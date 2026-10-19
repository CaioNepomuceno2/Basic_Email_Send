"""Configuração do envio, lida de variáveis de ambiente e/ou de um arquivo `.env`."""

from __future__ import annotations

import os
from collections.abc import Mapping
from dataclasses import dataclass, field
from email.utils import formataddr
from pathlib import Path

from dotenv import dotenv_values

from email_sender.message import MensagemError, validar_email

SEGURANCAS_VALIDAS = ("starttls", "ssl", "nenhuma")
PORTAS_PADRAO = {"starttls": 587, "ssl": 465, "nenhuma": 25}
HOSTS_LOCAIS = {"localhost", "127.0.0.1", "::1"}
REMETENTE_SIMULACAO = "remetente@example.com"


class ConfigError(ValueError):
    """Configuração ausente ou inválida."""


def carregar_variaveis(env_file: Path | None = None) -> dict[str, str]:
    """Combina o arquivo .env com as variáveis do ambiente.

    Variáveis já definidas no ambiente têm prioridade sobre o arquivo, o que
    permite sobrescrever valores em CI ou em execuções pontuais.
    Se `env_file` não for informado, usa `.env` do diretório atual, se existir.
    """
    valores: dict[str, str] = {}
    if env_file is not None and not env_file.is_file():
        raise ConfigError(f"Arquivo de configuração não encontrado: {env_file}")
    arquivo = env_file or Path(".env")
    if arquivo.is_file():
        valores.update({k: v for k, v in dotenv_values(arquivo).items() if v is not None})
    valores.update(os.environ)
    return valores


def _ler_numero(variaveis: Mapping[str, str], nome: str, padrao: float, tipo: type) -> float:
    bruto = variaveis.get(nome, "").strip()
    if not bruto:
        return padrao
    try:
        valor = tipo(bruto)
    except ValueError:
        raise ConfigError(f"{nome} deve ser numérico (recebido: {bruto!r}).") from None
    if valor < 0:
        raise ConfigError(f"{nome} não pode ser negativo.")
    return valor


@dataclass(frozen=True)
class SMTPConfig:
    """Dados de conexão com o servidor SMTP."""

    host: str
    porta: int
    usuario: str | None
    senha: str | None = field(repr=False)  # nunca aparece em logs ou no repr
    remetente: str
    nome_remetente: str | None = None
    seguranca: str = "starttls"
    timeout: float = 30.0

    @property
    def remetente_formatado(self) -> str:
        """Ex.: 'Caio <caio@exemplo.com>'."""
        return formataddr((self.nome_remetente or "", self.remetente))

    @classmethod
    def from_mapping(cls, variaveis: Mapping[str, str], *, exigir_credenciais: bool = True) -> SMTPConfig:
        seguranca = variaveis.get("SMTP_SEGURANCA", "starttls").strip().lower() or "starttls"
        if seguranca not in SEGURANCAS_VALIDAS:
            raise ConfigError(f"SMTP_SEGURANCA deve ser uma de {SEGURANCAS_VALIDAS} (recebido: {seguranca!r}).")

        host = variaveis.get("SMTP_HOST", "").strip()
        usuario = variaveis.get("SMTP_USUARIO", "").strip() or None
        senha = variaveis.get("SMTP_SENHA") or None
        remetente = (variaveis.get("EMAIL_REMETENTE", "").strip() or usuario or "").strip()
        porta = int(_ler_numero(variaveis, "SMTP_PORTA", PORTAS_PADRAO[seguranca], int))
        timeout = float(_ler_numero(variaveis, "SMTP_TIMEOUT", 30.0, float))

        if exigir_credenciais:
            faltando = []
            if not host:
                faltando.append("SMTP_HOST")
            if not remetente:
                faltando.append("EMAIL_REMETENTE (ou SMTP_USUARIO)")
            if usuario and not senha:
                faltando.append("SMTP_SENHA")
            if faltando:
                raise ConfigError("Configuração incompleta. Defina: " + ", ".join(faltando))
            try:
                validar_email(remetente)
            except MensagemError as erro:
                raise ConfigError(f"Remetente inválido: {erro}") from None
            if seguranca == "nenhuma" and senha and host not in HOSTS_LOCAIS:
                raise ConfigError("Por segurança, autenticação sem TLS só é permitida em servidores locais.")

        return cls(
            host=host or "localhost",
            porta=porta,
            usuario=usuario,
            senha=senha,
            remetente=remetente or REMETENTE_SIMULACAO,
            nome_remetente=variaveis.get("EMAIL_NOME_REMETENTE", "").strip() or None,
            seguranca=seguranca,
            timeout=timeout,
        )


@dataclass(frozen=True)
class EnvioConfig:
    """Parâmetros de ritmo e resiliência do envio."""

    intervalo_segundos: float = 1.0
    max_tentativas: int = 3
    backoff_segundos: float = 2.0

    @classmethod
    def from_mapping(cls, variaveis: Mapping[str, str]) -> EnvioConfig:
        tentativas = int(_ler_numero(variaveis, "ENVIO_MAX_TENTATIVAS", 3, int))
        if tentativas < 1:
            raise ConfigError("ENVIO_MAX_TENTATIVAS deve ser pelo menos 1.")
        return cls(
            intervalo_segundos=float(_ler_numero(variaveis, "ENVIO_INTERVALO_SEGUNDOS", 1.0, float)),
            max_tentativas=tentativas,
            backoff_segundos=float(_ler_numero(variaveis, "ENVIO_BACKOFF_SEGUNDOS", 2.0, float)),
        )
