"""Orquestra o envio: personaliza cada mensagem, respeita o ritmo e gera o relatório."""

from __future__ import annotations

import csv
import logging
import re
from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field
from email.message import EmailMessage
from pathlib import Path
from typing import Literal

from email_sender.message import MensagemError, montar_mensagem, resumo_mensagem
from email_sender.rate_limit import LimitadorTaxa
from email_sender.recipients import Destinatario
from email_sender.sender import EnviadorSMTP, EnvioError
from email_sender.templates import Conteudo, RenderizadorTemplates, TemplateError

logger = logging.getLogger(__name__)

Status = Literal["enviado", "simulado", "falha", "nao_enviado"]


@dataclass(frozen=True)
class ResultadoEnvio:
    email: str
    status: Status
    detalhe: str = ""


@dataclass
class Relatorio:
    resultados: list[ResultadoEnvio] = field(default_factory=list)

    def contar(self, status: Status) -> int:
        return sum(1 for r in self.resultados if r.status == status)

    @property
    def houve_falha(self) -> bool:
        return any(r.status in ("falha", "nao_enviado") for r in self.resultados)

    def resumo(self) -> str:
        partes = [f"{self.contar(s)} {s.replace('_', ' ')}" for s in ("enviado", "simulado", "falha", "nao_enviado")]
        return f"Total: {len(self.resultados)} | " + " | ".join(p for p in partes if not p.startswith("0 "))

    def salvar_csv(self, caminho: Path) -> None:
        caminho.parent.mkdir(parents=True, exist_ok=True)
        with caminho.open("w", newline="", encoding="utf-8") as arquivo:
            escritor = csv.writer(arquivo)
            escritor.writerow(["email", "status", "detalhe"])
            escritor.writerows((r.email, r.status, r.detalhe) for r in self.resultados)


@dataclass
class Campanha:
    """Define o que será enviado. O assunto e o corpo aceitam variáveis Jinja2 ({{ nome }})."""

    remetente: str
    assunto: str
    renderizador: RenderizadorTemplates
    template: str | None = None
    mensagem: str | None = None
    anexos: tuple[Path, ...] = ()
    variaveis: Mapping[str, str] = field(default_factory=dict)
    responder_para: str | None = None
    descadastro: str | None = None

    def __post_init__(self) -> None:
        if bool(self.template) == bool(self.mensagem):
            raise ValueError("Informe exatamente um entre 'template' e 'mensagem'.")

    def preparar(self, destinatario: Destinatario) -> EmailMessage:
        # Prioridade: dados do CSV > variáveis da linha de comando.
        contexto = {**self.variaveis, **destinatario.dados, "email": destinatario.email}
        assunto = self.renderizador.renderizar_texto(self.assunto, contexto)
        if self.template:
            conteudo = self.renderizador.renderizar(self.template, contexto)
        else:
            conteudo = Conteudo(texto=self.renderizador.renderizar_texto(self.mensagem or "", contexto))
        return montar_mensagem(
            remetente=self.remetente,
            para=destinatario.email,
            assunto=assunto,
            texto=conteudo.texto,
            html=conteudo.html,
            anexos=(*self.anexos, *destinatario.anexos),
            responder_para=self.responder_para,
            descadastro=self.descadastro,
        )


def _nome_arquivo_eml(indice: int, email: str) -> str:
    return f"{indice:04d}_{re.sub(r'[^A-Za-z0-9._-]', '_', email)}.eml"


def executar(
    campanha: Campanha,
    destinatarios: Iterable[Destinatario],
    *,
    enviador: EnviadorSMTP | None,
    limitador: LimitadorTaxa | None = None,
    pasta_eml: Path | None = None,
) -> Relatorio:
    """Envia (ou simula, se `enviador` for None) a campanha para cada destinatário.

    Erros de um destinatário não interrompem os demais, exceto falhas fatais
    (ex.: autenticação recusada), que tornariam todos os envios seguintes inúteis.
    """
    simulacao = enviador is None
    relatorio = Relatorio()
    pendentes = list(destinatarios)

    for indice, destinatario in enumerate(pendentes, start=1):
        prefixo = f"[{indice}/{len(pendentes)}]"
        try:
            mensagem = campanha.preparar(destinatario)
        except (TemplateError, MensagemError) as erro:
            logger.error("%s %s: mensagem não montada: %s", prefixo, destinatario.email, erro)
            relatorio.resultados.append(ResultadoEnvio(destinatario.email, "falha", str(erro)))
            continue

        if pasta_eml is not None:
            pasta_eml.mkdir(parents=True, exist_ok=True)
            (pasta_eml / _nome_arquivo_eml(indice, destinatario.email)).write_bytes(mensagem.as_bytes())

        if simulacao:
            logger.info("%s [SIMULAÇÃO] %s", prefixo, resumo_mensagem(mensagem))
            logger.debug("Corpo em texto:\n%s", mensagem.get_body(("plain",)).get_content())
            relatorio.resultados.append(ResultadoEnvio(destinatario.email, "simulado"))
            continue

        if limitador is not None:
            limitador.aguardar()
        try:
            enviador.enviar(mensagem)
        except EnvioError as erro:
            logger.error("%s %s: %s", prefixo, destinatario.email, erro)
            relatorio.resultados.append(ResultadoEnvio(destinatario.email, "falha", str(erro)))
            if erro.fatal:
                restantes = pendentes[indice:]
                logger.error("Envio interrompido. %d destinatário(s) não processado(s).", len(restantes))
                relatorio.resultados.extend(
                    ResultadoEnvio(d.email, "nao_enviado", "interrompido após erro fatal") for d in restantes
                )
                break
        else:
            logger.info("%s Enviado: %s", prefixo, resumo_mensagem(mensagem))
            relatorio.resultados.append(ResultadoEnvio(destinatario.email, "enviado"))

    return relatorio
