"""Leitura da lista de destinatários a partir de um CSV.

- A coluna `email` é obrigatória; as demais viram variáveis nos templates.
- Os nomes das colunas são normalizados (minúsculas, espaços viram "_").
- Aceita CSV separado por vírgula, ponto e vírgula (padrão do Excel em pt-BR) ou tab.
- Coluna opcional `anexos`: caminhos separados por "|", relativos à pasta do CSV.
- Linhas com e-mail vazio, inválido ou repetido são ignoradas e reportadas.
"""

from __future__ import annotations

import csv
import io
from dataclasses import dataclass, field
from pathlib import Path

from email_sender.message import MensagemError, validar_email

SEPARADOR_ANEXOS = "|"


@dataclass(frozen=True)
class Destinatario:
    email: str
    dados: dict[str, str] = field(default_factory=dict)
    anexos: tuple[Path, ...] = ()


@dataclass
class ListaDestinatarios:
    destinatarios: list[Destinatario] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)


def _normalizar_coluna(nome: str) -> str:
    return "_".join(nome.strip().lower().split())


def _detectar_dialeto(cabecalho: str) -> type[csv.Dialect] | csv.Dialect:
    try:
        return csv.Sniffer().sniff(cabecalho, delimiters=",;\t")
    except csv.Error:
        return csv.excel  # arquivo com uma única coluna, por exemplo


def carregar_destinatarios(
    caminho: Path | str, *, coluna_email: str = "email", coluna_anexos: str = "anexos"
) -> ListaDestinatarios:
    caminho = Path(caminho)
    if not caminho.is_file():
        raise FileNotFoundError(f"CSV de destinatários não encontrado: {caminho}")

    # utf-8-sig remove o BOM que o Excel costuma adicionar ao salvar CSV em UTF-8
    conteudo = caminho.read_text(encoding="utf-8-sig")
    if not conteudo.strip():
        raise ValueError(f"O arquivo {caminho} está vazio.")

    leitor = csv.DictReader(io.StringIO(conteudo), dialect=_detectar_dialeto(conteudo.splitlines()[0]))
    colunas = [_normalizar_coluna(c) for c in (leitor.fieldnames or [])]
    if coluna_email not in colunas:
        raise ValueError(f"O CSV precisa de uma coluna '{coluna_email}'. Colunas encontradas: {colunas}")
    leitor.fieldnames = colunas

    resultado = ListaDestinatarios()
    vistos: set[str] = set()
    for numero_linha, linha in enumerate(leitor, start=2):  # linha 1 é o cabeçalho
        dados = {chave: (valor or "").strip() for chave, valor in linha.items() if chave}
        bruto = dados.pop(coluna_email, "")
        if not bruto:
            resultado.avisos.append(f"Linha {numero_linha}: e-mail vazio, linha ignorada.")
            continue
        try:
            email = validar_email(bruto)
        except MensagemError as erro:
            resultado.avisos.append(f"Linha {numero_linha}: {erro}")
            continue
        if email.lower() in vistos:
            resultado.avisos.append(f"Linha {numero_linha}: {email} repetido, linha ignorada.")
            continue
        vistos.add(email.lower())

        texto_anexos = dados.pop(coluna_anexos, "")
        anexos = tuple(
            (caminho.parent / item.strip()).resolve() for item in texto_anexos.split(SEPARADOR_ANEXOS) if item.strip()
        )
        resultado.destinatarios.append(Destinatario(email=email, dados=dados, anexos=anexos))

    return resultado
