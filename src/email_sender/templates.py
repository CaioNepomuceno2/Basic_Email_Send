"""Renderização de templates Jinja2 para o assunto e o corpo (texto e HTML) dos e-mails.

Convenção de arquivos na pasta de templates:
    <nome>.txt.j2   -> corpo em texto puro
    <nome>.html.j2  -> corpo em HTML (com escape automático de variáveis)
Pelo menos um dos dois deve existir. Sem a versão em texto, ela é gerada a partir do HTML.
"""

from __future__ import annotations

import html
import re
from collections.abc import Mapping
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path

from jinja2 import (
    Environment,
    FileSystemLoader,
    StrictUndefined,
    TemplateNotFound,
    TemplateSyntaxError,
    UndefinedError,
    select_autoescape,
)

SUFIXO_TEXTO = ".txt.j2"
SUFIXO_HTML = ".html.j2"


class TemplateError(Exception):
    """Template inexistente, com erro de sintaxe ou com variável não fornecida."""


@dataclass(frozen=True)
class Conteudo:
    texto: str
    html: str | None = None


class _ExtratorTexto(HTMLParser):
    """Converte HTML simples em texto legível (usado como alternativa ao HTML)."""

    QUEBRAS = {"br", "p", "div", "h1", "h2", "h3", "h4", "table"}  # quebra antes e depois
    LINHAS = {"tr", "li"}  # quebra apenas antes
    CELULAS = {"td", "th"}
    IGNORAR = {"style", "script", "head", "title"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.partes: list[str] = []
        self._ignorando = 0
        self._links: list[str | None] = []

    def handle_starttag(self, tag, attrs):
        if tag in self.IGNORAR:
            self._ignorando += 1
        elif tag == "li":
            self.partes.append("\n- ")
        elif tag in self.QUEBRAS or tag in self.LINHAS:
            self.partes.append("\n")
        if tag == "a":
            self._links.append(dict(attrs).get("href"))

    def handle_endtag(self, tag):
        if tag in self.IGNORAR:
            self._ignorando = max(0, self._ignorando - 1)
        elif tag == "a" and self._links:
            href = self._links.pop()
            if href and not href.startswith("mailto:"):
                self.partes.append(f" ({href})")
        elif tag in self.CELULAS:
            self.partes.append(" ")
        elif tag in self.QUEBRAS:
            self.partes.append("\n")

    def handle_data(self, data):
        if not self._ignorando:
            self.partes.append(data)


def html_para_texto(conteudo_html: str) -> str:
    extrator = _ExtratorTexto()
    extrator.feed(conteudo_html)
    texto = html.unescape("".join(extrator.partes))
    linhas = [re.sub(r"[ \t]+", " ", linha).strip() for linha in texto.splitlines()]
    return re.sub(r"\n{3,}", "\n\n", "\n".join(linhas)).strip() + "\n"


class RenderizadorTemplates:
    def __init__(self, pasta: Path | str) -> None:
        self.pasta = Path(pasta)
        self._env = Environment(
            loader=FileSystemLoader(self.pasta),
            undefined=StrictUndefined,  # variável ausente gera erro em vez de texto vazio
            autoescape=select_autoescape(enabled_extensions=("html", "html.j2"), default_for_string=False),
            keep_trailing_newline=True,
            trim_blocks=True,
            lstrip_blocks=True,
        )

    def disponiveis(self) -> list[str]:
        if not self.pasta.is_dir():
            return []
        nomes = {
            p.name.removesuffix(SUFIXO_TEXTO).removesuffix(SUFIXO_HTML)
            for p in self.pasta.iterdir()
            if p.name.endswith((SUFIXO_TEXTO, SUFIXO_HTML))
        }
        return sorted(nomes)

    def _renderizar_arquivo(self, arquivo: str, contexto: Mapping[str, object]) -> str | None:
        try:
            template = self._env.get_template(arquivo)
        except TemplateNotFound:
            return None
        except TemplateSyntaxError as erro:
            raise TemplateError(f"Erro de sintaxe em {arquivo}, linha {erro.lineno}: {erro.message}") from erro
        try:
            return template.render(**contexto)
        except UndefinedError as erro:
            raise TemplateError(f"{arquivo}: {erro.message}") from erro

    def renderizar(self, nome: str, contexto: Mapping[str, object]) -> Conteudo:
        texto = self._renderizar_arquivo(nome + SUFIXO_TEXTO, contexto)
        corpo_html = self._renderizar_arquivo(nome + SUFIXO_HTML, contexto)
        if texto is None and corpo_html is None:
            disponiveis = ", ".join(self.disponiveis()) or "nenhum"
            raise TemplateError(f"Template '{nome}' não encontrado em {self.pasta}. Disponíveis: {disponiveis}.")
        if texto is None:
            texto = html_para_texto(corpo_html)
        return Conteudo(texto=texto, html=corpo_html)

    def renderizar_texto(self, modelo: str, contexto: Mapping[str, object]) -> str:
        """Renderiza uma string (ex.: o assunto ou uma mensagem passada na linha de comando)."""
        try:
            return self._env.from_string(modelo).render(**contexto)
        except TemplateSyntaxError as erro:
            raise TemplateError(f"Erro de sintaxe no texto informado: {erro.message}") from erro
        except UndefinedError as erro:
            raise TemplateError(erro.message) from erro
