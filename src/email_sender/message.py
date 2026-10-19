"""Montagem da mensagem MIME: texto + HTML alternativos e anexos."""

from __future__ import annotations

import mimetypes
import re
from collections.abc import Sequence
from email.message import EmailMessage
from email.utils import formatdate, make_msgid, parseaddr
from pathlib import Path

from email_sender.templates import html_para_texto

# Validação pragmática (não cobre 100% da RFC 5322, mas evita erros comuns de digitação).
_REGEX_EMAIL = re.compile(
    r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@"
    r"[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$"
)

# Limite comum de provedores como Gmail e Outlook para o tamanho dos anexos.
LIMITE_ANEXOS_BYTES = 25 * 1024 * 1024


class MensagemError(ValueError):
    """Dados inválidos para montar a mensagem."""


def validar_email(endereco: str) -> str:
    endereco = (endereco or "").strip()
    if not _REGEX_EMAIL.match(endereco) or ".." in endereco:
        raise MensagemError(f"Endereço de e-mail inválido: {endereco!r}")
    return endereco


def limpar_cabecalho(valor: str) -> str:
    """Remove quebras de linha (evita injeção de cabeçalhos) e espaços repetidos."""
    return re.sub(r"\s+", " ", valor).strip()


def montar_mensagem(
    *,
    remetente: str,
    para: str,
    assunto: str,
    texto: str | None,
    html: str | None = None,
    anexos: Sequence[Path] = (),
    responder_para: str | None = None,
    descadastro: str | None = None,
) -> EmailMessage:
    """Cria um EmailMessage pronto para envio.

    Com texto e HTML, gera um `multipart/alternative`: o cliente de e-mail exibe
    o HTML e usa o texto como alternativa. Com anexos, tudo fica dentro de um
    `multipart/mixed`.
    """
    if not (texto and texto.strip()) and not (html and html.strip()):
        raise MensagemError("A mensagem precisa de um corpo em texto ou HTML.")
    assunto_limpo = limpar_cabecalho(assunto)
    if not assunto_limpo:
        raise MensagemError("O assunto não pode ficar vazio.")

    _, endereco_remetente = parseaddr(remetente)
    dominio = endereco_remetente.rpartition("@")[2] or None

    msg = EmailMessage()
    msg["From"] = remetente
    msg["To"] = validar_email(para)
    msg["Subject"] = assunto_limpo
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=dominio)
    if responder_para:
        msg["Reply-To"] = validar_email(responder_para)
    if descadastro:
        # Cabeçalho reconhecido por Gmail/Outlook para exibir a opção "cancelar inscrição".
        msg["List-Unsubscribe"] = f"<{limpar_cabecalho(descadastro)}>"

    msg.set_content(texto if texto and texto.strip() else html_para_texto(html or ""))
    if html:
        msg.add_alternative(html, subtype="html")

    total = 0
    for caminho in map(Path, anexos):
        if not caminho.is_file():
            raise MensagemError(f"Anexo não encontrado: {caminho}")
        dados = caminho.read_bytes()
        total += len(dados)
        if total > LIMITE_ANEXOS_BYTES:
            raise MensagemError(f"Os anexos excedem o limite de {LIMITE_ANEXOS_BYTES // (1024 * 1024)} MB.")
        tipo, codificacao = mimetypes.guess_type(caminho.name)
        if tipo is None or codificacao is not None:
            tipo = "application/octet-stream"
        principal, subtipo = tipo.split("/", 1)
        msg.add_attachment(dados, maintype=principal, subtype=subtipo, filename=caminho.name)

    return msg


def resumo_mensagem(msg: EmailMessage) -> str:
    """Linha descritiva para logs e para o modo de simulação."""
    anexos = [parte.get_filename() for parte in msg.iter_attachments()]
    tamanho_kb = len(msg.as_bytes()) / 1024
    detalhe_anexos = f" | anexos: {', '.join(anexos)}" if anexos else ""
    return f"para={msg['To']} | assunto={msg['Subject']!r} | {tamanho_kb:.1f} KB{detalhe_anexos}"
