import pytest

from email_sender import message
from email_sender.message import MensagemError, montar_mensagem, resumo_mensagem, validar_email

BASE = {"remetente": "Equipe <equipe@exemplo.com>", "para": "ana@exemplo.com", "assunto": "Olá"}


@pytest.mark.parametrize("endereco", ["ana@exemplo.com", "a.b+tag@sub.exemplo.com.br", "  ana@exemplo.com  "])
def test_emails_validos(endereco):
    assert validar_email(endereco) == endereco.strip()


@pytest.mark.parametrize(
    "endereco", ["", "ana", "ana@", "@exemplo.com", "ana@exemplo", "a..b@exemplo.com", "a b@x.com"]
)
def test_emails_invalidos(endereco):
    with pytest.raises(MensagemError):
        validar_email(endereco)


def test_texto_e_html_viram_multipart_alternative():
    msg = montar_mensagem(**BASE, texto="Oi", html="<p>Oi</p>")

    assert msg.get_content_type() == "multipart/alternative"
    assert msg.get_body(("plain",)).get_content().strip() == "Oi"
    assert msg.get_body(("html",)).get_content().strip() == "<p>Oi</p>"
    assert msg["Message-ID"].endswith("@exemplo.com>")
    assert msg["Date"]


def test_somente_html_gera_alternativa_em_texto():
    msg = montar_mensagem(**BASE, texto=None, html="<h1>Título</h1><p>Corpo</p>")
    assert "Título" in msg.get_body(("plain",)).get_content()


def test_anexos(tmp_path):
    pdf = tmp_path / "relatorio.pdf"
    pdf.write_bytes(b"%PDF-1.4 conteudo")
    csv = tmp_path / "dados.csv"
    csv.write_text("a,b\n1,2\n", encoding="utf-8")

    msg = montar_mensagem(**BASE, texto="Segue anexo.", anexos=[pdf, csv])

    assert msg.get_content_type() == "multipart/mixed"
    anexos = {parte.get_filename(): parte.get_content_type() for parte in msg.iter_attachments()}
    assert anexos == {"relatorio.pdf": "application/pdf", "dados.csv": "text/csv"}
    assert "anexos: relatorio.pdf, dados.csv" in resumo_mensagem(msg)


def test_anexo_inexistente(tmp_path):
    with pytest.raises(MensagemError, match="não encontrado"):
        montar_mensagem(**BASE, texto="x", anexos=[tmp_path / "sumiu.pdf"])


def test_limite_de_tamanho_dos_anexos(tmp_path, monkeypatch):
    monkeypatch.setattr(message, "LIMITE_ANEXOS_BYTES", 10)
    grande = tmp_path / "grande.bin"
    grande.write_bytes(b"0" * 11)
    with pytest.raises(MensagemError, match="limite"):
        montar_mensagem(**BASE, texto="x", anexos=[grande])


def test_assunto_com_quebra_de_linha_nao_injeta_cabecalho():
    msg = montar_mensagem(**{**BASE, "assunto": "Oferta\r\nBcc: vitima@exemplo.com"}, texto="x")
    assert msg["Subject"] == "Oferta Bcc: vitima@exemplo.com"
    assert msg["Bcc"] is None


def test_cabecalhos_opcionais():
    msg = montar_mensagem(
        **BASE, texto="x", responder_para="suporte@exemplo.com", descadastro="https://exemplo.com/sair"
    )
    assert msg["Reply-To"] == "suporte@exemplo.com"
    assert msg["List-Unsubscribe"] == "<https://exemplo.com/sair>"


@pytest.mark.parametrize(("texto", "assunto"), [("", "Olá"), ("   ", "Olá"), ("corpo", "  ")])
def test_corpo_ou_assunto_vazio(texto, assunto):
    with pytest.raises(MensagemError):
        montar_mensagem(**{**BASE, "assunto": assunto}, texto=texto)
