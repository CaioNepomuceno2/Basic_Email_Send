import pytest

from email_sender.templates import RenderizadorTemplates, TemplateError, html_para_texto


@pytest.fixture
def renderizador(pasta_templates):
    return RenderizadorTemplates(pasta_templates)


def test_lista_templates_do_projeto(renderizador):
    assert renderizador.disponiveis() == ["boas_vindas", "lembrete_evento"]


def test_renderiza_texto_e_html(renderizador):
    conteudo = renderizador.renderizar("boas_vindas", {"nome": "Ana", "email": "ana@exemplo.com", "empresa": "ACME"})

    assert "Olá, Ana!" in conteudo.texto
    assert "à ACME" in conteudo.texto
    assert "<strong>ACME</strong>" in conteudo.html
    assert "Equipe de Atendimento" in conteudo.texto  # valor padrão via filtro default


def test_html_escapa_variaveis_mas_texto_nao(tmp_path):
    (tmp_path / "t.txt.j2").write_text("{{ nome }}", encoding="utf-8")
    (tmp_path / "t.html.j2").write_text("<p>{{ nome }}</p>", encoding="utf-8")

    conteudo = RenderizadorTemplates(tmp_path).renderizar("t", {"nome": "<script>x</script>"})

    assert conteudo.texto == "<script>x</script>"
    assert conteudo.html == "<p>&lt;script&gt;x&lt;/script&gt;</p>"


def test_gera_texto_a_partir_do_html_quando_nao_ha_txt(renderizador):
    contexto = {"nome": "Ana", "evento": "Meetup", "data": "01/01", "horario": "19h", "link": "https://x.example"}
    conteudo = renderizador.renderizar("lembrete_evento", contexto)

    assert conteudo.html is not None
    assert "Olá, Ana!" in conteudo.texto
    assert "Data 01/01" in conteudo.texto
    assert "(https://x.example)" in conteudo.texto
    assert "<" not in conteudo.texto


def test_variavel_ausente_gera_erro_claro(renderizador):
    with pytest.raises(TemplateError, match="nome"):
        renderizador.renderizar("boas_vindas", {"email": "a@b.com"})


def test_template_inexistente_lista_disponiveis(renderizador):
    with pytest.raises(TemplateError, match="boas_vindas"):
        renderizador.renderizar("nao_existe", {})


def test_renderizar_texto(renderizador):
    assert renderizador.renderizar_texto("Olá, {{ nome }}", {"nome": "Leo"}) == "Olá, Leo"
    with pytest.raises(TemplateError):
        renderizador.renderizar_texto("{{ nome ", {})


def test_html_para_texto_ignora_style_e_formata_listas():
    texto = html_para_texto("<style>p{color:red}</style><p>Itens:</p><ul><li>Um</li><li>Dois</li></ul>")
    assert texto == "Itens:\n\n- Um\n- Dois\n"
