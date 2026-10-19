from unittest.mock import MagicMock

import pytest

from email_sender.campanha import Campanha, executar
from email_sender.recipients import Destinatario
from email_sender.sender import EnvioError
from email_sender.templates import RenderizadorTemplates


@pytest.fixture
def campanha(pasta_templates):
    return Campanha(
        remetente="Equipe <equipe@exemplo.com>",
        assunto="Bem-vindo(a), {{ nome }}",
        renderizador=RenderizadorTemplates(pasta_templates),
        template="boas_vindas",
        variaveis={"empresa": "ACME", "nome": "Cliente"},
    )


def destinatarios(*nomes):
    return [Destinatario(email=f"{n.lower()}@exemplo.com", dados={"nome": n}) for n in nomes]


def test_personaliza_assunto_e_corpo_com_prioridade_do_csv(campanha):
    msg = campanha.preparar(Destinatario("ana@exemplo.com", {"nome": "Ana"}))

    assert msg["Subject"] == "Bem-vindo(a), Ana"
    assert "Olá, Ana!" in msg.get_body(("plain",)).get_content()
    assert "ACME" in msg.get_body(("html",)).get_content()


def test_exige_template_ou_mensagem(pasta_templates):
    with pytest.raises(ValueError):
        Campanha(remetente="a@b.com", assunto="x", renderizador=RenderizadorTemplates(pasta_templates))


def test_simulacao_nao_envia_e_salva_eml(campanha, tmp_path):
    relatorio = executar(campanha, destinatarios("Ana", "Leo"), enviador=None, pasta_eml=tmp_path / "eml")

    assert relatorio.contar("simulado") == 2
    arquivos = sorted(p.name for p in (tmp_path / "eml").iterdir())
    assert arquivos == ["0001_ana_exemplo.com.eml", "0002_leo_exemplo.com.eml"]


def test_envio_usa_limitador_e_continua_apos_falha_comum(campanha):
    enviador = MagicMock()
    enviador.enviar.side_effect = [None, EnvioError("550 caixa inexistente"), None]
    limitador = MagicMock()

    relatorio = executar(campanha, destinatarios("Ana", "Leo", "Bia"), enviador=enviador, limitador=limitador)

    assert [r.status for r in relatorio.resultados] == ["enviado", "falha", "enviado"]
    assert limitador.aguardar.call_count == 3
    assert relatorio.houve_falha


def test_falha_fatal_interrompe_o_lote(campanha):
    enviador = MagicMock()
    enviador.enviar.side_effect = EnvioError("autenticação", fatal=True)

    relatorio = executar(campanha, destinatarios("Ana", "Leo", "Bia"), enviador=enviador)

    assert [r.status for r in relatorio.resultados] == ["falha", "nao_enviado", "nao_enviado"]
    assert enviador.enviar.call_count == 1


def test_erro_de_template_vira_falha_do_destinatario(pasta_templates):
    campanha = Campanha(
        remetente="a@exemplo.com",
        assunto="Oi {{ apelido }}",
        renderizador=RenderizadorTemplates(pasta_templates),
        mensagem="corpo",
    )
    relatorio = executar(campanha, destinatarios("Ana"), enviador=MagicMock())

    assert relatorio.resultados[0].status == "falha"
    assert "apelido" in relatorio.resultados[0].detalhe


def test_relatorio_csv(campanha, tmp_path):
    relatorio = executar(campanha, destinatarios("Ana"), enviador=None)
    caminho = tmp_path / "saida" / "relatorio.csv"
    relatorio.salvar_csv(caminho)

    assert caminho.read_text(encoding="utf-8").splitlines() == ["email,status,detalhe", "ana@exemplo.com,simulado,"]
    assert relatorio.resumo() == "Total: 1 | 1 simulado"
