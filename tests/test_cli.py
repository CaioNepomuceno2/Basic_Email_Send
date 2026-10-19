"""Testes de ponta a ponta da CLI, com smtplib.SMTP substituído por um mock."""

import csv
from unittest.mock import patch

import pytest

from email_sender.cli import SAIDA_ERRO_USO, SAIDA_FALHA_ENVIO, SAIDA_OK, main


@pytest.fixture
def env_file(tmp_path, variaveis_smtp):
    caminho = tmp_path / "teste.env"
    linhas = [f"{chave}={valor}" for chave, valor in variaveis_smtp.items()]
    caminho.write_text("\n".join(linhas + ["ENVIO_INTERVALO_SEGUNDOS=0", "ENVIO_BACKOFF_SEGUNDOS=0"]), encoding="utf-8")
    return caminho


@pytest.fixture
def smtp_mock():
    with patch("email_sender.sender.smtplib.SMTP") as classe:
        yield classe


def args_lote(csv_exemplo, pasta_templates, *extras):
    return [
        "lote",
        "--csv",
        str(csv_exemplo),
        "--template",
        "lembrete_evento",
        "--assunto",
        "Lembrete: {{ evento }}",
        "--pasta-templates",
        str(pasta_templates),
        *extras,
    ]


def test_dry_run_nao_abre_conexao_nem_exige_credenciais(csv_exemplo, pasta_templates, smtp_mock, caplog):
    codigo = main(args_lote(csv_exemplo, pasta_templates, "--dry-run"))

    assert codigo == SAIDA_OK
    smtp_mock.assert_not_called()


def test_lote_envia_uma_mensagem_personalizada_por_destinatario(
    csv_exemplo, pasta_templates, env_file, smtp_mock, tmp_path
):
    relatorio = tmp_path / "relatorio.csv"

    codigo = main(args_lote(csv_exemplo, pasta_templates, "--env-file", str(env_file), "--relatorio", str(relatorio)))

    assert codigo == SAIDA_OK
    enviados = [c.args[0] for c in smtp_mock.return_value.send_message.call_args_list]
    assert [m["To"] for m in enviados] == [
        "ana.souza@example.com",
        "bruno.lima@example.com",
        "carla.mendes@example.com",
    ]
    assert enviados[0]["From"] == "Equipe Exemplo <remetente@exemplo.com>"
    assert [p.get_filename() for p in enviados[0].iter_attachments()] == ["programacao.txt"]
    assert list(enviados[1].iter_attachments()) == []
    with relatorio.open(encoding="utf-8") as arquivo:
        assert {linha["status"] for linha in csv.DictReader(arquivo)} == {"enviado"}


def test_limite_de_destinatarios(csv_exemplo, pasta_templates, env_file, smtp_mock):
    main(args_lote(csv_exemplo, pasta_templates, "--env-file", str(env_file), "--limite", "2"))
    assert smtp_mock.return_value.send_message.call_count == 2


def test_enviar_mensagem_simples_com_variaveis_e_anexo(env_file, smtp_mock, tmp_path):
    anexo = tmp_path / "nota.txt"
    anexo.write_text("conteúdo", encoding="utf-8")

    codigo = main(
        [
            "enviar",
            "--env-file",
            str(env_file),
            "--para",
            "ana@exemplo.com",
            "--para",
            "leo@exemplo.com",
            "--assunto",
            "Olá, {{ nome }}",
            "--mensagem",
            "Segue o arquivo, {{ nome }}.",
            "--var",
            "nome=Equipe",
            "--anexo",
            str(anexo),
        ]
    )

    assert codigo == SAIDA_OK
    mensagens = [c.args[0] for c in smtp_mock.return_value.send_message.call_args_list]
    assert len(mensagens) == 2
    assert mensagens[0]["Subject"] == "Olá, Equipe"
    assert mensagens[0].get_body(("plain",)).get_content().strip() == "Segue o arquivo, Equipe."


def test_falha_no_envio_retorna_codigo_1(env_file, smtp_mock):
    smtp_mock.return_value.send_message.side_effect = OSError("rede caiu")

    codigo = main(
        ["enviar", "--env-file", str(env_file), "--para", "a@exemplo.com", "--assunto", "x", "--mensagem", "y"]
    )

    assert codigo == SAIDA_FALHA_ENVIO


@pytest.mark.parametrize(
    "argumentos",
    [
        ["enviar", "--para", "a@exemplo.com", "--assunto", "x", "--mensagem", "y"],  # sem configuração SMTP
        ["enviar", "--para", "invalido", "--assunto", "x", "--mensagem", "y", "--dry-run"],
        ["enviar", "--para", "a@exemplo.com", "--assunto", "x", "--template", "inexistente", "--dry-run"],
        ["enviar", "--para", "a@exemplo.com", "--assunto", "x", "--mensagem", "y", "--anexo", "sumiu.pdf", "--dry-run"],
    ],
)
def test_erros_de_uso_retornam_codigo_2(argumentos, smtp_mock):
    assert main(argumentos) == SAIDA_ERRO_USO
    smtp_mock.assert_not_called()


def test_testar_conexao(env_file, smtp_mock):
    assert main(["testar-conexao", "--env-file", str(env_file)]) == SAIDA_OK
    smtp_mock.return_value.login.assert_called_once()
    smtp_mock.return_value.send_message.assert_not_called()


def test_listar_templates(pasta_templates, capsys):
    assert main(["templates", "--pasta-templates", str(pasta_templates)]) == SAIDA_OK
    assert capsys.readouterr().out.split() == ["boas_vindas", "lembrete_evento"]


def test_log_em_arquivo_nao_contem_senha(csv_exemplo, pasta_templates, env_file, smtp_mock, tmp_path):
    log = tmp_path / "envio.log"
    main(args_lote(csv_exemplo, pasta_templates, "--env-file", str(env_file), "--log-arquivo", str(log), "-v"))

    conteudo = log.read_text(encoding="utf-8")
    assert "Enviado" in conteudo
    assert "senha-de-teste" not in conteudo
