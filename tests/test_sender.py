"""Testes do EnviadorSMTP. O smtplib é substituído por mocks: nenhum e-mail real é enviado."""

import smtplib
from email.message import EmailMessage
from unittest.mock import MagicMock, call, patch

import pytest

from email_sender.config import SMTPConfig
from email_sender.sender import EnviadorSMTP, EnvioError, eh_erro_temporario


@pytest.fixture
def config(variaveis_smtp):
    return SMTPConfig.from_mapping(variaveis_smtp)


@pytest.fixture
def mensagem():
    msg = EmailMessage()
    msg["From"] = "remetente@exemplo.com"
    msg["To"] = "ana@exemplo.com"
    msg["Subject"] = "Teste"
    msg.set_content("Olá")
    return msg


@pytest.fixture
def smtp_mock():
    with patch("email_sender.sender.smtplib.SMTP") as classe:
        yield classe


def test_conecta_com_starttls_e_autentica(config, mensagem, smtp_mock):
    servidor = smtp_mock.return_value

    with EnviadorSMTP(config) as enviador:
        enviador.enviar(mensagem)

    smtp_mock.assert_called_once_with("smtp.exemplo.com", 587, timeout=30.0)
    servidor.starttls.assert_called_once()
    servidor.login.assert_called_once_with("remetente@exemplo.com", "senha-de-teste")
    servidor.send_message.assert_called_once_with(mensagem)
    servidor.quit.assert_called_once()


def test_reutiliza_a_mesma_conexao_para_varios_envios(config, mensagem, smtp_mock):
    with EnviadorSMTP(config) as enviador:
        for _ in range(3):
            enviador.enviar(mensagem)

    assert smtp_mock.call_count == 1
    assert smtp_mock.return_value.send_message.call_count == 3


def test_modo_ssl_usa_smtp_ssl(variaveis_smtp, mensagem):
    config = SMTPConfig.from_mapping({**variaveis_smtp, "SMTP_SEGURANCA": "ssl", "SMTP_PORTA": ""})
    with patch("email_sender.sender.smtplib.SMTP_SSL") as smtp_ssl, EnviadorSMTP(config) as enviador:
        enviador.enviar(mensagem)

    assert smtp_ssl.call_args.args == ("smtp.exemplo.com", 465)
    smtp_ssl.return_value.starttls.assert_not_called()
    smtp_ssl.return_value.send_message.assert_called_once()


def test_sem_usuario_nao_faz_login(mensagem, smtp_mock):
    config = SMTPConfig.from_mapping({"SMTP_HOST": "relay.interno", "EMAIL_REMETENTE": "app@exemplo.com"})
    with EnviadorSMTP(config) as enviador:
        enviador.enviar(mensagem)
    smtp_mock.return_value.login.assert_not_called()


def test_reconecta_e_reenvia_apos_queda_de_conexao(config, mensagem, smtp_mock):
    primeira, segunda = MagicMock(), MagicMock()
    primeira.send_message.side_effect = smtplib.SMTPServerDisconnected("conexão perdida")
    smtp_mock.side_effect = [primeira, segunda]
    dormir = MagicMock()

    with EnviadorSMTP(config, backoff_segundos=2, dormir=dormir) as enviador:
        enviador.enviar(mensagem)

    segunda.send_message.assert_called_once_with(mensagem)
    primeira.close.assert_called_once()
    dormir.assert_called_once_with(2)


def test_backoff_exponencial_e_erro_apos_esgotar_tentativas(config, mensagem, smtp_mock):
    smtp_mock.return_value.send_message.side_effect = smtplib.SMTPResponseException(451, b"tente mais tarde")
    dormir = MagicMock()

    enviador = EnviadorSMTP(config, max_tentativas=3, backoff_segundos=1, dormir=dormir)
    with enviador, pytest.raises(EnvioError, match="3 tentativa"):
        enviador.enviar(mensagem)

    assert dormir.call_args_list == [call(1), call(2)]


def test_erro_permanente_nao_e_repetido(config, mensagem, smtp_mock):
    erro = smtplib.SMTPRecipientsRefused({"ana@exemplo.com": (550, b"caixa inexistente")})
    smtp_mock.return_value.send_message.side_effect = erro
    dormir = MagicMock()

    enviador = EnviadorSMTP(config, dormir=dormir)
    with enviador, pytest.raises(EnvioError, match="permanente") as info:
        enviador.enviar(mensagem)

    assert not info.value.fatal
    assert smtp_mock.return_value.send_message.call_count == 1
    dormir.assert_not_called()


def test_falha_de_autenticacao_e_fatal_e_sem_retentativa(config, smtp_mock):
    smtp_mock.return_value.login.side_effect = smtplib.SMTPAuthenticationError(535, b"credenciais invalidas")
    dormir = MagicMock()

    with pytest.raises(EnvioError, match="senha de app") as info:
        EnviadorSMTP(config, dormir=dormir).conectar_com_retentativa()

    assert info.value.fatal
    assert smtp_mock.call_count == 1
    smtp_mock.return_value.close.assert_called_once()  # socket liberado mesmo com erro
    dormir.assert_not_called()


def test_servidor_inacessivel_e_fatal_apos_tentativas(config, smtp_mock):
    smtp_mock.side_effect = ConnectionRefusedError("recusado")

    with pytest.raises(EnvioError) as info:
        EnviadorSMTP(config, max_tentativas=2, dormir=lambda _: None).conectar_com_retentativa()

    assert info.value.fatal
    assert smtp_mock.call_count == 2


@pytest.mark.parametrize(
    ("erro", "temporario"),
    [
        (smtplib.SMTPServerDisconnected(), True),
        (smtplib.SMTPResponseException(421, b"servico indisponivel"), True),
        (smtplib.SMTPResponseException(554, b"rejeitado"), False),
        (smtplib.SMTPSenderRefused(553, b"remetente invalido", "x@y.com"), False),
        (smtplib.SMTPRecipientsRefused({"a@b.com": (450, b"greylist")}), True),
        (smtplib.SMTPAuthenticationError(454, b"temp"), False),
        (smtplib.SMTPNotSupportedError(), False),
        (TimeoutError(), True),
        (ValueError(), False),
    ],
)
def test_classificacao_de_erros(erro, temporario):
    assert eh_erro_temporario(erro) is temporario
