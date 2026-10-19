import pytest

from email_sender.config import ConfigError, EnvioConfig, SMTPConfig, carregar_variaveis


def test_carrega_config_completa(variaveis_smtp):
    config = SMTPConfig.from_mapping(variaveis_smtp)

    assert config.host == "smtp.exemplo.com"
    assert config.porta == 587
    assert config.seguranca == "starttls"
    assert config.remetente == "remetente@exemplo.com"  # usa SMTP_USUARIO quando EMAIL_REMETENTE falta
    assert config.remetente_formatado == "Equipe Exemplo <remetente@exemplo.com>"


def test_senha_nao_aparece_no_repr(variaveis_smtp):
    assert "senha-de-teste" not in repr(SMTPConfig.from_mapping(variaveis_smtp))


@pytest.mark.parametrize(("seguranca", "porta"), [("starttls", 587), ("ssl", 465), ("nenhuma", 25)])
def test_porta_padrao_por_tipo_de_seguranca(seguranca, porta):
    config = SMTPConfig.from_mapping(
        {"SMTP_HOST": "localhost", "EMAIL_REMETENTE": "a@b.com", "SMTP_SEGURANCA": seguranca}
    )
    assert config.porta == porta


def test_config_incompleta_lista_o_que_falta():
    with pytest.raises(ConfigError, match="SMTP_HOST.*SMTP_SENHA"):
        SMTPConfig.from_mapping({"SMTP_USUARIO": "a@b.com"})


def test_simulacao_dispensa_credenciais():
    config = SMTPConfig.from_mapping({}, exigir_credenciais=False)
    assert config.remetente == "remetente@example.com"


@pytest.mark.parametrize(
    "variaveis",
    [
        {"SMTP_SEGURANCA": "tls-magico"},
        {"SMTP_PORTA": "abc"},
        {"SMTP_TIMEOUT": "-1"},
        {"EMAIL_REMETENTE": "sem-arroba"},
    ],
)
def test_valores_invalidos(variaveis_smtp, variaveis):
    with pytest.raises(ConfigError):
        SMTPConfig.from_mapping({**variaveis_smtp, **variaveis})


def test_bloqueia_senha_sem_tls_em_servidor_remoto(variaveis_smtp):
    with pytest.raises(ConfigError, match="sem TLS"):
        SMTPConfig.from_mapping({**variaveis_smtp, "SMTP_SEGURANCA": "nenhuma"})
    # em servidor local (ex.: ambiente de testes) é permitido
    SMTPConfig.from_mapping({**variaveis_smtp, "SMTP_SEGURANCA": "nenhuma", "SMTP_HOST": "localhost"})


def test_envio_config():
    assert EnvioConfig.from_mapping({}) == EnvioConfig(1.0, 3, 2.0)
    config = EnvioConfig.from_mapping({"ENVIO_INTERVALO_SEGUNDOS": "0.5", "ENVIO_MAX_TENTATIVAS": "5"})
    assert (config.intervalo_segundos, config.max_tentativas) == (0.5, 5)
    with pytest.raises(ConfigError):
        EnvioConfig.from_mapping({"ENVIO_MAX_TENTATIVAS": "0"})


def test_ambiente_tem_prioridade_sobre_arquivo_env(tmp_path, monkeypatch):
    arquivo = tmp_path / ".env"
    arquivo.write_text("SMTP_HOST=do.arquivo\nSMTP_PORTA=2525\n", encoding="utf-8")
    monkeypatch.setenv("SMTP_HOST", "do.ambiente")

    variaveis = carregar_variaveis()  # lê o .env do diretório atual (tmp_path)

    assert variaveis["SMTP_HOST"] == "do.ambiente"
    assert variaveis["SMTP_PORTA"] == "2525"


def test_env_file_explicito_inexistente(tmp_path):
    with pytest.raises(ConfigError):
        carregar_variaveis(tmp_path / "nao_existe.env")
