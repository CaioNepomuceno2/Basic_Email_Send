from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
PASTA_TEMPLATES = RAIZ / "templates"
CSV_EXEMPLO = RAIZ / "exemplos" / "destinatarios.csv"

PREFIXOS_CONFIG = ("SMTP_", "EMAIL_", "ENVIO_")


@pytest.fixture(autouse=True)
def ambiente_isolado(monkeypatch, tmp_path):
    """Garante que nenhum teste leia credenciais reais do ambiente ou de um .env local."""
    import os

    for nome in list(os.environ):
        if nome.startswith(PREFIXOS_CONFIG):
            monkeypatch.delenv(nome)
    monkeypatch.chdir(tmp_path)


@pytest.fixture
def variaveis_smtp():
    return {
        "SMTP_HOST": "smtp.exemplo.com",
        "SMTP_PORTA": "587",
        "SMTP_USUARIO": "remetente@exemplo.com",
        "SMTP_SENHA": "senha-de-teste",
        "EMAIL_NOME_REMETENTE": "Equipe Exemplo",
    }


@pytest.fixture
def pasta_templates():
    return PASTA_TEMPLATES


@pytest.fixture
def csv_exemplo():
    return CSV_EXEMPLO
