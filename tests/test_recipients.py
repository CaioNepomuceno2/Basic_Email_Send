import pytest

from email_sender.recipients import carregar_destinatarios


def escrever(tmp_path, conteudo, nome="lista.csv", encoding="utf-8"):
    caminho = tmp_path / nome
    caminho.write_text(conteudo, encoding=encoding)
    return caminho


def test_csv_de_exemplo_do_projeto(csv_exemplo):
    lista = carregar_destinatarios(csv_exemplo)

    assert [d.email for d in lista.destinatarios] == [
        "ana.souza@example.com",
        "bruno.lima@example.com",
        "carla.mendes@example.com",
    ]
    assert lista.avisos == []
    ana = lista.destinatarios[0]
    assert ana.dados["nome"] == "Ana Souza"
    assert "email" not in ana.dados and "anexos" not in ana.dados
    assert ana.anexos[0].name == "programacao.txt" and ana.anexos[0].is_file()
    assert lista.destinatarios[1].anexos == ()


def test_ponto_e_virgula_bom_e_cabecalho_normalizado(tmp_path):
    caminho = escrever(tmp_path, " Email ;Nome Completo\nana@exemplo.com;Ana\n", encoding="utf-8-sig")

    lista = carregar_destinatarios(caminho)

    assert lista.destinatarios[0].email == "ana@exemplo.com"
    assert lista.destinatarios[0].dados == {"nome_completo": "Ana"}


def test_linhas_problematicas_sao_ignoradas_e_reportadas(tmp_path):
    caminho = escrever(
        tmp_path,
        "email,nome\n"
        "ana@exemplo.com,Ana\n"
        ",Sem Email\n"
        "invalido,Fulano\n"
        "ANA@exemplo.com,Ana de novo\n"
        "leo@exemplo.com,Leo\n",
    )

    lista = carregar_destinatarios(caminho)

    assert [d.email for d in lista.destinatarios] == ["ana@exemplo.com", "leo@exemplo.com"]
    assert len(lista.avisos) == 3
    assert lista.avisos[0].startswith("Linha 3")
    assert "repetido" in lista.avisos[2]


def test_multiplos_anexos_relativos_ao_csv(tmp_path):
    caminho = escrever(tmp_path, "email,anexos\nana@exemplo.com,docs/a.pdf | b.txt\n")

    destinatario = carregar_destinatarios(caminho).destinatarios[0]

    assert destinatario.anexos == ((tmp_path / "docs" / "a.pdf").resolve(), (tmp_path / "b.txt").resolve())


def test_csv_sem_coluna_email(tmp_path):
    with pytest.raises(ValueError, match="coluna 'email'"):
        carregar_destinatarios(escrever(tmp_path, "nome\nAna\n"))


def test_csv_inexistente_ou_vazio(tmp_path):
    with pytest.raises(FileNotFoundError):
        carregar_destinatarios(tmp_path / "nao_existe.csv")
    with pytest.raises(ValueError):
        carregar_destinatarios(escrever(tmp_path, "\n"))
