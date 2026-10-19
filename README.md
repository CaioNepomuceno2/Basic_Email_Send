# email-sender-python

[![CI](https://github.com/CaioNepomuceno2/email-sender-python/actions/workflows/ci.yml/badge.svg)](https://github.com/CaioNepomuceno2/email-sender-python/actions/workflows/ci.yml)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![Lint: ruff](https://img.shields.io/badge/lint-ruff-261230)

Ferramenta de linha de comando em Python para enviar e-mails via SMTP com
templates (texto e HTML), anexos e **envio em lote personalizado a partir de um CSV**.
Inclui modo de simulação, controle de ritmo, novas tentativas automáticas e logs.

Usa apenas a biblioteca padrão (`smtplib`, `email`, `ssl`) mais
[Jinja2](https://jinja.palletsprojects.com/) para os templates e
[python-dotenv](https://github.com/theskumar/python-dotenv) para a configuração.

## Funcionalidades

- **Templates Jinja2** para assunto e corpo, em texto puro e/ou HTML (`multipart/alternative`).
  Sem a versão em texto, ela é gerada automaticamente a partir do HTML.
- **Envio em lote** a partir de CSV: cada coluna vira uma variável (`{{ nome }}`, `{{ evento }}`...).
  Aceita CSV com `,` ou `;` (padrão do Excel em português) e arquivos com BOM.
- **Anexos** globais (`--anexo`) e por destinatário (coluna `anexos` do CSV), com tipo MIME
  detectado e limite total de 25 MB.
- **Modo simulação** (`--dry-run`): monta todas as mensagens, valida templates e anexos, sem conectar ao servidor.
  Com `--salvar-eml`, grava cada mensagem como `.eml` para conferência em um cliente de e-mail.
- **Controle de ritmo** (rate limiting): intervalo mínimo entre envios e `--limite` de destinatários por execução.
- **Novas tentativas** com backoff exponencial para falhas temporárias (queda de conexão, timeout,
  códigos SMTP 4xx) e reconexão automática. Erros permanentes (5xx) não são repetidos, e falha de
  autenticação interrompe o lote.
- **Validação** de e-mails, remoção de duplicados e relatório das linhas ignoradas.
- **Logs** no terminal e, opcionalmente, em arquivo; **relatório CSV** com o status de cada envio.
- **TLS** via STARTTLS (porta 587) ou SSL (porta 465), com verificação de certificado.

## Estrutura

```
email-sender-python/
├── src/email_sender/
│   ├── cli.py          # interface de linha de comando (argparse)
│   ├── config.py       # leitura e validação do .env / variáveis de ambiente
│   ├── templates.py    # renderização Jinja2 e conversão HTML -> texto
│   ├── message.py      # montagem da mensagem MIME, anexos e validação de e-mail
│   ├── recipients.py   # leitura do CSV de destinatários
│   ├── sender.py       # conexão SMTP, TLS, login e novas tentativas
│   ├── rate_limit.py   # intervalo mínimo entre envios
│   └── campanha.py     # orquestra o envio e gera o relatório
├── templates/          # boas_vindas (txt + html), lembrete_evento (html)
├── exemplos/           # destinatarios.csv e anexo de exemplo
├── tests/              # pytest; o smtplib é simulado com mocks
├── .env.example
└── pyproject.toml
```

## Instalação

Requer Python 3.10 ou superior.

```bash
git clone https://github.com/CaioNepomuceno2/email-sender-python.git
cd email-sender-python

python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -e .                   # instala o comando `email-sender`
```

Também é possível executar sem instalar o comando: `pip install -r requirements.txt` e
`PYTHONPATH=src python -m email_sender ...`.

## Configuração

```bash
cp .env.example .env
```

Edite o `.env` com os dados do seu provedor:

| Variável | Descrição | Padrão |
|---|---|---|
| `SMTP_HOST` | Servidor SMTP | obrigatório |
| `SMTP_PORTA` | Porta | 587 (`starttls`) / 465 (`ssl`) |
| `SMTP_SEGURANCA` | `starttls`, `ssl` ou `nenhuma` (apenas servidores locais) | `starttls` |
| `SMTP_USUARIO` / `SMTP_SENHA` | Credenciais | obrigatórias se houver autenticação |
| `EMAIL_REMETENTE` | Endereço do remetente | `SMTP_USUARIO` |
| `EMAIL_NOME_REMETENTE` | Nome exibido no remetente | vazio |
| `SMTP_TIMEOUT` | Timeout da conexão, em segundos | 30 |
| `ENVIO_INTERVALO_SEGUNDOS` | Intervalo mínimo entre envios | 1.0 |
| `ENVIO_MAX_TENTATIVAS` | Tentativas por mensagem em falhas temporárias | 3 |
| `ENVIO_BACKOFF_SEGUNDOS` | Espera inicial entre tentativas (dobra a cada nova tentativa) | 2.0 |

Variáveis definidas no ambiente têm prioridade sobre o arquivo `.env`. Use
`--env-file caminho` para escolher outro arquivo.

### Gmail: senha de app

O Gmail não aceita a senha normal da conta para SMTP. É preciso gerar uma **senha de app**:

1. Ative a **verificação em duas etapas** na sua Conta Google.
2. Acesse <https://myaccount.google.com/apppasswords>, crie uma senha de app (ex.: "email-sender")
   e copie os 16 caracteres gerados.
3. No `.env`:
   ```env
   SMTP_HOST=smtp.gmail.com
   SMTP_PORTA=587
   SMTP_SEGURANCA=starttls
   SMTP_USUARIO=seu.email@gmail.com
   SMTP_SENHA=sua_senha_de_app
   ```
4. Teste com `email-sender testar-conexao`.

Contas corporativas (Google Workspace) podem ter senhas de app bloqueadas pelo administrador.

### Outlook / Microsoft 365

```env
SMTP_HOST=smtp.office365.com
SMTP_PORTA=587
SMTP_SEGURANCA=starttls
```

A Microsoft vem restringindo a autenticação básica (usuário e senha) no SMTP em favor de OAuth2.
Dependendo do tipo de conta e das políticas da organização, o login pode ser recusado mesmo com a
senha correta. Esta ferramenta suporta apenas usuário e senha; em contas Microsoft 365, o
administrador precisa manter o *SMTP AUTH* habilitado para a caixa de correio.

## Uso

Sempre comece pela **simulação**:

```bash
email-sender lote \
  --csv exemplos/destinatarios.csv \
  --template lembrete_evento \
  --assunto "Lembrete: {{ evento }}" \
  --dry-run
```

```
00:05:10 INFO    Modo simulação: nenhuma mensagem será enviada.
00:05:10 INFO    [1/3] [SIMULAÇÃO] para=ana.souza@example.com | assunto='Lembrete: Workshop de Python para Dados' | 2.8 KB | anexos: programacao.txt
00:05:10 INFO    [2/3] [SIMULAÇÃO] para=bruno.lima@example.com | assunto='Lembrete: Workshop de Python para Dados' | 2.2 KB
00:05:10 INFO    [3/3] [SIMULAÇÃO] para=carla.mendes@example.com | assunto='Lembrete: Workshop de Python para Dados' | 2.5 KB | anexos: programacao.txt
00:05:10 INFO    Total: 3 | 3 simulado
```

Envio real em lote, com intervalo de 2 s, relatório e log em arquivo:

```bash
email-sender lote --csv exemplos/destinatarios.csv --template lembrete_evento \
  --assunto "Lembrete: {{ evento }}" --intervalo 2 \
  --relatorio saida/relatorio.csv --log-arquivo saida/envio.log
```

Envio individual com mensagem direta, variável e anexo:

```bash
email-sender enviar --para ana@example.com \
  --assunto "Relatório mensal" \
  --mensagem "Olá, {{ nome }}! Segue o relatório em anexo." \
  --var nome=Ana --anexo relatorio.pdf
```

Outros comandos:

```bash
email-sender templates                     # lista os templates disponíveis
email-sender testar-conexao                # conecta e autentica, sem enviar nada
email-sender lote ... --dry-run --salvar-eml saida/eml   # grava as mensagens em .eml
email-sender lote ... --limite 50          # processa só os 50 primeiros
email-sender enviar --help                 # todas as opções
```

| Opção | Efeito |
|---|---|
| `--template NOME` / `--mensagem TEXTO` | corpo a partir de `templates/NOME.*.j2` ou de um texto direto |
| `--var CHAVE=VALOR` | variável extra para os templates (o CSV tem prioridade) |
| `--anexo ARQUIVO` | anexo enviado a todos (repetível) |
| `--responder-para EMAIL` | define o cabeçalho `Reply-To` |
| `--descadastro URL` | adiciona `List-Unsubscribe` (recomendado para envios em massa) |
| `--tentativas N` | sobrescreve `ENVIO_MAX_TENTATIVAS` |
| `-v` | logs detalhados, incluindo o corpo em texto na simulação |

**Códigos de saída:** `0` sucesso, `1` houve falha em algum envio, `2` erro de configuração ou de uso.

### Formato do CSV

```csv
email,nome,evento,data,horario,link,anexos
ana.souza@example.com,Ana Souza,Workshop de Python para Dados,20/03/2027,19h00,https://example.com/evento/python,anexos/programacao.txt
```

- `email` é obrigatória; as demais colunas viram variáveis (nomes em minúsculas, espaços viram `_`).
- `anexos` (opcional): um ou mais caminhos separados por `|`, relativos à pasta do CSV.
- Linhas com e-mail vazio, inválido ou repetido são ignoradas e listadas como aviso.

### Criando templates

Crie `templates/<nome>.txt.j2` e/ou `templates/<nome>.html.j2`:

```jinja
Olá, {{ nome }}!
{% if empresa is defined and empresa %}Bem-vindo(a) à {{ empresa }}.{% endif %}
{{ assinatura | default("Equipe de Atendimento") }}
```

- Variável usada e não fornecida gera erro (Jinja2 `StrictUndefined`), evitando e-mails com lacunas.
- No HTML, os valores são escapados automaticamente, então dados do CSV não injetam marcação.

## Testes

```bash
pip install -r requirements-dev.txt
pytest
ruff check . && ruff format --check .
```

Os testes substituem o `smtplib` por mocks (`unittest.mock`), então **nenhum e-mail é enviado**.
Eles cobrem configuração, templates, montagem MIME, leitura do CSV, rate limiting, novas tentativas,
classificação de erros SMTP e a CLI de ponta a ponta. O workflow do GitHub Actions roda lint,
testes (Python 3.10, 3.12 e 3.13) e uma simulação do envio em lote a cada push.

## Segurança e boas práticas

- **Nunca versione o `.env`**: ele já está no `.gitignore`. O repositório contém apenas o `.env.example`, sem credenciais.
- Use **senhas de app** em vez da senha principal da conta e revogue-as quando não forem mais usadas.
- A senha não aparece em logs nem na representação dos objetos de configuração.
- A conexão usa TLS com verificação de certificado (`ssl.create_default_context()`). Enviar senha sem
  TLS é bloqueado, exceto para servidores locais de teste.
- Quebras de linha são removidas do assunto, e endereços com espaços ou quebras de linha são rejeitados,
  o que impede injeção de cabeçalhos.
- Rode `--dry-run` antes de qualquer envio em lote e comece com um `--limite` pequeno.
- Respeite os limites diários de envio do seu provedor e envie apenas para quem consentiu em receber
  as mensagens (LGPD), oferecendo uma forma de descadastro (`--descadastro`).

## Autor

**Caio Henrique Nepomuceno Santos**: [github.com/CaioNepomuceno2](https://github.com/CaioNepomuceno2)
