"""Interface de linha de comando.

Exemplos:
    email-sender enviar --para ana@exemplo.com --assunto "Olá" --mensagem "Oi, {{ nome }}!" --var nome=Ana
    email-sender lote --csv exemplos/destinatarios.csv --template lembrete_evento \\
        --assunto "Lembrete: {{ evento }}" --dry-run
    email-sender testar-conexao
    email-sender templates
"""

from __future__ import annotations

import argparse
import logging
import sys
from collections.abc import Sequence
from pathlib import Path

from email_sender import __version__
from email_sender.campanha import Campanha, Relatorio, executar
from email_sender.config import ConfigError, EnvioConfig, SMTPConfig, carregar_variaveis
from email_sender.message import MensagemError, validar_email
from email_sender.rate_limit import LimitadorTaxa
from email_sender.recipients import Destinatario, carregar_destinatarios
from email_sender.sender import EnviadorSMTP, EnvioError
from email_sender.templates import RenderizadorTemplates

logger = logging.getLogger("email_sender")

SAIDA_OK, SAIDA_FALHA_ENVIO, SAIDA_ERRO_USO = 0, 1, 2


def configurar_logging(verbose: bool, arquivo: Path | None) -> None:
    formato = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s", datefmt="%H:%M:%S")
    raiz = logging.getLogger("email_sender")
    raiz.handlers.clear()
    raiz.setLevel(logging.DEBUG if verbose else logging.INFO)
    raiz.propagate = False

    console = logging.StreamHandler(sys.stderr)
    console.setFormatter(formato)
    raiz.addHandler(console)
    if arquivo is not None:
        arquivo.parent.mkdir(parents=True, exist_ok=True)
        manipulador = logging.FileHandler(arquivo, encoding="utf-8")
        manipulador.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
        raiz.addHandler(manipulador)


def _par_chave_valor(texto: str) -> tuple[str, str]:
    chave, sep, valor = texto.partition("=")
    if not sep or not chave.strip():
        raise argparse.ArgumentTypeError(f"use o formato chave=valor (recebido: {texto!r})")
    return chave.strip(), valor


def criar_parser() -> argparse.ArgumentParser:
    comum = argparse.ArgumentParser(add_help=False)
    comum.add_argument("--env-file", type=Path, help="arquivo .env (padrão: .env no diretório atual)")
    comum.add_argument("-v", "--verbose", action="store_true", help="exibe logs detalhados")
    comum.add_argument("--log-arquivo", type=Path, help="também grava os logs neste arquivo")

    conteudo = argparse.ArgumentParser(add_help=False)
    conteudo.add_argument("--assunto", required=True, help="assunto; aceita variáveis, ex.: 'Olá, {{ nome }}'")
    corpo = conteudo.add_mutually_exclusive_group(required=True)
    corpo.add_argument("--template", help="nome do template em --pasta-templates (sem extensão)")
    corpo.add_argument("--mensagem", help="corpo em texto puro informado diretamente")
    conteudo.add_argument("--pasta-templates", type=Path, default=Path("templates"), help="padrão: ./templates")
    conteudo.add_argument("--anexo", action="append", type=Path, default=[], help="arquivo anexo (repetível)")
    conteudo.add_argument(
        "--var", action="append", type=_par_chave_valor, default=[], metavar="CHAVE=VALOR", help="variável extra"
    )
    conteudo.add_argument("--responder-para", help="endereço do cabeçalho Reply-To")
    conteudo.add_argument("--descadastro", help="URL ou mailto para o cabeçalho List-Unsubscribe")
    conteudo.add_argument("--dry-run", action="store_true", help="simula: monta as mensagens sem enviar")
    conteudo.add_argument("--salvar-eml", type=Path, metavar="PASTA", help="salva cada mensagem como .eml")
    conteudo.add_argument("--tentativas", type=int, help="máximo de tentativas por envio (padrão: 3)")

    parser = argparse.ArgumentParser(
        prog="email-sender",
        description="Envio de e-mails via SMTP com templates, anexos e envio em lote a partir de CSV.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = parser.add_subparsers(dest="comando", required=True)

    enviar = sub.add_parser("enviar", parents=[comum, conteudo], help="envia para um ou mais endereços")
    enviar.add_argument("--para", action="append", required=True, help="destinatário (repetível)")

    lote = sub.add_parser("lote", parents=[comum, conteudo], help="envio personalizado a partir de um CSV")
    lote.add_argument("--csv", type=Path, required=True, help="CSV com a coluna 'email' e variáveis")
    lote.add_argument("--intervalo", type=float, help="segundos entre envios (padrão: 1.0)")
    lote.add_argument("--limite", type=int, help="processa no máximo N destinatários")
    lote.add_argument("--relatorio", type=Path, help="grava o resultado de cada envio em CSV")

    sub.add_parser("testar-conexao", parents=[comum], help="conecta e autentica, sem enviar nada")

    templates = sub.add_parser("templates", parents=[comum], help="lista os templates disponíveis")
    templates.add_argument("--pasta-templates", type=Path, default=Path("templates"))
    return parser


def _destinatarios(args: argparse.Namespace) -> list[Destinatario]:
    if args.comando == "enviar":
        return [Destinatario(email=validar_email(email)) for email in args.para]
    lista = carregar_destinatarios(args.csv)
    for aviso in lista.avisos:
        logger.warning("%s: %s", args.csv.name, aviso)
    destinatarios = lista.destinatarios
    if args.limite is not None:
        if args.limite < 1:
            raise ValueError("--limite deve ser maior que zero.")
        if len(destinatarios) > args.limite:
            logger.info("Limite de %d destinatário(s) aplicado (%d no CSV).", args.limite, len(destinatarios))
        destinatarios = destinatarios[: args.limite]
    return destinatarios


def _executar_envio(args: argparse.Namespace) -> int:
    variaveis = carregar_variaveis(args.env_file)
    smtp_config = SMTPConfig.from_mapping(variaveis, exigir_credenciais=not args.dry_run)
    envio_config = EnvioConfig.from_mapping(variaveis)

    renderizador = RenderizadorTemplates(args.pasta_templates)
    # Validações antecipadas: evitam descobrir o problema só no meio do lote.
    if args.template and args.template not in renderizador.disponiveis():
        disponiveis = ", ".join(renderizador.disponiveis()) or "nenhum"
        raise ValueError(
            f"Template '{args.template}' não encontrado em {args.pasta_templates}. Disponíveis: {disponiveis}."
        )
    for anexo in args.anexo:
        if not anexo.is_file():
            raise FileNotFoundError(f"Anexo não encontrado: {anexo}")

    campanha = Campanha(
        remetente=smtp_config.remetente_formatado,
        assunto=args.assunto,
        renderizador=renderizador,
        template=args.template,
        mensagem=args.mensagem,
        anexos=tuple(args.anexo),
        variaveis=dict(args.var),
        responder_para=args.responder_para,
        descadastro=args.descadastro,
    )
    destinatarios = _destinatarios(args)
    if not destinatarios:
        logger.error("Nenhum destinatário válido para processar.")
        return SAIDA_FALHA_ENVIO

    intervalo = getattr(args, "intervalo", None)
    limitador = LimitadorTaxa(intervalo if intervalo is not None else envio_config.intervalo_segundos)

    if args.dry_run:
        logger.info("Modo simulação: nenhuma mensagem será enviada.")
        relatorio = executar(campanha, destinatarios, enviador=None, pasta_eml=args.salvar_eml)
    else:
        enviador = EnviadorSMTP(
            smtp_config,
            max_tentativas=args.tentativas or envio_config.max_tentativas,
            backoff_segundos=envio_config.backoff_segundos,
        )
        with enviador:
            relatorio = executar(
                campanha, destinatarios, enviador=enviador, limitador=limitador, pasta_eml=args.salvar_eml
            )

    _finalizar(relatorio, getattr(args, "relatorio", None))
    return SAIDA_FALHA_ENVIO if relatorio.houve_falha else SAIDA_OK


def _finalizar(relatorio: Relatorio, caminho: Path | None) -> None:
    logger.info(relatorio.resumo())
    if caminho is not None:
        relatorio.salvar_csv(caminho)
        logger.info("Relatório salvo em %s", caminho)


def _testar_conexao(args: argparse.Namespace) -> int:
    config = SMTPConfig.from_mapping(carregar_variaveis(args.env_file))
    with EnviadorSMTP(config, max_tentativas=1):
        logger.info("Conexão e autenticação OK. Remetente: %s", config.remetente_formatado)
    return SAIDA_OK


def _listar_templates(args: argparse.Namespace) -> int:
    nomes = RenderizadorTemplates(args.pasta_templates).disponiveis()
    if not nomes:
        logger.warning("Nenhum template encontrado em %s", args.pasta_templates)
    for nome in nomes:
        print(nome)
    return SAIDA_OK


def main(argv: Sequence[str] | None = None) -> int:
    args = criar_parser().parse_args(argv)
    configurar_logging(args.verbose, args.log_arquivo)
    comandos = {
        "enviar": _executar_envio,
        "lote": _executar_envio,
        "testar-conexao": _testar_conexao,
        "templates": _listar_templates,
    }
    try:
        return comandos[args.comando](args)
    except EnvioError as erro:
        logger.error("%s", erro)
        return SAIDA_FALHA_ENVIO
    except (ConfigError, MensagemError, FileNotFoundError, ValueError) as erro:
        logger.error("%s", erro)
        return SAIDA_ERRO_USO
    except KeyboardInterrupt:
        logger.warning("Interrompido pelo usuário.")
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
