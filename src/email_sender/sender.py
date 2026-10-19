"""Conexão SMTP com TLS, autenticação e reenvio automático em falhas temporárias."""

from __future__ import annotations

import contextlib
import logging
import smtplib
import ssl
import time
from collections.abc import Callable
from email.message import EmailMessage

from email_sender.config import SMTPConfig

logger = logging.getLogger(__name__)


class EnvioError(Exception):
    """Falha ao enviar. `fatal=True` indica que não adianta tentar os próximos envios."""

    def __init__(self, mensagem: str, *, fatal: bool = False) -> None:
        super().__init__(mensagem)
        self.fatal = fatal


def eh_erro_temporario(erro: BaseException) -> bool:
    """Decide se vale a pena tentar novamente.

    - Códigos SMTP 4xx são temporários (ex.: 421 serviço indisponível, 451 limite atingido).
    - Códigos 5xx são permanentes (ex.: 550 caixa inexistente, 535 autenticação inválida).
    - Quedas de conexão e timeouts são temporários.
    """
    if isinstance(erro, smtplib.SMTPAuthenticationError):
        return False
    if isinstance(erro, smtplib.SMTPRecipientsRefused):
        codigos = [codigo for codigo, _ in erro.recipients.values()]
        return bool(codigos) and all(400 <= codigo < 500 for codigo in codigos)
    if isinstance(erro, smtplib.SMTPResponseException):
        return 400 <= erro.smtp_code < 500
    if isinstance(erro, smtplib.SMTPServerDisconnected):
        return True
    if isinstance(erro, smtplib.SMTPException):
        return False
    return isinstance(erro, OSError)  # timeout, conexão recusada/resetada, DNS


class EnviadorSMTP:
    """Mantém uma conexão aberta para vários envios e reconecta quando necessário.

    Uso:
        with EnviadorSMTP(config) as enviador:
            enviador.enviar(mensagem)
    """

    def __init__(
        self,
        config: SMTPConfig,
        *,
        max_tentativas: int = 3,
        backoff_segundos: float = 2.0,
        dormir: Callable[[float], None] = time.sleep,
    ) -> None:
        if max_tentativas < 1:
            raise ValueError("max_tentativas deve ser pelo menos 1.")
        self.config = config
        self.max_tentativas = max_tentativas
        self.backoff_segundos = backoff_segundos
        self._dormir = dormir
        self._smtp: smtplib.SMTP | None = None

    # ---------------- conexão ----------------

    def conectar(self) -> None:
        cfg = self.config
        contexto_tls = ssl.create_default_context()  # valida o certificado do servidor
        logger.debug("Conectando a %s:%s (%s)", cfg.host, cfg.porta, cfg.seguranca)
        if cfg.seguranca == "ssl":
            smtp = smtplib.SMTP_SSL(cfg.host, cfg.porta, timeout=cfg.timeout, context=contexto_tls)
        else:
            smtp = smtplib.SMTP(cfg.host, cfg.porta, timeout=cfg.timeout)
        try:
            smtp.ehlo()
            if cfg.seguranca == "starttls":
                smtp.starttls(context=contexto_tls)
                smtp.ehlo()  # o servidor anuncia novos recursos após o TLS
            if cfg.usuario:
                smtp.login(cfg.usuario, cfg.senha or "")
        except BaseException:
            self._fechar_silenciosamente(smtp)
            raise
        self._smtp = smtp
        logger.info("Conectado a %s:%s como %s", cfg.host, cfg.porta, cfg.usuario or "(sem autenticação)")

    def fechar(self) -> None:
        if self._smtp is not None:
            try:
                self._smtp.quit()
            except smtplib.SMTPException:
                self._fechar_silenciosamente(self._smtp)
            finally:
                self._smtp = None

    @staticmethod
    def _fechar_silenciosamente(smtp: smtplib.SMTP) -> None:
        with contextlib.suppress(Exception):  # o objetivo é apenas liberar o socket
            smtp.close()

    def __enter__(self) -> EnviadorSMTP:
        self.conectar_com_retentativa()
        return self

    def __exit__(self, *exc_info) -> None:
        self.fechar()

    def conectar_com_retentativa(self) -> None:
        self._executar_com_retentativa(self.conectar, "conexão")

    # ---------------- envio ----------------

    def enviar(self, mensagem: EmailMessage) -> None:
        def tentativa() -> None:
            if self._smtp is None:
                self.conectar()
            self._smtp.send_message(mensagem)

        self._executar_com_retentativa(tentativa, f"envio para {mensagem['To']}")

    def _executar_com_retentativa(self, operacao: Callable[[], None], descricao: str) -> None:
        for numero in range(1, self.max_tentativas + 1):
            try:
                operacao()
                return
            except smtplib.SMTPAuthenticationError as erro:
                raise EnvioError(
                    f"Falha de autenticação ({erro.smtp_code}). Verifique SMTP_USUARIO e SMTP_SENHA "
                    "(no Gmail, use uma senha de app).",
                    fatal=True,
                ) from erro
            except (smtplib.SMTPException, OSError) as erro:
                if not eh_erro_temporario(erro):
                    raise EnvioError(f"Erro permanente no {descricao}: {erro}") from erro
                # Conexão possivelmente corrompida: descarta e reconecta na próxima tentativa.
                if self._smtp is not None:
                    self._fechar_silenciosamente(self._smtp)
                    self._smtp = None
                if numero == self.max_tentativas:
                    raise EnvioError(
                        f"Falha no {descricao} após {self.max_tentativas} tentativa(s): {erro}",
                        fatal=isinstance(erro, OSError) and not isinstance(erro, smtplib.SMTPException),
                    ) from erro
                espera = self.backoff_segundos * 2 ** (numero - 1)  # backoff exponencial: 2s, 4s, 8s...
                logger.warning(
                    "Erro temporário no %s (tentativa %d/%d): %s. Nova tentativa em %.1fs.",
                    descricao,
                    numero,
                    self.max_tentativas,
                    erro,
                    espera,
                )
                self._dormir(espera)
