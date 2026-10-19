"""Controle de ritmo de envio (rate limiting).

Provedores como Gmail e Outlook limitam a quantidade de mensagens por minuto e
por dia. Enviar rápido demais pode gerar bloqueio temporário da conta, por isso
garantimos um intervalo mínimo entre envios consecutivos.
"""

from __future__ import annotations

import time
from collections.abc import Callable


class LimitadorTaxa:
    def __init__(
        self,
        intervalo_segundos: float,
        *,
        relogio: Callable[[], float] = time.monotonic,
        dormir: Callable[[float], None] = time.sleep,
    ) -> None:
        if intervalo_segundos < 0:
            raise ValueError("O intervalo não pode ser negativo.")
        self.intervalo = intervalo_segundos
        self._relogio = relogio  # injetáveis para facilitar os testes
        self._dormir = dormir
        self._ultimo: float | None = None

    def aguardar(self) -> float:
        """Bloqueia o tempo necessário para respeitar o intervalo. Retorna a espera aplicada."""
        espera = 0.0
        if self._ultimo is not None:
            espera = self.intervalo - (self._relogio() - self._ultimo)
            if espera > 0:
                self._dormir(espera)
            else:
                espera = 0.0
        self._ultimo = self._relogio()
        return espera
