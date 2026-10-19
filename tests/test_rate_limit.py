import pytest

from email_sender.rate_limit import LimitadorTaxa


class RelogioFalso:
    def __init__(self):
        self.agora = 100.0
        self.esperas = []

    def __call__(self):
        return self.agora

    def dormir(self, segundos):
        self.esperas.append(segundos)
        self.agora += segundos


def test_primeira_chamada_nao_espera_e_seguintes_respeitam_intervalo():
    relogio = RelogioFalso()
    limitador = LimitadorTaxa(2.0, relogio=relogio, dormir=relogio.dormir)

    assert limitador.aguardar() == 0
    relogio.agora += 0.5  # o envio levou 0,5 s
    assert limitador.aguardar() == pytest.approx(1.5)
    relogio.agora += 3  # já passou mais que o intervalo
    assert limitador.aguardar() == 0
    assert relogio.esperas == [pytest.approx(1.5)]


def test_intervalo_zero_nunca_espera():
    relogio = RelogioFalso()
    limitador = LimitadorTaxa(0, relogio=relogio, dormir=relogio.dormir)
    for _ in range(3):
        limitador.aguardar()
    assert relogio.esperas == []


def test_intervalo_negativo():
    with pytest.raises(ValueError):
        LimitadorTaxa(-1)
