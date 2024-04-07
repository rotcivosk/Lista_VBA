import pytest
from calculadora import somar, subtrair, multiplicar, dividir, potencia

def test_somar():
    assert somar(3,4) == 7
def test_subtrair():
    assert subtrair(9, 11) == -2
def test_multiplicar():
    assert multiplicar(9, 0.5) == 4.5
def test_dividir():
    assert dividir(8,2) == 4
def test_potencia():
    assert potencia(4,2) == 16
def test_dividir_por_zero():
    with pytest.raises(ValueError):
        dividir(9,0)
