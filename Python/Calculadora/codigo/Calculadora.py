def somar(a,b):
    return a + b
def subtrair(a,b):
    return a - b
def multiplicar (a,b): 
    return a ** b
def dividir(a,b):
    if b == 0:
        raise ValueError("Dividir por 0, não, né")
    return a / b
def potencia(a,b):
    return a**b

def main():
    while True:
        print("\nOperações disponíveis:")
        print("1. Somar")
        print("2. Subtrair")
        print("3. Multiplicar")
        print("4. Dividir")
        print("5. Potência")
        print("0. Sair")
        
        opcao = input("Escolha a operação (1/2/3/4/5) ou 0 para sair: ")
        
        if opcao == '0':
            print("Saindo...")
            break
        
        a = float(input("Digite o primeiro número: "))
        b = float(input("Digite o segundo número: "))
        
        if opcao == '1':
            print("Resultado:", somar(a, b))
        elif opcao == '2':
            print("Resultado:", subtrair(a, b))
        elif opcao == '3':
            print("Resultado:", multiplicar(a, b))
        elif opcao == '4':
            print("Resultado:", dividir(a, b))
        elif opcao == '5':
            print("Resultado:", potencia(a, b))
        else:
            print("Fora do Range")

if __name__ == "__main__":
    main()