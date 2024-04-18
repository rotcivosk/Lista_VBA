import random

def escolher_palavra():
    palavras = ['python', 'hangman', 'computador', 'programacao', 'teclado']
    return random.choice(palavras)

def iniciar_jogo():
    palavra = escolher_palavra()
    palavra_oculta = ['_' for _ in palavra]
    tentativas = 0
    max_tentativas = 6  # Supondo um jogo de forca com 6 partes do corpo
    tentativas_erradas = []
    
    print("Bem-vindo ao jogo de Forca!")
    print("Tente adivinhar a palavra. Boa sorte!")
    
    return palavra, palavra_oculta, tentativas, max_tentativas, tentativas_erradas

def exibir_estado(palavra_oculta, tentativas_erradas):
    print("Palavra: " + ' '.join(palavra_oculta))
    print("Erros: " + ', '.join(tentativas_erradas))
    print(f"Tentativas restantes: {max_tentativas - len(tentativas_erradas)}")

# Inicia o jogo
palavra, palavra_oculta, tentativas, max_tentativas, tentativas_erradas = iniciar_jogo()
exibir_estado(palavra_oculta, tentativas_erradas)