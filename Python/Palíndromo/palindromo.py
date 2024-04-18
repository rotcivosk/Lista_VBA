import string

def clean_input(text):
    # Remove espaços e pontuações e converte para minúsculas
    translator = str.maketrans('', '', string.punctuation)
    cleaned_text = text.translate(translator).replace(' ', '').lower()
    return cleaned_text

def is_palindrome(text):
    # Verifica se o texto é um palíndromo
    cleaned_text = clean_input(text)
    return cleaned_text == cleaned_text[::-1]

def main():
    user_input = input("Digite uma palavra ou frase para verificar se é um palíndromo: ")
    if is_palindrome(user_input):
        print("É um palíndromo!")
    else:
        print("Não é um palíndromo.")

if __name__ == "__main__":
    main()
