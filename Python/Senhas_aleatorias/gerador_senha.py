import random
import string

def generate_password(length, use_uppercase, use_lowercase, use_digits, use_specials):
    # Criando o conjunto de caracteres possíveis baseado nas escolhas do usuário
    possible_chars = ""
    if use_uppercase:
        possible_chars += string.ascii_uppercase
    if use_lowercase:
        possible_chars += string.ascii_lowercase
    if use_digits:
        possible_chars += string.digits
    if use_specials:
        possible_chars += string.punctuation

    # Gerando a senha
    password = ''.join(random.choice(possible_chars) for _ in range(length))
    return password

def main():
    print("Bem-vindo ao Gerador de Senhas!")
    length = int(input("Digite o comprimento da senha desejada: "))
    use_uppercase = input("Incluir letras maiúsculas? (s/n): ").lower() == 's'
    use_lowercase = input("Incluir letras minúsculas? (s/n): ").lower() == 's'
    use_digits = input("Incluir números? (s/n): ").lower() == 's'
    use_specials = input("Incluir símbolos especiais? (s/n): ").lower() == 's'

    password = generate_password(length, use_uppercase, use_lowercase, use_digits, use_specials)
    print("Sua nova senha é:", password)

if __name__ == "__main__":
    main()
