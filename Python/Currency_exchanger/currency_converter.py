import requests

def get_rates(base_currency):
    """
    Recebe o input que quer cotar, e recebe a informação do url do exchangerate-api que é conforme o exemplo em anexo
    """
    url = f"https://v6.exchangerate-api.com/v6/b416e66d5b3a34228460f851/latest/{base_currency}"
    response = requests.get(url)
    data = response.json()
    return data['conversion_rates']

def main():
    # Inputs no geral é com input()
    base_currency = input("Enter the base currency (BRL, USD etc): ").upper() #upper para ser tudo maiúsculo
    rates = get_rates(base_currency)
    amount = float(input("Enter the amount to convert: "))
    target_currency = input("Enter the target currency (BRL, USD etc): ").upper()

    # calcula
    converted_amount = amount * rates[target_currency]
    print(f"{amount} {base_currency} is equal to {converted_amount:.2f} {target_currency}")

if __name__ == "__main__":
    # Entry point of the program
    main()
