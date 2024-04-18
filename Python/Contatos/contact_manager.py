import sqlite3

def add_contact(name, phone, email):
    try:
        conn = sqlite3.connect('contacts.db')
        cursor = conn.cursor()
        cursor.execute("INSERT INTO contacts (name, phone, email) VALUES (?, ?, ?)", (name, phone, email))
        conn.commit()
    except sqlite3.Error as e:
        print(f"An error occurred: {e}")
    finally:
        conn.close()


def edit_contact(contact_id, new_name, new_phone, new_email):
    conn = sqlite3.connect('contacts.db')
    cursor = conn.cursor()
    cursor.execute("UPDATE contacts SET name = ?, phone = ?, email = ? WHERE id = ?", (new_name, new_phone, new_email, contact_id))
    conn.commit()
    conn.close()


def delete_contact(contact_id):
    conn = sqlite3.connect('contacts.db')
    cursor = conn.cursor()
    cursor.execute("DELETE FROM contacts WHERE id = ?", (contact_id,))
    conn.commit()
    conn.close()


def search_contacts(search_term):
    conn = sqlite3.connect('contacts.db')
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM contacts WHERE name LIKE ? OR phone LIKE ? OR email LIKE ?", ('%'+search_term+'%', '%'+search_term+'%', '%'+search_term+'%'))
    results = cursor.fetchall()
    conn.close()
    return results

def main():
    while True:
        print("\nGerenciador de Contatos")
        print("1. Adicionar Contato")
        print("2. Editar Contato")
        print("3. Deletar Contato")
        print("4. Procurar Contato")
        print("5. Sair")
        choice = input("Escolha uma opção: ")

        if choice == '1':
            name = input("Nome: ")
            phone = input("Telefone: ")
            email = input("Email: ")
            add_contact(name, phone, email)
            print("Contato adicionado com sucesso.")

        elif choice == '2':
            contact_id = input("Qual ID que gostaria de alterar? ")
            new_name = input("Novo nome: ")
            new_phone = input("Novo telefone: ")
            new_email = input("Novo email: ")
            edit_contact(contact_id, new_name, new_phone, new_email)
            print("Contato atualizado com sucesso.")

        elif choice == '3':
            contact_id = input("Qual id do contato gostaria de deletar? ")
            delete_contact(contact_id)
            print("Contato deletado com sucesso.")

        elif choice == '4':
            search_term = input("Procurar: ")
            results = search_contacts(search_term)
            if results:
                print("\nResultados encontrados:")
                for result in results:
                    print(f"ID: {result[0]}, Nome: {result[1]}, Telefone: {result[2]}, Email: {result[3]}")
            else:
                print("Nenhum contato encontrado.")

        elif choice == '5':
            print("Saindo do gerenciador de contatos.")
            break

        else:
            print("Opção inválida")

if __name__ == "__main__":
    main()

