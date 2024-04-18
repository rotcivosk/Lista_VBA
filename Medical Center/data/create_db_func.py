import sqlite3

def create_db():
    # Conectar ao banco de dados SQLite (será criado se não existir)
    conn = sqlite3.connect('contacts.db')
    # Criar um cursor para executar comandos SQL
    cursor = conn.cursor()
    # Criar a tabela de contatos se não existir
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS contacts (
            idFuncionario INTEGER PRIMARY KEY,
            nomeFuncionario TEXT NOT NULL,
            cargoFuncionario TEXT NOT NULL,
            dataAdmissao TEXT
        )
    ''')
    # Salvar as mudanças e fechar a conexão com o banco
    conn.commit()
    conn.close()

create_db()