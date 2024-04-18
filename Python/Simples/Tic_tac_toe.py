def print_board(board):
    for row in board:
        print(" | ".join(row))
        print("-" * 9)

def check_win(board):
    # Verifique as linhas
    for row in board:
        if row.count(row[0]) == len(row) and row[0] != " ":
            return True
    # Verifique as colunas
    for col in range(len(board[0])):
        check = []
        for row in board:
            check.append(row[col])
        if check.count(check[0]) == len(check) and check[0] != " ":
            return True
    # Verifique as diagonais
    if board[0][0] == board[1][1] == board[2][2] and board[0][0] != " ":
        return True
    if board[0][2] == board[1][1] == board[2][0] and board[0][2] != " ":
        return True
    return False

def tic_tac_toe():
    board = [[" " for _ in range(3)] for _ in range(3)]
    player = "X"

    while True:
        print_board(board)
        print("Vez do jogador", player)
        row = int(input("Digite o número da linha: "))
        col = int(input("Digite o número da coluna: "))
        if board[row][col] == " ":
            board[row][col] = player
            if check_win(board):
                print_board(board)
                print("Parabéns, jogador", player, "você ganhou!")
                break
            player = "X" if player == "O" else "O"
        else:
            print("Posição inválida, tente novamente.")

tic_tac_toe()
