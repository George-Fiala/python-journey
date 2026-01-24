board = [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
]
from random import randrange

board[1][1] = "X"

def display_board(board):
    for b in board:
        print(b)

def enter_move(board):
    move = int(input("Enter Your move (1-9): "))
    for row in range(3):
        for col in range(3):
            if board[row][col] == move:
                board[row][col] = "O"

def make_list_of_free_fields(board):
    free = []
    for row in range(3):
        for col in range(3):
            if board[row][col] != "X" and board[row][col] != "O":
                free.append((row, col))   
    return free

def draw_move(board):
    free = make_list_of_free_fields(board)
    if len(free) > 0:
        idx = randrange(len(free)) 
        row, col = free[idx]
        board[row][col] = "X" 

def victory_for(board, sign):
    for r in range(3):
        if board[r][0] == sign and board[r][1] == sign and board[r][2] == sign:
            return True
    for c in range(3):
        if board[0][c] == sign and board[1][c] == sign and board[2][c] == sign:
            return True
    if board[0][0] == sign and board[1][1] == sign and board[2][2] == sign:
        return True
    if board[0][2] == sign and board[1][1] == sign and board[2][0] == sign:
        return True
    
    return False

human_turn = True

while True:
    display_board(board)

    if human_turn:
        enter_move(board)
        sign = 'O'
    else:
        print("PC is playing")
        draw_move(board)
        sign = 'X'

    if victory_for(board, sign):
        display_board(board)
        print("Player won", sign)
        break

    free = make_list_of_free_fields(board)
    if len(free) == 0:
        display_board(board)
        print("Tie! Noone won.")
        break

    human_turn = not human_turn
