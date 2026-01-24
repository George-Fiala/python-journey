EMPTY = 0

PAWN = 1

ROOK = 2

KNIGHT = 3

BISHOP = 4

QUEEN = 5

KING = 6

pawn_count = 0
rook_count = 0
total_pieces = 0


board = [[EMPTY for i in range(8)] for j in range(8)]
board[0][0] = ROOK
board[0][1] = KNIGHT
board[0][2] = BISHOP
board[0][3] = QUEEN
board[0][4] = KING
board[0][5] = BISHOP
board[0][6] = KNIGHT
board[0][7] = ROOK
board[1] = [PAWN for i in range(8)]

for row in board:
    for piece in row:
        if piece == PAWN:
            pawn_count += 1
        if piece == ROOK:
            rook_count += 1
        if piece != EMPTY:
            total_pieces += 1

        

for row in board:
    print(row)
print(pawn_count)
print(rook_count)
print(f"Total pices:{total_pieces}")