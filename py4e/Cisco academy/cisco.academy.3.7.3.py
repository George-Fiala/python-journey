buildings = 3
floors = 15
rooms_per_floor = 20
free_rooms = 0
rooms = [[[False for r in range(20)] for f in range(15)] for b in range(3)]

rooms[0][3][5] = True
rooms[1][7][18] = True
rooms[2][13][0] = True

for buildings in rooms:
    for floors in buildings:
        for rooms_per_floor in floors:
            if not rooms_per_floor:
                free_rooms += 1
print(free_rooms)