my_list = [5, 1, 4, 2, 8]
swapped = True

while swapped:
    swapped = False
    for i in range (len(my_list) - 1):
        if my_list[i] > my_list[i + 1]:
            swapped = True
            my_list[i], my_list[i + 1] = my_list[i + 1], my_list[i]

        print(my_list)
print("Po pruchodu:", my_list)
