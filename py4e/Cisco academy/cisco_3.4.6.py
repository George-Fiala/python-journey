hat_list = [1, 2, 3, 4, 5]
hat_list[2] = int(input("Enter number: "))
print("New hat list:", hat_list)
del hat_list[-1]
print("Updated hat list:", hat_list)
print("Updated length of Hat list:", len(hat_list))     