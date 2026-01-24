my_list = [1, 2, 4, 4, 1, 4, 2, 6, 2, 9]
unique = []

for x in my_list:
    if x not in unique:
        unique.append(x)

print(f"Unique list: {unique}")