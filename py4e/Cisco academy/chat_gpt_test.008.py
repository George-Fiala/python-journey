numbers = [10, 20, 30, 40]
target = 30

found = False
found_index = None

for i in range(len(numbers)):
    if numbers[i] == target:
        found = True
        found_index = i
        break

if found:
    print(f"Found index is: {found_index}")

else:
    print(f"Not found")

