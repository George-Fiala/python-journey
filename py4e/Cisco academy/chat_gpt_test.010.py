numbers = [10, 20, 30, 40, 50]
target = 40

found = False
found_index = None

for i in range(len(numbers)):
    if numbers[i] == target:
        found = True
        found_index = i
        break

if found:
    print(f"Found index = {found_index}")

else:
    print("Index not found")
