numbers = []

while True:
    text = input("Enter number or stop:")
    if text == "stop":
        break
    value = float(text)
    numbers.append(value)

target_text = input("Enter number to search:")
target = float(target_text)

found = False
found_index = None
for i in range(len(numbers)):
    if numbers[i] == target:
        found = True
        found_index = i
        break
if found:
    print(f"Found index: {found_index}")

else:
    print("Not found")

print(numbers)