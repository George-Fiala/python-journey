numbers = []
unique_numbers = []
found = False
found_index = None

while True:
    text = input("Enter number or done:")
    if text == "done":
        break

    value = int(text)
    numbers.append(value)
    if value not in unique_numbers:
        unique_numbers.append(value)

target_text = input("Enter number we looking for:")
target = int(target_text)

for i in range(len(unique_numbers)):
    if unique_numbers[i] == target:
        found = True
        found_index = i
        break

if found:
    print(f"Numbers:{numbers}")
    print(f"Unique numbers:{unique_numbers}")
    print(f"Found index= {found_index}")
    print(f"Value:{target}")

else:
    print("Number not found.")
