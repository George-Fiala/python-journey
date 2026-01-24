numbers = []

found = False
target_index = None

while True:
    text = input("Enter number or stop:")
    if text =="stop":
        break
    value = int(text)
    numbers.append(value)

target_text = input("Enter number that we looking for:")
target = int(target_text)
for i in range(len(numbers)):
    if numbers[i] == target:
        found = True
        target_index = i
        break
if found:
    print(f"Numbers={numbers}")
    print(f"Value:{numbers[target_index]}")
    print("Target index:", target_index)
else:
    print("Number not found.")

