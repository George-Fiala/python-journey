numbers = []
unique_numbers = []
count = 0
total = 0
total_unique = 0
found = False
target_index = None
min_value = None
max_value = None

while True:
    text = input("Enter number or stop:")
    if text =="stop":
        break

    value = int(text)
    count += 1
    total += value
    if min_value is None or value < min_value:
        min_value = value
    if max_value is None or value > max_value:
        max_value = value
    numbers.append(value)
    if value not in unique_numbers:
        unique_numbers.append(value)
        total_unique += value
if count == 0:
    print("No numbers entered")
    

else:
    target_text = input("Enter number we looking for:")
    target = int(target_text)
    for i in range(len(numbers)):
        if numbers[i] == target:
            found = True
            target_index = i
            break
    if found:
        print(f"Target index:{target_index}")
        print(f"Value:{numbers[target_index]}")
    
    else:
        print("Not found")

    average = total / count
    print(total, count, total_unique, average)
    print(f"Max value:{max_value}")
    print(f"Min value:{min_value}")
    print(numbers)
    print(unique_numbers)

