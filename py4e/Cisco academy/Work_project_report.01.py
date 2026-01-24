def save_int(text):
    try:
        return int(text)
    except ValueError:
        return None
    
numbers = []
count = 0
total = 0
average = 0
even_numbers = []
min_number = None
max_number = None

while True:
    text = input("Enter number or done: ")
    if text == "done":
        break

    value = save_int(text)
    if value is None:
        print("Enter number.")
        continue
    count += 1
    total += value
    numbers.append(value)
    if value % 2 == 0:
        even_numbers.append(value)

    if min_number is None or value < min_number:
        min_number = value
    if max_number is None or value > max_number:
        max_number = value

if count == 0:
    print("No numbers entered.")

else:
    average = total / count
    print("Numbers:", numbers)
    print("Count:", count)
    print("Total:", total)
    print("Average:", average)
    print("Even numbers:", even_numbers)
    print("Max number:", max_number)
    print("Min number:", min_number)

    
    while True:
        found = False
        target_index = None

        target_text = input("Enter target number or done: ")
        if target_text == "done":
            break
        target = save_int(target_text)
        
        if target is None:
            print("Enter number.")
            continue
        
        for i in range(len(numbers)):
            if numbers[i] == target:
                found = True
                target_index = i
                break
        if found:
            print("Target index:", target_index)
            print("Target number:", numbers[target_index])

        else:
            print("Target index not found.")

    
