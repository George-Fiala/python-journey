numbers = []
even_numbers = []
unique_numbers = []
total = 0
min_value = None
max_value = None
count = 0
found = False
target_index = None

while True:
    text = input("Enter number or stop: ")
    if text == "stop":
        break

    value = int(text)
    total += value
    count += 1
    numbers.append(value)
    if value not in unique_numbers:
            unique_numbers.append(value)      
    if min_value is None or value < min_value:
        min_value = value
    if max_value is None or value > max_value:
        max_value = value


if count == 0:
    print("No numbers entered.")

else:
    for n in numbers: 
        if n % 2 == 0:
            even_numbers.append(n)
    
    target_text = input("Enter number that we looking for: ")
    target = int(target_text)
    for i in range(len(numbers)):
        if numbers[i] == target:
            found = True
            target_index = i
            break
    if found:
        print("Target index:", target_index)
        print("Value:", target)

    else:
        print("Not found")

    average = total / count
    print("You entered:", numbers)
    print("Count:", count)
    print("Average:", round(average, 2))
    print("Min value:", min_value)
    print("Max value:", max_value)
    print("Total:", total)
    print("Even numbers:", even_numbers)
    print("Unique numbers:", unique_numbers)