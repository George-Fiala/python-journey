
my_list = []
total = 0.0
count = 0
smallest = None
largest = None

while True:
    text = input("Enter number or stop:")
    if text == "stop":
        break
    value = float(text)
    count += 1
    total += value
    my_list.append(value)
    if smallest is None or value < smallest:
        smallest = value
    if largest is None or value > largest:
        largest = value

if count == 0:
    print("No numbers added.")

else:
    average = total / count
    print(total, count, average)
    print(smallest)
    print(largest)
    print(my_list)
        