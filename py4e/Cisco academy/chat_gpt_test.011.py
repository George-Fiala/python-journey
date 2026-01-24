numbers = [10, 20, 10, 30, 20, 40, 30]
unique_numbers = []

total = 0
count = 0
min_value = None
max_value = None

for n in numbers:
    total += n
    count += 1
    if min_value is None or n < min_value:
        min_value = n
    if max_value is None or n > max_value:
        max_value = n
    if n not in unique_numbers:
        unique_numbers.append(n)

if count == 0:
    print("No numbers.")

else:
    average = total / count

    print(f"Maximal value = {max_value}")
    print(f"Minimal value = {min_value}")
    print(f"Total: {total}")
    print(f"Count: {count}")
    print(f"Average: {average}")
    print(f"Unique numbers: {unique_numbers}")
    print(f"Numbers:{numbers}")