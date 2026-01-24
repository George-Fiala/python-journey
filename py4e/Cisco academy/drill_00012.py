
numbers = [
    [3, 8, 15, 22],
    [4, 9, 12, 18],
    [7, 6, 10, 25]
]

total = 0
even_count = 0
even_numbers = []
max_value = None
min_value = None

for row in numbers:
    for n in row:
        total += n
        if n % 2 == 0:
            even_count += 1
            even_numbers.append(n)
        if max_value is None or n > max_value:
            max_value = n
        if min_value is None or n < min_value:
            min_value = n        

print("Total:", total)
print("Even numbers count:", even_count)
print("Even numbers:", even_numbers)
print("Max number=", max_value)
print("Min number=", min_value)
