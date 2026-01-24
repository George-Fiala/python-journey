count_pos = 0
total = 0

while True:
    number = int(input("Enter number: "))
    if number < 0:
        break
    if number > 0:
        count_pos += 1
        total += number

print("Total of positive numbers:", total)
print("Number of positive numbers:", count_pos)