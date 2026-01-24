nums = [12, -4, 7, 0, 19, -10, 3, 0]

total = 0
max_value = None
min_value = None
count_positive = 0
count_negative = 0
count_zeroes = 0
sum_positive = 0
sum_negative = 0

for n in nums:
    total += n

    if max_value is None or n > max_value:
        max_value = n
    if min_value is None or n < min_value:
        min_value = n

    if n > 0:
        count_positive += 1
        sum_positive += n
    elif n < 0:
        count_negative += 1
        sum_negative += n
    else:
        count_zeroes += 1

# průměry až PO cyklu
average_all = total / len(nums)

if count_positive > 0:
    average_positive = sum_positive / count_positive
else:
    average_positive = None  # nebo 0, podle toho, jak se rozhodneš

if count_negative > 0:
    average_negative = sum_negative / count_negative
else:
    average_negative = None

print("Total:", total)
print("Max:", max_value)
print("Min:", min_value)
print("Positive count:", count_positive)
print("Positive sum:", sum_positive)
print("Negative count:", count_negative)
print("Negative sum:", sum_negative)
print("Zeroes:", count_zeroes)
print("Total items:", len(nums))
print("Average (all):", average_all)
print("Average positive:", average_positive)
print("Average negative:", average_negative)