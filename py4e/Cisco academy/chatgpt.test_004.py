temps = [12, 15, 18, 9, 0, -2, 7, 14]
total = 0
max_value = None
min_value = None
max_index = None
min_index = None
count_above_zero = 0
count_below_zero = 0
count_zero = 0
index = 0

for t in temps:
    total += t
    if max_value is None or t > max_value:
        max_value = t
        max_index = index
    if min_value is None or t < min_value:
        min_value = t
        min_index = index
    index += 1
    if t > 0:
        count_above_zero += 1
    elif t < 0:
        count_below_zero += 1
    else:
        count_zero += 1

average_all = total / len(temps)

print("Total:", total)
print("Max value:", max_value)
print("Min value:", min_value)
print("Count above zero:", count_above_zero)
print("Count below zero:", count_below_zero)
print("Count zero:", count_zero)
print("Average all:", average_all)
print("Temps count:", len(temps))
print("Max index:", max_index)
print("Min index:", min_index)