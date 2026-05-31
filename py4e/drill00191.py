def count_digits(items):
    count = 0
    for item in items:
        if item.isdigit():
            count += 1
    return count
print(count_digits("PO123-GF45-A9"))
