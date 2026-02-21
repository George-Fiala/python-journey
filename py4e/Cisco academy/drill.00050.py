def sum_positive(numbers):
    if not numbers:
        return 0
    total = 0

    for num in numbers:
        if num > 0:
            total += num
    return total
        