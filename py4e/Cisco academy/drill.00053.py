def find_min(numbers):
    if not numbers:
        return None
    min_value = None
    for num in numbers:
        if min_value is None or num < min_value:
            min_value = num

    return min_value