def find_max(numbers):
    max_number = None
    for number in numbers:
        if max_number is None or number > max_number:
            max_number = number
    return max_number
