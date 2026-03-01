def remove_duplicates(numbers):
    if not numbers:
        return []
    unique_numbers = []
    for num in numbers:
        if num not in unique_numbers:
            unique_numbers.append(num)

    return unique_numbers