def find_max(numbers):
    if not numbers:
        return None
    
    max_num = None

    for num in numbers:
        if max_num is None or num > max_num:
            max_num = num

    return max_num

