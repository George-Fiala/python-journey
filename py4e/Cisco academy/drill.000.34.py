def safe_index(numbers, target):
    for i in range(len(numbers)):
        if numbers[i] == target:
            return i
    return None

