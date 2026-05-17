def count_digits(text):
    count = 0
    for char in text:
        if char.isdigit():
            count += 1
    return count
