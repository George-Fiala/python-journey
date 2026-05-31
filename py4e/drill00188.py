def count_digits(text):
    count = 0
    for char in text:
        if char.isdigit():
            count += 1
    return count

print (count_digits('ter5dha7as8df9'))