def find_max(numbers):
    max_num = None
    for number in numbers:
        if max_num is None or number > max_num:
            max_num = number
    return max_num
print(find_max([4, -2, 99, 15, 0]))