numbers = [42, 18, 76, 5, 33, 12]


def find_min_number(numbers):
    min_num = None
    for number in numbers:
        if min_num is None or min_num > number:
            min_num = number
        
    return min_num
       
def find_max_number(numbers):
    max_num = None
    for number in numbers:
        if max_num is None or max_num < number:
            max_num = number
        
    return max_num

print(find_min_number(numbers))
print(find_max_number(numbers))
print(find_max_number([]))
print(find_min_number([]))