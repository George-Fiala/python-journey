def safe_int(text):
    try:
        return int(text)
    except ValueError:
        return None
    
def analyze_numbers(numbers):
    total_sum = 0
    frequency = {}
    max_value = None
    min_value = None

    for num in numbers:
        total_sum += num

        if num in frequency:
            frequency[num] += 1

        else:
            frequency[num] = 1

        if max_value is None or num > max_value:
            max_value = num

        if min_value is None or num < min_value:
            min_value = num
    
    return total_sum, frequency, max_value, min_value

test_data = [3, 1, 3, -2, 1, 3]
my_sum, my_freq, my_max, my_min = analyze_numbers(test_data)
print(my_sum)
print(my_freq)
print(my_max)
print(my_min)
