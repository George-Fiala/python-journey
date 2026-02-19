numbers = [1, 2, 3, 4, 5]

def calculate_average(numbers):
    if not numbers:
        return None
    
    total_sum = 0
    count = 0

    for num in numbers:
        total_sum += num
        count += 1

    return total_sum / count
    
print(calculate_average(numbers))
