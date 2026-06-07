def find_max_min_and_avg_numbers(cost):
    max_num = None
    min_num = None
    
    for number in cost:
        if max_num is None or number > max_num:
            max_num = number
        if min_num is None or number < min_num:
            min_num = number
    avg_num = sum(cost) / len(cost)
    return max_num, min_num, avg_num 

print(find_max_min_and_avg_numbers([120, 350, 99, 450, 220]))