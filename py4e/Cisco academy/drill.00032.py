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
    count = 0

    for num in numbers:
        total_sum += num
        count += 1

        if num in frequency:
            frequency[num] += 1

        else:
            frequency[num] = 1
        
        if max_value is None or num > max_value:
            max_value = num
        if min_value is None or num < min_value:
            min_value = num
    
    avg = total_sum / count
    return total_sum, frequency, max_value, min_value, count, avg

def main():
    data = []

    while True:
        user_input = input("Enter a number (or 'stop'): ")
        if user_input.lower().strip() == 'stop':
            break

        clean_num = safe_int(user_input)
        if clean_num is not None:
            data.append(clean_num)
        else:
            print("Error: Enter valid integer.")

    if data:
        my_sum, my_freq, my_max, my_min, my_count, my_avg = analyze_numbers(data)

        print("-" * 30)
        print(f"Sum: {my_sum}")
        print("Frequency:")
        for key, value in my_freq.items():
            print(key, ":", value)
        print(f"Max: {my_max}")
        print(f"Min: {my_min}")
        print(f"Count: {my_count}")
        print(f"Average: {my_avg}")

        print("-" * 30)
        search_input = input("Enter a number to find its index: ")
        search_num = safe_int(search_input)
        if search_num is None:
            print("Error: Enter a valid integer to search.")
        elif search_num in data:
            found_index = data.index(search_num)
            print(f"Number {search_num} is at index: {found_index}")

        else:
             print(f"Number {search_num} is NOT in the list.")
    else:
        print("No data provided.")

if __name__ == "__main__":
    main()

