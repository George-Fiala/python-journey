def safe_int(text):
    try:
        return int(text)
    except ValueError:
        return None
    
def analyze_numbers(numbers):
    total_sum = 0
    frequency = {}

    for num in numbers:
        total_sum += num

        if num in frequency:
            frequency[num] += 1
        else:
            frequency[num] = 1

    return total_sum, frequency, max(numbers), min(numbers)

def main():
    data = []

    while True:
        user_input = input("Enter a number (or 'stop'): ")

        if user_input.lower() == 'stop':
            break

        clean_num = safe_int(user_input)
        if clean_num is not None:
            data.append(clean_num)
        else:
            print("Error: Enter a proper integrer.")

    if data:
        my_sum, my_freq, my_max, my_min = analyze_numbers(data)

        print("-" * 30)
        print(f"Sum: {my_sum}")
        print(f"Frequency: {my_freq}")
        print(f"Max: {my_max}")
        print(f"Min: {my_min}")

        print("-" * 30)
        search_input = input("Enter a number to find its index: ")
        search_num = safe_int(search_input)

        if search_num in data:
            found_index = data.index(search_num)
            print(f"Number {search_num} is at index: {found_index}")
        else:
            print(f"Number {search_num} is NOT in the list.")

    else:
        print("No data provided")

if __name__ == "__main__":
    main()
    