def safe_int(text):
    try:
        return int(text)
    except ValueError:
        return None
    
def safe_index(numbers, target):
    for i in range(len(numbers)):
        if numbers[i] == target:
            return i
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

        
    avg = total_sum / count
    return total_sum, count, avg

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
        my_sum, my_count, my_avg = analyze_numbers(data)

        print("-" * 30)
        print(f"Sum: {my_sum}")
        print(f"Count: {my_count}")
        print(f"Average: {my_avg}")

        print("-" * 30)
        search_input = input("Enter a number to find its index: ")
        search_num = safe_int(search_input)
        if search_num is None:
            print("Error: Enter a valid integer to search.")
        else:
            idx = safe_index(data, search_num)   # ✅ PUT idx HERE (this line)

            if idx is None:
                print(f"Number {search_num} is NOT in the list.")
            else:
                print(f"Number {search_num} is at index: {idx}")
    else:
        print("No data provided.")

if __name__ == "__main__":
    main()