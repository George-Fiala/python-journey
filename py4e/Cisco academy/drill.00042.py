def safe_int(text):
    try:
        return int(text)
    except ValueError:
        return None
def find_all_indexes(numbers, target):
    indexes = []

    for i in range(len(numbers)):
        if numbers[i] == target:
            indexes.append(i)
    return indexes

def analyze_numbers(numbers):
    total_sum = 0
    count = 0
    frequency = {}
    max_value = None
    min_value = None

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

    average = total_sum / count if count > 0 else None
    return total_sum, count, average, frequency, max_value, min_value

def print_frequency_line_by_line(freq):
    for key, value in freq.items():
        print(key, ":", value) 

def main():
    data = []

    while True:
        user_input = input("Enter a number or stop: ").strip()
        if user_input.lower() == "stop":
            break
        clean_num = safe_int(user_input)

        if clean_num is not None:
            data.append(clean_num)
        else:
            print("Error: Enter a valid integer.")
            continue
    if not data:
        print("No data provided.")
        return
    total_sum, count, average, freq, max_value, min_value = analyze_numbers(data)

    print("-" * 30)
    print(f"Data: {data}")
    print(f"Total sum: {total_sum}")
    print(f"Count: {count}")
    print(f"Average: {average}")
    print(f"Max num: {max_value}")
    print(f"Min num: {min_value}")
    print("-" * 30)
    print("Frequency: ")
    print_frequency_line_by_line(freq)
    print("-" * 30)

    search_index = input("Enter a number to find its index or stop: ").strip()
    if search_index.lower() == "stop":
        return
    search_num = safe_int(search_index)
    if search_num is None:
        print("Error: Enter a valid integer.")
        return
    indexes = find_all_indexes(data, search_num)

    if not indexes:
        print(f"NUmber {search_num} is not on the list.")
        return
    
    for i in indexes:
        print(f"Indexes: {i}, Value: {data[i]}")
    print(f"All indexes: {indexes}")
    print(f"Count of matches: {len(indexes)}")

if __name__ == "__main__":
    main()




    