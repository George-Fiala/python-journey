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

    return total_sum, frequency

def main():
    data = []

    while True:
        user_input = input("Enter a number (or 'stop' to finish): ")

        if user_input.lower() == 'stop':
            break

        clean_num = safe_int(user_input)

        if clean_num is not None:
            data.append(clean_num)

        else:
            print("Error: Please enter a valid integer.")

    if data:
        my_sum, my_freq = analyze_numbers(data)

        print("-" * 30)
        print(f"Data: {data}")
        print(f"Total Sum: {my_sum}")
        print(f"Frequency: {my_freq}")

    else:
        print("No data provided.")

if __name__ == "__main__":
    main()
