def safe_int(text):
    """Safely converts string to integer."""
    try:
        return int(text)
    except ValueError:
        return None

def analyze_numbers(numbers):
    """Calculates sum, frequency, max, and min."""
    total_sum = 0
    frequency = {}
    
    for num in numbers:
        # 1. Sum calculation
        total_sum += num
        
        # 2. Dictionary Logic (Frequency counting)
        if num in frequency:
            frequency[num] += 1
        else:
            frequency[num] = 1
            
    # Return 4 values now!
    return total_sum, frequency, max(numbers), min(numbers)

def main():
    data = []
    
    # --- Part 1: Data Collection ---
    while True:
        user_input = input("Enter a number (or 'stop'): ")
        
        if user_input.lower() == 'stop':
            break
            
        clean_num = safe_int(user_input)
        
        if clean_num is not None:
            data.append(clean_num)
        else:
            print("Error: Please enter a valid integer.")

    # --- Part 2: Analysis & Output ---
    if data:
        # Unpacking 4 values from the function
        my_sum, my_freq, my_max, my_min = analyze_numbers(data)
        
        print("-" * 30)
        print(f"Data: {data}")
        print(f"Sum: {my_sum}")
        print(f"Max: {my_max}")
        print(f"Min: {my_min}")
        print(f"Frequency: {my_freq}")
        
        # --- Part 3: Search (Index Finder) ---
        print("-" * 30)
        search_input = input("Enter a number to find its index: ")
        search_num = safe_int(search_input)
        
        # Safety check: Is the number in our list?
        if search_num in data:
            # .index() finds the FIRST occurrence
            found_index = data.index(search_num)
            print(f"Number {search_num} is at index: {found_index}")
        else:
            print(f"Number {search_num} is NOT in the list.")
            
    else:
        print("No data provided.")

if __name__ == "__main__":
    main()