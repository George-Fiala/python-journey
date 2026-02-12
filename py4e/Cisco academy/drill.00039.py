def safe_int(text):
    try:
        return int(text)
    except ValueError:
        return None
    
total_leaks = 0

for i in range(3):
    user_input = input(f"Enter number of leaks for machine {i+1}: ")
    total_num = safe_int(user_input)
    if total_num is not None:
        total_leaks += total_num

    else:
        print("Error: Enter valid integer")
print(f"Total leaks found today: {total_leaks}")
