numbers = []
unique_numbers = []



while True:
    text = input("Enter number or stop:")
    if text == "stop":
        break

    value = float(text)
    numbers.append(value)
    if value not in unique_numbers:
        unique_numbers.append(value)

if not unique_numbers:
    print("No unique numbers")

else:
    print(f"Unique numbers: {unique_numbers}")
    print(f"All inputs: {numbers}")
                  