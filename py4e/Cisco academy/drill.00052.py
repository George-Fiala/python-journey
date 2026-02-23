def count_even(numbers):
    even_numbers = 0
    if not numbers:
        return 0
    for num in numbers:
        if num % 2 == 0:
            even_numbers += 1

    return even_numbers

    
