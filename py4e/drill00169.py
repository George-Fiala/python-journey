def count_consonants(text):
    count = 0
    vowels = 'aeiouAEIOU'
    for char in text:
        if char.isalpha() and char not in vowels:
            count += 1
    return count
