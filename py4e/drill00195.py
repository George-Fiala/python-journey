def count_consonants(text):
    count = 0
    vowels = 'aeiouAEIOU'
    for char in text.lower():
        if char.isalpha() and char not in vowels:
            count += 1
    return count
print(count_consonants("Procurement Automation"))