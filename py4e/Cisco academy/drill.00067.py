def count_vowels(text):
    count = 0
    vowels = "aeiouAEIOU"
    for letter in text:
        if letter in vowels:
            count += 1
    return count