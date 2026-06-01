def is_palindrome(text):
    return text.lower() == text.lower()[::-1]
print(is_palindrome("Level"))
print(is_palindrome("Python"))