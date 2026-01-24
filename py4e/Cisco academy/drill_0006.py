def safe_int(text):
    try:
        return int(text)
    except ValueError:
        return None

print(safe_int(" 7 "))
print(safe_int("abc"))
print(safe_int(""))
print(safe_int("-2"))