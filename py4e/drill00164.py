def capitalize_words(text):
    words = text.split()
    result = []
    for word in words:
        result.append(word.capitalize())
    return " ".join(result)