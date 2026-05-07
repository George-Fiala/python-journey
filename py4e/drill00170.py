def word_count(text):
    result = {}
    words = text.split()
    for word in words:
        key = word
        if key not in result:
            result[key] = 1
        else:
            result[key] += 1
    return result