def capitalize_words(sentence):
    words = sentence.split()
    result = []
    for word in words:
        result.append(word.capitalize())
    return " ".join(result)



