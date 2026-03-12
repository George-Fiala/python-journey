def get_longest_word(sentence):
    words = sentence.split()
    result = None
    for word in words:
        if result is None or len(word) > len(result):
            result = word
    return result