def word_count(sentence):
    words = sentence.split()
    count = 0
    for word in words:
        count += 1
    return count


def word_count(sentence):
    return len(sentence.split())