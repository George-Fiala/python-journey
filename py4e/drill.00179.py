def word_count(text):
    words = text.split()
    results = {}
    for word in words:
        if word not in results:
            results[word] = 1
        else:
            results[word] += 1
    return results
           

