def get_unique(text):
    results = []
    for char in text:
        if char not in results:
            results.append(char)
    return results

print (get_unique([1, 2, 2, 3, 1, 4]))