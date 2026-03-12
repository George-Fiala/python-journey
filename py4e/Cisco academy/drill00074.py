def group_by_forst_letter(words):
    result = {}
    for word in words:
        key = word[0]
        if key not in result:
            result[key] = []
        result[key].append(word)
    return result