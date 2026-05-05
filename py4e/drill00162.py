def group_by_first_letter(words):
    result = {}
    for word in words:
        if not word:
            continue
        key = word[0].lower()
        if key not in result:
            result[key] = []
        result[key].append(word)
    return result