def get_unique(items):
    result = []
    for char in items:
        if char not in result:
            result.append(char)
    return result