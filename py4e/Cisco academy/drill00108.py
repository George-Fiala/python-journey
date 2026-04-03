def get_unique(items):
    unique_items =[]
    for char in items:
        if char not in unique_items:
            unique_items.append(char)
    return unique_items