def get_unique(items):
    unique_item = []
    for i in items:
        if i not in unique_item:
            unique_item.append(i)
    return unique_item