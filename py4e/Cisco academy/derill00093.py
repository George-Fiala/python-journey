def get_unique(items):
    uni_item = []
    for char in items:
        if char not in uni_item:
            uni_item.append(char)
    return uni_item

    
