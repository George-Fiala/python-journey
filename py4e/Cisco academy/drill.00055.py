def flatten_list(nested):
    if not nested:
        return []
    
    result = []
    for sublist in nested:
        for num in sublist:
            result.append(num)
    return result

        

