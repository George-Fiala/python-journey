def analyze_text(text):
    words = text.split()
    result = {}
    for word in words:
        if word not in result:
            result[word] = 1
        else:
            result[word] += 1
    return result
    
    