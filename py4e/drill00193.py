def get_unique(words):
    
    result = []
    for word in words:
        if word not in result:
            result.append(word)
    return result
print(get_unique(["EMBA", "BOBST", "EMBA", "RS", "BOBST", "IFM"]))