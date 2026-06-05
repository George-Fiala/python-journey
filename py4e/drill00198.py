def group_by_first_letter(words):
    
    results = {}
    for word in words:
        key = word[0]
        if key not in results:
            results[key] = []
        results[key].append(word)
    return results
print(group_by_first_letter(["Bobst", "Belt", "Emba", "Encoder", "RS"]))