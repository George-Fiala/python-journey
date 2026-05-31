def get_unique(items):
    result = []
    for item in items:
        if item not in result:
            result.append(item)
    return result

print(get_unique(["EMBA", "BOBST", "EMBA", "RS", "BOBST"]))