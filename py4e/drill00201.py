def only_unique_items(items):
    results = []
    for item in items:
        if item not in results:
            results.append(item)
    return results
print(only_unique_items(items = ["Bobst", "Emba", "Bobst", "Asahi"]))