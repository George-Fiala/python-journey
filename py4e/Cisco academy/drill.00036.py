def safe_index(data, target):
    for i in range(len(data)):
        if data[i] == target:
            return i
    return None

data = [3, 1, -2, 3, 1, 3]

seach_num = 1

idx = safe_index(data, seach_num)

if idx is None:
    print("Idx is not in data")
else:
    print("Nasel jsem na indexu:", idx)


search_num = 999
idx = safe_index(data, search_num)

if idx is None:
    print("Číslo není v seznamu")
else:
    print("Našel jsem na indexu:", idx)