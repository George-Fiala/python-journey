def safe_index(data, target):
    for i in range(len(data)):
        if data[i] == target:
            return i

    return None

data = [3, 1, -2, 3, 1, 3]
print(safe_index(data, 1))     # má vypsat 1 (první výskyt čísla 1)
print(safe_index(data, 999))   # má vypsat None