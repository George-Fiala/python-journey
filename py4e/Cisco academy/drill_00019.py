def binarni_hledani(seznam, cil):
    low = 0
    high = len(seznam) - 1

    while low <= high:
        mid = (low + high) // 2
        odhad = seznam[mid]

        if odhad == cil:
            return mid
    
        elif odhad > cil:
            high = mid -1 
        else:
            low = mid + 1
    return None

muj_seznam = [2, 5, 8, 12, 16, 23, 38, 56, 72, 91]
print(binarni_hledani(muj_seznam, 23))
