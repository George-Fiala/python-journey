# 1. Vytvoříme seznam 10 čísel
my_list = [0, 10, 20, 30, 40, 50, 60, 70, 80, 90]
length = len(my_list)

print("Startovní stav:", my_list)
print("-" * 50)

# 2. Smyčka jede do poloviny (length // 2), tedy 5x
for i in range(length // 2):
    
    # Vypočítáme index na druhé straně (pravá ruka)
    opposite = length - i - 1
    
    # Uložíme si hodnoty jen pro výpis (abychom viděli, co měníme)
    val_left = my_list[i]
    val_right = my_list[opposite]
    
    # --- MAGICKÁ VÝMĚNA (SWAP) ---
    # Tady se to stane: Levá se stane Pravou a Pravá se stane Levou
    my_list[i], my_list[opposite] = my_list[opposite], my_list[i]
    
    # Výpis aktuálního kroku
    print(f"Krok {i}: Měním {val_left} (index {i}) <--> {val_right} (index {opposite})")
    print(f"Stav:   {my_list}")
    print("-")

print("-" * 50)
print("HOTOVO. Otočený seznam:", my_list)