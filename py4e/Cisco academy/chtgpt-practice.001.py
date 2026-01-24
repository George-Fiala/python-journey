secret_number = 42
count = 0
number = int(input("Enter secret number: "))


while number !=42:
    if number == 0:
        print("Vzdal ses, konec hry.")
        break
    elif number < 100:
        print("Číslo mimo rozsah, zkus to znovu.")
        continue
    elif number > 100:
        print("Číslo mimo rozsah, zkus to znovu.")
        continue
    else:
        if number < 42:
            print("Příliš malé.")
            count += 1
            break
        elif number > 42:
            print("Příliš velké.")
            count += 1
            break
            
        else:
            print("Správně, vyhrál jsi!")
            break
        