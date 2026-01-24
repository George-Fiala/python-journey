while True : 
    guess = input(f"Guess: ")
    try:
        num = int(guess)
        if num > 10:
            print("Too high!")
        elif num <= 0:
            print("Too low!")
        else:
            print("Spot on-WIN!")
            break
    except:
        print("Not a number-guess again")
        continue
print("Game over-smart guesses!")
        
    
    