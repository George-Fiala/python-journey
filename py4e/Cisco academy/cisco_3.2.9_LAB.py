secret_word = "dragon"
attempts = 0

while True:
    word = input("Enter word: ")
    if word == secret_word:
        print("You found the dragon.")
        break
    attempts += 1

print("Attempts:", attempts)

