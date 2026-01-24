num = int(input("Enter number: "))
if num < 0:
    if num % 2 == 0:
        print("Negative even")
    else:
        print("Negative odd")
elif num == 0:
    print("Zero even")
else:
    if num % 2 == 0:
        print("Positive even")
    else:
        print("Positive odd")
