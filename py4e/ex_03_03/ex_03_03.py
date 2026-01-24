score = input("Enter Score: ")
sr = float(score)
if sr > 1:
    print("Error")
if sr < 0:
    print("Error")
elif sr >= 0.9:
    print("A")
elif sr >= 0.8:
    print("B")
elif sr >= 0.7:
    print("C")
elif sr >= 0.6:
    print("D")
elif sr < 0.6:
    print("F")