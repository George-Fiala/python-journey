c0 = int(input("Enter any non-negative and non-zero integer number: "))
steps = 0

while c0 != 1:
    if c0 % 2 == 1:
        c0 = c0 / 2
        steps += 1
    else:
        c0 = 3 * c0 + 1       
        steps += 1

    
print("Steps: ", steps)
