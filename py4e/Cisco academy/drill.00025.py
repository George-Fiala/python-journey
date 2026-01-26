while True:
    first_number = input("Enter number or q to end: ")
    if first_number == "q":
        print("ending...")
        break
    else:
        first_number = float(first_number)

    operation = input("+, -, / , * or q to end: ")
    if operation == "q":
        print("Ending...")
        break
    
    second_number = input("Enter number or q to end: ")
    if second_number == "q":
        print("Ending")
        break
    else:
        second_number = float(second_number)
                              
   
    if operation == "+":
        print("Adding", first_number + second_number)
    
    elif operation == "-":
        print("Minus", first_number - second_number)
    
    elif operation == "/":
        print("Divide", first_number / second_number)
    
    elif operation == "*":
        print("Multiply", first_number * second_number)
    
    else:
        print("Nothing to calculate.")        


