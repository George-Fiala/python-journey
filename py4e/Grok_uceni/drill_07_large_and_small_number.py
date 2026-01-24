largest = None
smallest = None
while True:
    entry = input("Number or 'done': ")
    if entry == 'done':
        break
    try:
        num = int(entry)
        if largest is None or num > largest:
            largest = num 
           
        if smallest is None or num < smallest:
            smallest = num
         
    except:
        print("Bad number-skip!")
if largest is None:          # no numbers were entered
    print("No numbers entered!")
else:
    print(f"Largest: {largest}")
    print(f"Smallest: {smallest}")