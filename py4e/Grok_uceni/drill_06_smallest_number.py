smallest = None
while True:
    entry = input("Number or 'done': ")
    if entry == 'done':
        break
    try:
        num = int(entry)
        if smallest is None or num < smallest:
            smallest = num
    except:
        print("Bad number-skip!")
if smallest is not None:
    print(f"Smallest: {smallest}")
else:
    print("No numbers entered!")