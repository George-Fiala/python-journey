largest = None
while True:
    entry = input("Number or 'done': ")
    if entry == 'done':
        break
    try:
        num = int(entry)
        if largest is None or num > largest:
            largest = num
    except:
        print("Bad number-skip!")
if largest is not None:
    print(f"Largest: {largest}")
else:
    print("No numbers entered!")