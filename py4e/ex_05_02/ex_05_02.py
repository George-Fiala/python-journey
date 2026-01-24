largest = None
smallest = None
had_invalid = False
while True:
    num = input('Enter a number: ')
    if num == 'done':
        break
    #print(num)
    try:
        num = (int(num))
        if largest is None or num > largest:
            largest = num
        if smallest is None or num < smallest:
            smallest = num
    except:
        had_invalid = True
        continue
            
if had_invalid: print('Invalid input')            
print('Maximum is', largest)
print('Minimum is', smallest)
