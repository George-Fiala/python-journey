try:
    value = int(input('Enter a natural number: '))
    print('The reciprocal of', value, 'is', 1/value)

    seznam = [] 
    print(seznam[0])

except ValueError:
    print('I do not know what to do.')
except ZeroDivisionError:
    print('Division by zero is not allowed in our Universe.')
except:
    print("Something strange hapebed here... Sorry!")


