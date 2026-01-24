numbers = [10, 5, 7, 2, 1]
print("Original list contents:", numbers)

numbers[0] = 111
print("New list of contents:", numbers)

numbers[1] = numbers[4]
print("New list contents:", numbers)

numbers = [10, 5, 7, 2, 1]
print("Original listy contents:", numbers)

numbers[0] = 111
print("\nPrevious list contents:", numbers)

numbers[1] = numbers[4]
print("Previous listy contents:", numbers)

print("\nList length:", len(numbers))

print(numbers)

del numbers[1]
print(len(numbers))
print(numbers)

print(numbers[4])
numbers[4] = 1