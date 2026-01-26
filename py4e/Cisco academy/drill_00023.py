try:
    book = {'a': 1, 'b': 2}
    print(book['c'])

except KeyError:
    print("Key not there")