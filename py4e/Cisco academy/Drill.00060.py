text = "hello beatiful word"
words = text.split()

result = []
for word in words:
    result.append(word.capitalize())

final = " ".join(result)
print(final)