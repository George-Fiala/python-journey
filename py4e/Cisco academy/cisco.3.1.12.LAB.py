blocks = int(input("Enter blocks: "))

h = 0
layer = 1
remaining = blocks

while remaining >= layer:
    remaining -= layer
    h += 1
    layer += (h ** 2)
    
print("height:", h)
print("Remaining blocks: ", remaining)
