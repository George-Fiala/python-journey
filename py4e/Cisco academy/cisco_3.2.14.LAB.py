blocks = int(input("Enter the number of blocks: "))
height = 0
layer = 1
remaining = blocks

While remaining >= layer:
    remaining -= layer
    height += 1
    layer += 1

print("The height of the pyramid:", height)

