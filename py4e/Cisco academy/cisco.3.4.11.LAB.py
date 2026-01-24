beatles = []
print("Step 1:", beatles)
beatles.append("John Lennon")
beatles.append("Paul McCartney")
beatles.append("George Harrison")
print("Step 2:", beatles)
for i in range(2):
    new_member = input("Enter new band member: ")
    beatles.append(new_member)
print("Step 3:", beatles)
del beatles[-1]
del beatles[-1]
print("Step 4:", beatles)
beatles.insert(0, "Ringo Starr")
print("Step 5:", beatles)