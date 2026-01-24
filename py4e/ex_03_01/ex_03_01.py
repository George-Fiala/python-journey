hrs = input("Enter Hours:")
rate = input("Enter rate")
h = float(hrs)
rt = float(rate)
if h <= 40:
    print(h * rt)
else:
    print((h - 40) * (rt * 1.5) + (40 * rt))
    