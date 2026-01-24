def computepay(h, r):
        if h <= 40:
            return(h * r)
        elif h > 40:
            return((40 * r) + ((h - 40) * (r * 1.5)))
hrs = input("Enter Hours:")
rate = input("Enter rate:")
try:
    h = float(hrs)
    r = float(rate)
except:
    print("Please enter numbers")
    quit()
p = computepay(h, r)
print("Pay", p)