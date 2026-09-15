part = {
    "name": "Bearing", "stock": 7, "minimum_stock": 6
}
part2 = {
    "name": "Sensor",
    "stock": 2,
    "minimum_stock": 5
}

def check_stock(stock, minimum_stock):
    if stock <= minimum_stock:
        return "Reorder"
    return "Stock OK"

stock = part["stock"]
minimum_stock = part["minimum_stock"]
stock2 = part2["stock"]
minimum_stock2 = part2["minimum_stock"]
status = check_stock(stock, minimum_stock)
status2 = check_stock(stock2, minimum_stock2)
print(status)
print(status2)