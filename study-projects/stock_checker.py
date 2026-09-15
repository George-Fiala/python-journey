part = {
    "name": "Bearing", "stock": 7, "minimum_stock": 6
}

def check_stock(stock, minimum_stock):
    if stock <= minimum_stock:
        return "Reorder"
    return "Stock OK"

stock = part["stock"]
minimum_stock = part["minimum_stock"]
status = check_stock(stock, minimum_stock)
print(status)