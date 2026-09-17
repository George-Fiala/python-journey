part = {
    "name": "Sensor",
    "stock": 8,
    "minimum_stock": 5
}

def get_reorder_quantity(stock, minimum_stock):
    if stock >= minimum_stock:
        return 0
    return minimum_stock - stock

name = part["name"]
stock = part["stock"]
minimum_stock = part["minimum_stock"]
quantity_to_order = get_reorder_quantity(stock, minimum_stock)

print(f"{name} - Quantity to order: {quantity_to_order}")