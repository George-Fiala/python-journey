parts = [
    {"name": "Bearing", "stock": 8, "minimum_stock": 6},
    {"name": "Sensor", "stock": 3, "minimum_stock": 5}
]


def get_reorder_quantity(stock, minimum_stock):
    if stock >= minimum_stock:
        return 0
        
    return minimum_stock - stock


for part in parts:
    name = part["name"]
    stock = part["stock"]
    minimum_stock = part["minimum_stock"]
    reorder_quantity = get_reorder_quantity(stock, minimum_stock)
    if reorder_quantity > 0:
        print(f"{name} - Reorder quantity: {reorder_quantity}")