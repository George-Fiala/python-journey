parts = [
    {"name": "Bearing", "stock": 8, "minimum_stock": 6, "unit_price": 20},
    {"name": "Sensor", "stock": 3, "minimum_stock": 5, "unit_price": 25},
    {"name": "Belt", "stock": 1, "minimum_stock": 4, "unit_price": 10}
]


def get_reorder_quantity(stock, minimum_stock):
    if stock >= minimum_stock:
        return 0
    return minimum_stock - stock



total_cost = 0


for part in parts:
    name = part["name"]
    stock=part["stock"]
    minimum_stock = part["minimum_stock"]
    reorder_quantity = get_reorder_quantity(stock, minimum_stock)
    unit_price = part["unit_price"]
    line_cost = reorder_quantity * unit_price
    total_cost += line_cost
    if reorder_quantity > 0:
        print(f"{name} - Order {reorder_quantity} - Cost £{line_cost:.2f}")
print(f"Total cost: £{total_cost:.2f}")