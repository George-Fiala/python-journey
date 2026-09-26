parts = [
    {"name": "Bearing", "stock": 2, "minimum": 5, "unit_price": 40},
    {"name": "Sensor", "stock": 3, "minimum": 5, "unit_price": 25},
    {"name": "Belt", "stock": 1, "minimum": 4, "unit_price": 10},
]


def get_reorder_quantity(stock, minimum):
    if stock >= minimum:
        return 0
    return minimum - stock

total_cost = 0
budget = 180


for part in parts:
    name = part["name"]
    stock = part["stock"]
    minimum = part["minimum"]
    unit_price = part["unit_price"]
    reorder_quantity = get_reorder_quantity(stock, minimum)
    line_cost = reorder_quantity * unit_price
    total_cost += line_cost



print(f"Total cost = £{total_cost:.2f}")
if total_cost <= budget:
    print("Within budget")
else:
    print("Over budget")