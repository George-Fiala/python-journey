parts = [
    {"name": "Bearing", "stock": 8, "minimum": 6, "unit_price": 20, "critical": True},
    {"name": "Sensor", "stock": 3, "minimum": 5, "unit_price": 25, "critical": True},
    {"name": "Belt", "stock": 1, "minimum": 4, "unit_price": 10, "critical": False}
    
]


def get_reorder_quantity(stock, minimum):
    if stock >= minimum:
        return 0
    return minimum - stock

def get_part_priority(critical):
    if critical:
        return "Priority"
    return "Normal"

total_cost = 0


for part in parts:
    name = part["name"]
    stock = part["stock"]
    minimum = part["minimum"]
    unit_price = part["unit_price"]
    critical = part["critical"]
    part_priority = get_part_priority(critical)
    reorder_quantity = get_reorder_quantity(stock, minimum)
    if reorder_quantity > 0:
        line_cost = reorder_quantity * unit_price
        total_cost+= line_cost
        print(f"{name} - Order {reorder_quantity} - £{line_cost:.2f} - {part_priority}")
print(f"Total cost of order: £{total_cost:.2f}")



