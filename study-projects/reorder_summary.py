parts = [
    {"name": "Bearing", "stock": 2, "minimum" : 5, "unit_price": 40, "critical": True},
    {"name": "Sensor", "stock": 3, "minimum" : 5, "unit_price": 25, "critical": True},
    {"name": "Belt", "stock": 6, "minimum" : 4, "unit_price": 10, "critical": False},
    {"name": "Motor", "stock": 1, "minimum" : 2, "unit_price": 80, "critical": False},
]


def get_reorder_quantity(stock, minimum):
    if stock >= minimum:
        return 0
    return minimum - stock


total_cost = 0
order_parts_count = 0
critical_parts_to_order = 0
all_parts_to_order_count = 0
all_critical_parts_to_order_count = 0


for part in parts:
    name = part["name"]
    stock = part["stock"]
    minimum = part["minimum"]
    unit_price = part["unit_price"]
    critical = part["critical"]
    reorder_quantity = get_reorder_quantity(stock, minimum)
    line_cost = reorder_quantity * unit_price
    total_cost += line_cost
    if reorder_quantity > 0:
        order_parts_count += 1
        all_parts_to_order_count += reorder_quantity
        if critical:
            critical_parts_to_order += 1
            all_critical_parts_to_order_count += reorder_quantity
   


        print(f"{name} - Order amount: {reorder_quantity} - Cost: £{line_cost:.2f}")
print(f"Total cost: £{total_cost:.2f}")
print(f"Critical part types to order : {critical_parts_to_order}")
print(f"Part types to order: {order_parts_count}")
print(f"Total units to order: {all_parts_to_order_count}")
print(f"Critical units to order: {all_critical_parts_to_order_count}")