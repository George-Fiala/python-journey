parts = [
    {"part_name": "Bearing 6204", "stock_quantity": 3, "minimum_stock": 5, "is_critical": True},
    {"part_name": "M8 Bolt", "stock_quantity": 40, "minimum_stock": 20, "is_critical": False},
    {"part_name": "Drive Belt", "stock_quantity": 2, "minimum_stock": 2, "is_critical": True},
]

def get_reorder_action(part):
    stock_quantity = part["stock_quantity"]
    minimum_stock = part["minimum_stock"]
    is_critical = part["is_critical"]
    if stock_quantity < minimum_stock and is_critical:
        return "Urgent reorder"
    elif stock_quantity < minimum_stock:
        return "Reorder"
    return "Stock OK"


def get_reorder_quantity(part):
    minimum_stock = part["minimum_stock"]
    stock_quantity = part["stock_quantity"]
    if stock_quantity < minimum_stock:
        return minimum_stock - stock_quantity
    return 0


for part in parts:
    reorder_action = get_reorder_action(part)
    reorder_quantity = get_reorder_quantity(part)
    part_name = part["part_name"]
    if reorder_quantity > 0:
        order_message = f"Order quantity: {reorder_quantity}"
    else:
        order_message = "No order required"
    print(f"{part_name} - {reorder_action} - {order_message}")