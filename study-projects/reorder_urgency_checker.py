parts = [
    {"name": "Bearing", "stock": 2, "minimum": 5},
    {"name": "Sensor", "stock": 5, "minimum": 5},
    {"name": "Motor", "stock": 0, "minimum": 6}
]


def get_reorder_quantity(stock, minimum):
    if stock >= minimum:
        return 0
    return minimum - stock



def get_urgency(reorder_quantity):
    if reorder_quantity == 0:
        return "No order"
    elif reorder_quantity >= 5:
        return "High"
    elif reorder_quantity >= 2:
        return "Medium"
    else:
        return "Low"


for part in parts:
    name = part["name"]
    stock = part["stock"]
    minimum = part["minimum"]
    reorder_quantity = get_reorder_quantity(stock, minimum)
    urgency = get_urgency(reorder_quantity)

    print(f"{name} - Order {reorder_quantity} - {urgency} ")
