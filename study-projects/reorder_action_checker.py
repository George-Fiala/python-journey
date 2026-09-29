parts = [
    {"name": "Bearing", "stock": 2, "minimum": 5, "critical": True},
    {"name": "Sensor", "stock": 4, "minimum": 5, "critical": False},
    {"name": "Motor", "stock": 0, "minimum": 6, "critical": False},
    {"name": "Belt", "stock": 5, "minimum": 5, "critical": True},
]


def get_reorder_quantity(stock, minimum):
    if stock >= minimum:
        return 0
    return minimum - stock


def get_reorder_urgency(reorder_quantity):
    if reorder_quantity == 0:
        return "No order"
    elif reorder_quantity >= 5:
        return "High"
    elif reorder_quantity >=2:
        return "Medium"
    else:
        return "Low"


def get_reorder_action(reorder_quantity, reorder_urgency, critical):
    if reorder_quantity == 0:
        return "No action"
    elif reorder_urgency == "High" or critical:
        return "Immediate"
    elif reorder_urgency == "Medium":
        return "Plan"
    else:
        return "Routine"


for part in parts:
    name = part["name"]
    stock = part["stock"]
    minimum = part["minimum"]
    critical = part["critical"]
    reorder_quantity = get_reorder_quantity(stock, minimum)
    reorder_urgency = get_reorder_urgency(reorder_quantity)
    reorder_action = get_reorder_action(reorder_quantity, reorder_urgency, critical)

    print(f"{name} - Order: {reorder_quantity} - {reorder_urgency} - {reorder_action}")