parts = [
    {"name": "Emba Bearing", "stock": 3, "unit_price": 70, "critical": True},
    {"name": "Bobst Sensor", "stock": 1, "unit_price": 260, "critical": True},
    {"name": "Gopfert Belt", "stock": 1, "unit_price":190, "critical": True},
    {"name": "Emba Workshop Bolts", "stock": 0, "unit_price": 3, "critical": False}
]

def find_priority_part(parts):
    minimal_stock = None
    for part in parts:
        critical = part["critical"]
        stock = part["stock"]
        unit_price = part["unit_price"]
        if critical:
            if minimal_stock is None or minimal_stock["stock"] > stock:
                minimal_stock = part
            elif minimal_stock["stock"] == stock and minimal_stock["unit_price"] < unit_price:
                minimal_stock = part
    return minimal_stock


priority_part = find_priority_part(parts)


if priority_part is not None:
    name = priority_part["name"]
    print(name)

else:
    print("No priority part")
            


