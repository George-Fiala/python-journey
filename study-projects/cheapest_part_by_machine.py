
parts = [
    {"part": "Bearing", "machine": "Emba 2", "price": 180, "available": True},
    {"part": "Sensor", "machine": "Bobst 1", "price": 450, "available": True},
    {"part": "Belt", "machine": "Emba 2", "price": 95, "available": True},
    {"part": "Motor", "machine": "Bobst 1", "price": 700, "available": False},
    {"part": "Valve", "machine": "Emba 2", "price": 240, "available": True},
    {"part": "Roller", "machine": "Mosca 5", "price": 120, "available": True}
]


def find_cheapest_by_machine(parts):
    cheapest_parts = {}

    for part in parts:
        machine = part["machine"]
        price = part["price"]
        available = part["available"]

        if available:
            if machine not in cheapest_parts:
                cheapest_parts[machine] = part
            elif price < cheapest_parts[machine]["price"]:
                cheapest_parts[machine] = part

    return cheapest_parts


cheapest = find_cheapest_by_machine(parts)

if cheapest:
    for machine in cheapest:
        details = cheapest[machine]
        print(
            f"{machine}: Part: {details['part']}, "
            f"Machine: {details['machine']}, "
            f"Price: £{details['price']:.2f}, "
            f"Available: {details['available']}"
        )
else:
    print("No parts")
