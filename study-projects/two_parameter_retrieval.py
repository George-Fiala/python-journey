parts = [
    {"part_name": "Bearing", "price": 400, "critical": True},
    {"part_name": "Belt", "price": 150, "critical": False},
    {"part_name": "Sensor", "price": 650, "critical": True}
]


def get_part_priority(price, critical):
    if price >= 500 and critical:
        return "Priority"
    return "Normal"


for part in parts:
    part_name = part["part_name"]
    price = part["price"]
    critical = part["critical"]
    part_priority = get_part_priority(price, critical)
    print(f"{part_name} - £{price:.2f} - {part_priority}")