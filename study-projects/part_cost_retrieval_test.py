parts = [
    {"part_name": "Bearing", "quantity": 4, "unit_price": 80},
    {"part_name": "Belt", "quantity": 2, "unit_price": 60}
]


def get_full_price(part):
    quantity = part["quantity"]
    unit_price = part["unit_price"]
    return quantity * unit_price


def get_part_value(full_price):
    if full_price >= 250:
        return "Expensive"
    return "Standard"


expensive_parts_count = 0

for part in parts:
    full_price = get_full_price(part)
    part_value = get_part_value(full_price)
    part_name = part["part_name"]
    print(f"{part_name} - £{full_price:.2f} - {part_value}")
    if part_value == "Expensive":
        expensive_parts_count += 1
print(f"Expensive parts count : {expensive_parts_count}")