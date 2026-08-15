parts = [
    {"part_name": "Bearing", "quantity": 3, "unit_price": 120},
    {"part_name": "Belt", "quantity": 2, "unit_price": 45},
    {"part_name": "Sensor", "quantity": 1, "unit_price": 310}
]


def get_full_price(part):
    quantity = part["quantity"]
    unit_price = part["unit_price"]
    return quantity * unit_price


def get_part_value(full_price):
    if full_price >= 300:
        return "High value"
    return "Standard"


high_value_count = 0
all_parts_added_price = 0
total_parts = len(parts)


for part in parts:
    full_price = get_full_price(part)
    part_value = get_part_value(full_price)
    part_name = part["part_name"]
    all_parts_added_price += full_price
    if part_value == "High value":
        high_value_count += 1
    print(f"{part_name} - £{full_price:.2f} - {part_value}")
print(f"Complete price: £{all_parts_added_price:.2f}")
print(f"High value part count: {high_value_count}")

average_part_value = all_parts_added_price / len(parts)
print(f"Average part price: £{average_part_value:.2f}")

high_value_percentage = high_value_count / total_parts * 100
print(f"High value percentage: {high_value_percentage:.2f}%")