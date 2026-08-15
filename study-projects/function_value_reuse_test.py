part = {
    "quantity": 5, "unit_price": 60
}

def get_full_price(part):
    quantity = part["quantity"]
    unit_price = part["unit_price"]
    return quantity * unit_price


def get_part_value(full_price):
    if full_price >= 250:
        return "High"
    return "Low"


full_price = get_full_price(part)
part_value = get_part_value(full_price)
print(f"£{full_price:.2f} - {part_value}")