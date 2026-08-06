orders = [
    {"po_number": "334-GF001", "supplier_name": "RS Components", "quantity": 10, "unit_price": 12.50},
    {"po_number": "371-GF002", "supplier_name": "Gopfert", "quantity": 2, "unit_price": 450},
    {"po_number": "362A-GF003", "supplier_name": "Festo", "quantity": 5, "unit_price": 38.40},
]


def get_order_value(order):
    quantity = order["quantity"]
    unit_price = order["unit_price"]
    return quantity * unit_price


def get_value_level(order_value):
    if order_value >= 500:
        return "High value"
    elif order_value >= 200:
        return "Medium value"
    return "Standard Value"

total_order_value = 0

for order in orders:
    order_value = get_order_value(order)
    value_level = get_value_level(order_value)
    po_number = order["po_number"]
    supplier_name = order["supplier_name"]
    total_order_value += order_value
    print(f"{po_number} - {supplier_name} - £{order_value:.2f} - {value_level}")
print(f"Total order value: £{total_order_value:.2f}")