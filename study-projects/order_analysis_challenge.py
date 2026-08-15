orders = [
    {"PO_number": "334-GF001", "supplier_name": "RS Components", "quantity": 10, "unit_price": 12.50},
    {"PO_number": "371-GF002", "supplier_name": "Gopfert", "quantity": 2, "unit_price": 450.00},
    {"PO_number": "362A-GF003", "supplier_name": "Festo", "quantity": 5, "unit_price": 38.40},
    {"PO_number": "321-GF004", "supplier_name": "Cromwell", "quantity": 8, "unit_price": 75.00},
]


def get_order_price(order):
    quantity = order["quantity"]
    unit_price = order["unit_price"]
    return quantity * unit_price


def get_value_level(price_per_order):
    if price_per_order >= 500:
        return "High value"
    elif price_per_order >= 200:
        return "Medium value"
    return "Standard value"


all_orders_value = 0
high_value_order_count = 0


for order in orders:
    supplier_name = order["supplier_name"]

    price_per_order = get_order_price(order)
    value_level = get_value_level(price_per_order)

    all_orders_value += price_per_order

    if value_level == "High value":
        high_value_order_count += 1

    print(f"{supplier_name} - £{price_per_order:.2f} - {value_level}")


average_order_value = all_orders_value / len(orders)


print(f"Value of all orders: £{all_orders_value:.2f}")
print(f"Average value per order: £{average_order_value:.2f}")
print(f"High value orders count: {high_value_order_count}")