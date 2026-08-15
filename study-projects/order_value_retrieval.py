orders = [
    {"po_number": "PO001", "quantity": 5, "unit_price": 40},
    {"po_number": "PO002", "quantity": 2, "unit_price": 300},
    {"po_number": "PO003", "quantity": 10, "unit_price": 15},
    {"po_number": "PO004", "quantity": 1, "unit_price": 750},
 ]


def get_full_price_per_po(order):
    quantity = order["quantity"]
    unit_price = order["unit_price"]
    return quantity * unit_price


def get_value_per_po(full_price_per_po):
    if full_price_per_po >= 500:
        return "High value"
    return "Standard"




total_price = 0
order_count = len(orders)
high_value_po_count = 0



for order in orders:
    full_price_per_po = get_full_price_per_po(order)
    value_per_po = get_value_per_po(full_price_per_po)
    po_number = order["po_number"]
    total_price += full_price_per_po
    if value_per_po == "High value":
        high_value_po_count += 1
    print(f"{po_number} - £{full_price_per_po:.2f} - {value_per_po}")


average_order_value = total_price / order_count
percentage_of_high_value = high_value_po_count / order_count * 100
print(f"Total value of all orders: £{total_price:.2f}")
print(f"Average order value: £{average_order_value:.2f}")
print(f"High value orders count: {high_value_po_count}")
print(f"Percentage of high value orders: {percentage_of_high_value:.2f}%")