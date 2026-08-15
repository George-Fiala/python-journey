order = {
    "quantity": 3,"unit_price": 100
}


def get_order_price(order):
    quantity = order["quantity"]
    unit_price = order["unit_price"]
    return quantity * unit_price




def get_order_value(get_order_price):
    if get_order_price >= 500:
        return "High"
    elif get_order_price >= 200:
        return "Medium"
    return "Low"


order_price = get_order_price(order)
order_value = get_order_value(order_price)
print(f"£{order_price:.2f} - {order_value}")



