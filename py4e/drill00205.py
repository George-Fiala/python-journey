def filter_big_orders(orders):
    result = []
    for order in orders:
        amount = order["amount"]
        if amount > 100:
           result.append(order)

            
    return result

print(filter_big_orders([
    {"supplier": "RS", "amount": 120},
    {"supplier": "RS", "amount": 200}
]))