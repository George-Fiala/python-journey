def total_spend_by_supplier(orders):
    result = {}
    for order in orders:
        supplier = order["supplier"]
        amount = order["amount"]
        if supplier not in result:
            result[supplier] = amount
        else:
            result[supplier] += amount
    return result

print(total_spend_by_supplier (orders = [
    {"supplier": "RS", "amount": 120},
    {"supplier": "Amazon", "amount": 80},
    {"supplier": "RS", "amount": 200},
    {"supplier": "Brabazon", "amount": 50}
]))