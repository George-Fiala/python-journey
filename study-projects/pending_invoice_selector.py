invoices = [
    {"supplier": "NORD", "order_invoice": "INV-101", "amount": 420, "approved": False},
    {"supplier": "ERIKS", "order_invoice": "INV-102", "amount": 900, "approved": True},
    {"supplier": "Mosca", "order_invoice": "INV-103", "amount": 680, "approved": False},
    {"supplier": "Rubix", "order_invoice": "INV-104", "amount": 620, "approved": False}
]

def find_highest_pending_invoice(invoices):
    max_not_approved = None
    for invoice in invoices:
        approved = invoice["approved"]
        amount = invoice["amount"]
        if not approved:
            if max_not_approved is None or max_not_approved["amount"] < amount:
                max_not_approved = invoice
    return max_not_approved


highest_invoice = find_highest_pending_invoice(invoices)


if highest_invoice is not None:
    supplier = highest_invoice["supplier"]
    order_invoice = highest_invoice["order_invoice"]
    amount = highest_invoice["amount"]
    print(f"{supplier} - {order_invoice} - £{amount:.2f}")
else:
    print("No pending invoice")
