invoices = [
    {"invo": "INV-201", "supplier": "NORD", "approved": False},
    {"invo": "INV-202", "supplier": "ERIKS", "approved": False},
    {"invo": "INV-203", "supplier": "NORD", "approved": True},
    {"invo": "INV-204", "supplier": "NORD", "approved": False},
    {"invo": "INV-205", "supplier": "Rubix", "approved": True},
    {"invo": "INV-206", "supplier": "ERIKS", "approved": False},
    {"invo": "INV-207", "supplier": "Mosca", "approved": False}
]

def count_unapproved_by_supplier(invoices):
    unapproved = {}
    for invoice in invoices:
        supplier = invoice["supplier"]
        approved = invoice["approved"]
        if not approved:
            unapproved[supplier] = unapproved.get(supplier, 0) + 1
    return unapproved


unapproved_invoices = count_unapproved_by_supplier(invoices)

if unapproved_invoices:
    
    print(unapproved_invoices) 
else:
    print("No unapproved invoices")