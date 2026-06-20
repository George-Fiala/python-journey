orders = [
    {"po": "001", "supplier": "RS", "cost": "25.50"},
    {"po": "002", "supplier": "IFM", "cost": "80.00"},
    {"po": "003", "supplier": "RS", "cost": "12.99"}
]


search_supplier = input("Supplier: ")
print("Searching for:", search_supplier)

found = False
for order in orders:
    if order["supplier"] == search_supplier:
        print("PO:", order["po"], "| Supplier:", order["supplier"], "| Cost:", order["cost"])
        found = True

if found is False:
    print("No order found for supplier.")
