orders = [
    {"po": "001","supplier": "RS","cost": 25.50},
    {"po": "002","supplier": "IFM","cost": 80.00},
    {"po": "003","supplier": "RS","cost": 12.99},
    {"po": "004","supplier": "Amazon","cost": 45.00},
    {"po": "003","supplier": "Brammer","cost": 30.00}
]

seen_pos = []

for order in orders:
    po = order["po"]
    if po in seen_pos:
        print("Duplicate PO found:", po)
    else:
        seen_pos.append(po)