suppliers = [
    {"supplier_name": "Pack support", "lead_time_days": 3},
    {"supplier_name": "D2", "lead_time_days": 4},
    {"supplier_name": "EMBA", "lead_time_days": 5},
    {"supplier_name": "Techbelt", "lead_time_days": 2},
    {"supplier_name": "Vulcatech", "lead_time_days": 7},

]

suppliers_with_lead_time_longer_than_3 = 0
print("Suppliers with lead time longer than 3 days:")
for supplier in suppliers:
    supplier_name = supplier["supplier_name"]
    lead_time_days = supplier["lead_time_days"]
    if lead_time_days > 3:
        suppliers_with_lead_time_longer_than_3 += 1
        print(supplier_name)
print(f"Count of how many supplier have lead time longer than 3 days: {suppliers_with_lead_time_longer_than_3}")


