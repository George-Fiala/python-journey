suppliers = [
    {"supplier_name":"Pack Support", "lead_time_days": 3, "is_24_7": False},
    {"supplier_name":"D2", "lead_time_days": 4, "is_24_7": True},
    {"supplier_name":"EMBA", "lead_time_days": 8, "is_24_7": False},
    {"supplier_name":"Techbelt", "lead_time_days": 6, "is_24_7": True},
    {"supplier_name":"Vulcatech", "lead_time_days": 7, "is_24_7": False},
]
count_24_7_suppliers = 0
for supplier in suppliers:
    supplier_name = supplier["supplier_name"]
    lead_time_days = supplier["lead_time_days"]
    is_24_7 = supplier["is_24_7"]
    if is_24_7 is True:
        count_24_7_suppliers += 1

        print(f"Supplier: {supplier_name} - Lead time: {lead_time_days} days")
print(f"Count of 24/7 suppliers : {count_24_7_suppliers}")
