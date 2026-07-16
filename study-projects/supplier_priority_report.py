suppliers = [
    {"supplier_name": "Pack Support", "lead_time_days": 3},
    {"supplier_name": "D2", "lead_time_days": 4},
    {"supplier_name": "EMBA", "lead_time_days": 5},
    {"supplier_name": "Techbelt", "lead_time_days": 6},
    {"supplier_name": "Vulcatech", "lead_time_days": 7},
]

for supplier in suppliers:
    supplier_name = supplier["supplier_name"]
    lead_time_days = supplier["lead_time_days"]
    if lead_time_days <= 3:
        priority = "Fast"
    elif lead_time_days <= 5:
        priority = "Medium"
    else:
        priority = "Slow"

    print(f"{supplier_name}: {priority}")