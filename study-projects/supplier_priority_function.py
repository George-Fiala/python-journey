suppliers = [
    {"supplier_name": "Pack Support", "lead_time_days": 3, "is_24_7": False},
    {"supplier_name": "D2", "lead_time_days": 4, "is_24_7": True},
    {"supplier_name": "EMBA", "lead_time_days": 8, "is_24_7": False},
]

def get_priority(supplier):
    lead_time_days = supplier["lead_time_days"]
    if lead_time_days >= 7:
        return "High"
        
    elif lead_time_days >= 4 and lead_time_days <= 6:
        return "Medium"
    else:
        return "Low"
for supplier in suppliers:
    supplier_name = supplier["supplier_name"]
    priority = get_priority(supplier)
    print(supplier_name, priority)

        