suppliers = [
    {"supplier_name": "Pack Support", "lead_time_days": 3, "is_24_7": False},
    {"supplier_name": "D2", "lead_time_days": 4, "is_24_7": True},
    {"supplier_name": "EMBA", "lead_time_days": 8, "is_24_7": False}
]

def get_order_action(supplier):
    lead_time_days = supplier["lead_time_days"]
    is_24_7 = supplier["is_24_7"]
    if lead_time_days >= 7 and not is_24_7:
        return "Chase supplier"
    return "Monitor"

def get_priority(supplier):
    lead_time_days = supplier["lead_time_days"]
    if lead_time_days >= 7:
        return "High"
    elif lead_time_days >= 4:
        return "Medium"
    return "Low"

for supplier in suppliers:
    order_action = get_order_action(supplier)
    priority = get_priority(supplier)
    supplier_name = supplier["supplier_name"]
    print(f"{supplier_name} - {order_action} - {priority}")