suppliers = [
    {"supplier_name":"Pack Support", "lead_time_days": 3, "is_24_7": False},
    {"supplier_name":"D2", "lead_time_days": 4, "is_24_7": True},
    {"supplier_name":"EMBA", "lead_time_days": 8, "is_24_7": False},
    {"supplier_name":"Techbelt", "lead_time_days": 6, "is_24_7": True},
    {"supplier_name":"Vulcatech", "lead_time_days": 7, "is_24_7": False},
]

for supplier in suppliers:
    supplier_name = supplier["supplier_name"]
    lead_time_days = supplier["lead_time_days"]
    is_24_7 = supplier["is_24_7"]
    if lead_time_days >= 7:
        action = "Chase supplier"
    elif lead_time_days >= 4:
        action = "Monitor"
    else:
        action = "No action"
    print(f"{supplier_name}: Action requered {action}")
    if is_24_7 is True:
        print("Emergency support available")
        