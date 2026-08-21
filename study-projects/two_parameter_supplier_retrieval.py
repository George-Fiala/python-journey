suppliers = [
    {"supplier_name": "D2", "lead_time": 3, "is_24_7": True},
    {"supplier_name": "EMBA", "lead_time": 8, "is_24_7": False},
    {"supplier_name": "Festo", "lead_time": 6, "is_24_7": False}
]


def get_action_needed(lead_time, is_24_7):
    if lead_time >= 7 and not is_24_7:
        return "Chase supplier"
    return "Monitor"


for supplier in suppliers:
    lead_time = supplier["lead_time"]
    is_24_7 = supplier["is_24_7"]
    supplier_name = supplier["supplier_name"]
    action_needed = get_action_needed(lead_time, is_24_7)
    print(f"{supplier_name} - {action_needed}")