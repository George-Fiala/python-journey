suppliers = [
    {"supplier_name": "Pack Support", "lead_time_days": 3, "is_24_7": False, "late_deliveries": 1},
    {"supplier_name": "D2", "lead_time_days": 4, "is_24_7": True, "late_deliveries": 0},
    {"supplier_name": "EMBA", "lead_time_days": 8, "is_24_7": False, "late_deliveries": 3},
    
]

def get_risk_level(supplier):
    late_deliveries = supplier["late_deliveries"]
    if late_deliveries >= 3:
        return "High risk"
    elif late_deliveries >= 1:
        return "Medium risk"
    return "Low risk"


for supplier in suppliers:
    risk_level = get_risk_level(supplier)
    supplier_name = supplier["supplier_name"]
    late_deliveries = supplier["late_deliveries"]
    if late_deliveries == 1:
        delivery_word = "delivery"
    else:
        delivery_word = "deliveries"
        
    print(f"{supplier_name} - {late_deliveries} late {delivery_word} - {risk_level}")
