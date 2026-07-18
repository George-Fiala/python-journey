suppliers = [
    {"supplier_name":"Pack Support", "lead_time_days": 3, "is_24_7": False},
    {"supplier_name":"D2", "lead_time_days": 4, "is_24_7": True},
    {"supplier_name":"EMBA", "lead_time_days": 8, "is_24_7": False},
    {"supplier_name":"Techbelt", "lead_time_days": 6, "is_24_7": True},
    {"supplier_name":"Vulcatech", "lead_time_days": 7, "is_24_7": False},
]

low_risk = 0
medium_risk = 0
high_risk = 0

for supplier in suppliers:
    supplier_name = supplier["supplier_name"]
    lead_time_days = supplier["lead_time_days"]
    if lead_time_days <= 3:
        risk = "Low"
        low_risk += 1
    elif lead_time_days <= 6:
        risk = "Medium"
        medium_risk +=1
    else:
        risk = "High"
        high_risk += 1
    print(f"{supplier_name}: {risk}" )
print(f"Low risk suppliers: {low_risk}")
print(f"Medium risk suppliers: {medium_risk}")
print(f"High risk suppliers: {high_risk}")