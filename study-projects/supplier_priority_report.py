suppliers = [
    {"supplier_name": "Pack Support", "lead_time_days": 3},
    {"supplier_name": "D2", "lead_time_days": 4},
    {"supplier_name": "EMBA", "lead_time_days": 5},
    {"supplier_name": "Techbelt", "lead_time_days": 6},
    {"supplier_name": "Vulcatech", "lead_time_days": 7},
]

fast_lead_time = 0
medium_lead_time = 0
slow_lead_time = 0

for supplier in suppliers:
    supplier_name = supplier["supplier_name"]
    lead_time_days = supplier["lead_time_days"]
    if lead_time_days <= 3:
        priority = "Fast"
        fast_lead_time += 1
    elif lead_time_days <= 5:
        priority = "Medium"
        medium_lead_time += 1
    else:
        priority = "Slow"
        slow_lead_time += 1

    print(f"{supplier_name}: {priority}")
print(f"Fast lead time count: {fast_lead_time}")
print(f"Medium lead time count: {medium_lead_time}")
print(f"Slow lead time count: {slow_lead_time}")