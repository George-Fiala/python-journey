suppliers = [
    {"supplier_name":"Pack Support", "lead_time_days": 3, "is_24_7": False},
    {"supplier_name":"D2", "lead_time_days": 4, "is_24_7": True},
    {"supplier_name":"EMBA", "lead_time_days": 8, "is_24_7": False},
    {"supplier_name":"Techbelt", "lead_time_days": 6, "is_24_7": True},
    {"supplier_name":"Vulcatech", "lead_time_days": 7, "is_24_7": False},
]

countsupplier_with_lead_time_days_higher_or_same_as_6 = 0
print("Supplier with lead time 6 days or more:")
for supplier in suppliers:
    supplier_name = supplier["supplier_name"]
    lead_time_days = supplier["lead_time_days"]
    if lead_time_days >= 6:
        countsupplier_with_lead_time_days_higher_or_same_as_6 += 1
        print(f"{supplier_name} - Lead time: {lead_time_days} days")
print(f"Follow-up required: {countsupplier_with_lead_time_days_higher_or_same_as_6}")
