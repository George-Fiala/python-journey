suppliers = [
    {"supplier_name": "RS Components", "lead_time_days": 3, "late_deliveries": 1, "total_orders": 20},
    {"supplier_name": "Festo", "lead_time_days": 5, "late_deliveries": 4, "total_orders": 18},
    {"supplier_name": "Gopfert", "lead_time_days": 9, "late_deliveries": 2, "total_orders": 8},
    {"supplier_name": "D2", "lead_time_days": 4, "late_deliveries": 0, "total_orders": 15}
]


def get_late_delivery_rate(supplier):
    late_deliveries = supplier["late_deliveries"]
    total_orders = supplier["total_orders"]
    late_delivery_rate = (late_deliveries / total_orders) * 100
    return late_delivery_rate


def get_performance_level(late_delivery_rate):
    if late_delivery_rate >= 20:
        return "Poor"
    elif late_delivery_rate >= 10:
        return "Watch"
    return "Good"


def get_lead_time_level(lead_time_days):
    if lead_time_days >= 7:
        return "Slow"
    elif lead_time_days >= 4:
        return "Average"
    return "Fast"

total_late_deliveries = 0
total_orders = 0
total_lead_time = 0
poor_supplier_count = 0
watch_supplier_count = 0
good_supplier_count = 0

for supplier in suppliers:
    late_delivery_rate = get_late_delivery_rate(supplier)
    performance_level = get_performance_level(late_delivery_rate)
    lead_time_days = supplier["lead_time_days"]
    lead_time_level = get_lead_time_level(lead_time_days)
    supplier_name = supplier["supplier_name"]
    late_deliveries = supplier["late_deliveries"]
    supplier_total_orders = supplier["total_orders"]
    total_late_deliveries += late_deliveries
    total_orders += supplier_total_orders
    total_lead_time += lead_time_days
    if performance_level == "Poor":
        poor_supplier_count += 1
    elif performance_level == "Watch":
        watch_supplier_count += 1
    elif performance_level == "Good":
        good_supplier_count += 1
    print(f"{supplier_name} - {late_delivery_rate:.2f}% - {performance_level} - {lead_time_level}")
print(f"Total late deliveries: {total_late_deliveries}")
print(f"Total orders: {total_orders}")
print(f"Total lead time: {total_lead_time}")
print(f"Poor suppliers: {poor_supplier_count}")
print(f"Suppliers to watch: {watch_supplier_count}")
print(f"Good suppliers: {good_supplier_count}")

overall_late_delivery_rate = total_late_deliveries / total_orders * 100
overall_performance = get_performance_level(overall_late_delivery_rate)
print(f"Overall late delivery rate: {overall_late_delivery_rate:.2f}% - {overall_performance}")


average_lead_time = total_lead_time / len(suppliers)
print(f"Average lead time: {average_lead_time:.2f}")

classified_supplier_count = poor_supplier_count + watch_supplier_count + good_supplier_count

if classified_supplier_count == len(suppliers):
    print("Supplier counts match")
else:
    print("Supplier count mismatch")