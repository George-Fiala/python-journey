suppliers = [
    {"supplier_name": "RS Components", "late_deliveries": 1, "total_orders": 20},
    {"supplier_name": "Festo", "late_deliveries": 4, "total_orders": 18},
    {"supplier_name": "Gopfert", "late_deliveries": 2, "total_orders": 8},
    {"supplier_name": "D2", "late_deliveries": 0, "total_orders": 15}
]


def get_late_deliveries_percentage(late_delivery_percentage):
    if late_delivery_percentage >= 20:
        return "Poor"
    return "Good"



total_orders_count = 0
total_late_deliveries_count = 0
poor_suppliers_count = 0


for supplier in suppliers:
    total_orders = supplier["total_orders"]
    late_deliveries = supplier["late_deliveries"]
    supplier_name = supplier["supplier_name"]
    late_delivery_percentage = late_deliveries / total_orders * 100
    supplier_status = get_late_deliveries_percentage(late_delivery_percentage)
    total_orders_count += total_orders
    total_late_deliveries_count += late_deliveries
    if supplier_status == "Poor":
        poor_suppliers_count += 1
    print(f"{supplier_name} - {late_delivery_percentage:.2f}% - {supplier_status}")
print(f"Late deliveries count: {total_late_deliveries_count}")
print(f"Total orders count: {total_orders_count}")
print(f"Poor suppliers count: {poor_suppliers_count}")
    


late_deliveries_percentage = total_late_deliveries_count / total_orders_count * 100
print(f"Late deliveries percentage: {late_deliveries_percentage:.2f}%")