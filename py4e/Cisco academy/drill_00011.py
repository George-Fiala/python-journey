leak_record = {
    "id": 101,
    "status": "Active",
    "pressure_loss": 5.5
}
leak_record["status"] = "Fixed"
leak_record["pressure_loss"] = 0.0
leak_record["technician"] = "Martin"

for key, value in leak_record.items():
    print(key, ":", value)
