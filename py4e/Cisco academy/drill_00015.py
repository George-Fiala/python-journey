leak_record = {
    "id": 101,
    "status": "Fixed",
    "pressure_loss": 0.0
}
if "severity" not in leak_record:
    leak_record["severity"] = "High"


del leak_record["id"]

for key, value in leak_record.items():
    print(key, ":", value)
