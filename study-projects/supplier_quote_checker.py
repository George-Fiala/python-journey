quotations = [
    {"supplier": "NORD", "price": 420, "lead_time": 5, "approved": True},
    {"supplier": "ERIKS", "price": 390, "lead_time": 12, "approved": True},
    {"supplier": "Rubix", "price": 450, "lead_time": 3, "approved": False},
    {"supplier": "BearingCo", "price": 410, "lead_time": 7, "approved": True}
]


def get_quote_status(price, lead_time, approved):
    if not approved:
        return "Rejected"
    elif price <= 400 and lead_time <= 7:
        return "Best"
    elif price <=450 or lead_time <= 7:
        return "Acceptable"
    else:
        return "Poor"


best_quotes = 0
rejected_quotes = 0
approved_suppliers_count = 0


for quote in quotations:
    supplier = quote["supplier"]
    price = quote["price"]
    lead_time = quote["lead_time"]
    approved = quote["approved"]
    quote_status = get_quote_status(price, lead_time, approved)
    if quote_status == "Best":
        best_quotes += 1

    if quote_status == "Rejected":
        rejected_quotes += 1

    if approved:
        approved_suppliers_count += 1

    print(f"{supplier} - £{price:.2f} - {lead_time} days - {quote_status}")

print(f"Best quotes count: {best_quotes}")
print(f"Rejected quotes count: {rejected_quotes}")
print(f"Approved quotes suppliers count: {approved_suppliers_count}")