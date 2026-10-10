quotes = [
    {"supplier": "NORD", "price": 420, "lead_time": 5, "approved": True},
    {"supplier": "ERIKS", "price": 390, "lead_time": 12, "approved": True},
    {"supplier": "Rubix", "price": 370, "lead_time": 3, "approved": False},
    {"supplier": "BearingCo", "price": 410, "lead_time": 7, "approved": True}
]



def find_best_quote(quotes):
    best_quote = None
    for quote in quotes:
        price = quote.get("price")
        lead_time = quote.get("lead_time")
        approved = quote.get("approved")
        if approved and lead_time <= 7:
           if best_quote is None:
               best_quote = quote
           elif best_quote["price"] > price:
               best_quote = quote
    return best_quote


winner = find_best_quote(quotes)


if winner is not None:
    supplier = winner.get("supplier")
    price = winner.get("price")
    lead_time = winner.get("lead_time")
    print(f"Best supplier: {supplier} - £{price:.2f} - {lead_time} days")
else:
    print("No suitable supplier found")