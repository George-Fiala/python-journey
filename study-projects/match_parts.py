ordered_parts = ["BRG001", "SNS002", "BLT003", "MTR004"]
stock_parts = ["BRG001", "BLT003", "FLT005"]

for ordered_part in ordered_parts:
    if ordered_part in stock_parts:
        print(f"Ordered part {ordered_part} already in stock.")
    else:
        print(f"Ordered part {ordered_part} not in stock yet.")
