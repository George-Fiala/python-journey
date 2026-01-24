import csv
from datetime import datetime, timedelta
from pathlib import Path

FILE_NAME = "air_leaks.csv"

SEVERITY_SLA_DAYS = {
    "HIGH": 2,    # fix within 48 hours
    "MEDIUM": 7,  # fix within 1 week
    "LOW": 14,    # fix within 2 weeks
}

FIELDNAMES = [
    "id",
    "date_found",
    "machine",
    "location",
    "severity",
    "status",
    "due_date",
    "work_order",
    "date_fixed",
    "notes",
]


def load_leaks(filename=FILE_NAME):
    path = Path(filename)
    if not path.exists():
        # create empty file with header
        with path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
            writer.writeheader()
        return []

    with path.open("r", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        return list(reader)


def save_leaks(leaks, filename=FILE_NAME):
    path = Path(filename)
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(leaks)


def generate_next_id(leaks):
    if not leaks:
        return 1
    max_id = max(int(l["id"]) for l in leaks)
    return max_id + 1


def prompt_severity():
    while True:
        print("Severity options: [H]igh (48h), [M]edium (7 days), [L]ow (14 days)")
        choice = input("Enter severity (H/M/L): ").strip().upper()
        if choice in ("H", "M", "L"):
            return {"H": "HIGH", "M": "MEDIUM", "L": "LOW"}[choice]
        print("Invalid choice. Please enter H, M, or L.")


def add_leak(leaks):
    print("\n--- Add New Leak ---")
    machine = input("Machine/Area (e.g. Emba 2, Corrugator): ").strip()
    location = input("Exact location (bay, side, reference): ").strip()
    severity = prompt_severity()
    work_order = input("Work order number (press Enter if none yet): ").strip()
    notes = input("Notes (optional): ").strip()

    today = datetime.today().date()
    days_to_fix = SEVERITY_SLA_DAYS[severity]
    due_date = today + timedelta(days=days_to_fix)

    leak = {
        "id": str(generate_next_id(leaks)),
        "date_found": today.isoformat(),
        "machine": machine,
        "location": location,
        "severity": severity,
        "status": "OPEN",
        "due_date": due_date.isoformat(),
        "work_order": work_order,
        "date_fixed": "",
        "notes": notes,
    }
    leaks.append(leak)
    print(f"Leak #{leak['id']} added. Due by {leak['due_date']} ({severity}).")


def list_leaks(leaks, only_open=False):
    print("\n--- Leak List ---")
    filtered = [l for l in leaks if (not only_open or l["status"] == "OPEN")]
    if not filtered:
        print("No leaks to show.")
        return

    # small ordering: open first, then by due date, then by severity
    severity_order = {"HIGH": 0, "MEDIUM": 1, "LOW": 2}
    filtered.sort(key=lambda l: (
        l["status"] != "OPEN",
        l["due_date"],
        severity_order.get(l["severity"], 99),
    ))

    today = datetime.today().date()

    for leak in filtered:
        # work out if it's overdue
        due_date_obj = datetime.fromisoformat(leak["due_date"]).date()
        is_overdue = (leak["status"] == "OPEN") and (due_date_obj < today)
        overdue_tag = "  **OVERDUE**" if is_overdue else ""

        print(f"#{leak['id']} | {leak['status']:5} | {leak['severity']:6} | "
              f"Found: {leak['date_found']} | Due: {leak['due_date']}{overdue_tag}")
        print(f"   Machine: {leak['machine']} | Location: {leak['location']}")
        if leak["work_order"]:
            print(f"   WO: {leak['work_order']}")
        if leak["date_fixed"]:
            print(f"   Fixed: {leak['date_fixed']}")
        if leak["notes"]:
            print(f"   Notes: {leak['notes']}")
        print("-" * 60)



def mark_fixed(leaks):
    print("\n--- Mark Leak as Fixed ---")
    leak_id = input("Enter leak ID to close: ").strip()
    leak = next((l for l in leaks if l["id"] == leak_id), None)
    if not leak:
        print("No leak found with that ID.")
        return
    if leak["status"] == "CLOSED":
        print("This leak is already closed.")
        return

    today = datetime.today().date().isoformat()
    leak["status"] = "CLOSED"
    leak["date_fixed"] = today

    if not leak["work_order"]:
        wo = input("Work order number (optional): ").strip()
        leak["work_order"] = wo

    print(f"Leak #{leak_id} marked as CLOSED on {today}.")


def main_menu():
    leaks = load_leaks()
    while True:
        print("\n=== IFM Air Leak Tracker ===")
        print("1) Add new leak")
        print("2) List OPEN leaks")
        print("3) List ALL leaks")
        print("4) Mark leak as fixed")
        print("5) Save & Exit")

        choice = input("Choose option (1-5): ").strip()
        if choice == "1":
            add_leak(leaks)
        elif choice == "2":
            list_leaks(leaks, only_open=True)
        elif choice == "3":
            list_leaks(leaks, only_open=False)
        elif choice == "4":
            mark_fixed(leaks)
        elif choice == "5":
            save_leaks(leaks)
            print("Data saved. Goodbye.")
            break
        else:
            print("Invalid choice, try again.")


if __name__ == "__main__":
    main_menu()

