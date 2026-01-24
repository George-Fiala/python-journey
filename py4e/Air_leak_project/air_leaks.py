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

    for leak in filtered:
        print(f"#{leak['id']} | {leak['status']:5} | {leak['severity']:6} | "
              f"Found: {leak['date_found']} | Due: {leak['due_date']}")
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

def summary_report(leaks):
    print("\n--- Summary Report ---")

    total = len(leaks)
    open_leaks = [l for l in leaks if l["status"] == "OPEN"]
    closed_leaks = [l for l in leaks if l["status"] == "CLOSED"]

    # counts by severity for OPEN leaks
    sev_counts = {"HIGH": 0, "MEDIUM": 0, "LOW": 0}
    for l in open_leaks:
        sev = l.get("severity", "").upper()
        if sev in sev_counts:
            sev_counts[sev] += 1

    # overdue = OPEN and due_date before today
    today = datetime.today().date()
    overdue = []
    for l in open_leaks:
        try:
            due = datetime.fromisoformat(l["due_date"]).date()
            if due < today:
                overdue.append(l)
        except Exception:
            # if due_date is missing or broken, ignore it
            pass

    print(f"Total leaks: {total}")
    print(f"OPEN:  {len(open_leaks)}")
    print(f"CLOSED:{len(closed_leaks)}")
    print("\nOPEN by severity:")
    print(f"  HIGH:   {sev_counts['HIGH']}")
    print(f"  MEDIUM: {sev_counts['MEDIUM']}")
    print(f"  LOW:    {sev_counts['LOW']}")
    print(f"\nOverdue (OPEN past due date): {len(overdue)}")

    if overdue:
        print("Overdue IDs: " + ", ".join(l["id"] for l in overdue))


def export_open_leaks(leaks, filename="open_air_leaks_report.csv"):
    open_leaks = [l for l in leaks if l["status"] == "OPEN"]
    if not open_leaks:
        print("\nNo OPEN leaks to export.")
        return

    path = Path(filename)
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(open_leaks)

    print(f"\nExported {len(open_leaks)} OPEN leaks to: {path.resolve()}")
def update_leak(leaks):
    print("\n--- Update Leak (edit existing) ---")
    leak_id = input("Enter leak ID to edit: ").strip()
    leak = next((l for l in leaks if l["id"] == leak_id), None)

    if not leak:
        print("No leak found with that ID.")
        return

    print("\nPress Enter to keep current value.\n")

    # MACHINE
    current_machine = leak.get("machine", "")
    new_machine = input(f"Machine/Area [{current_machine}]: ").strip()
    if new_machine:
        leak["machine"] = new_machine

    # LOCATION
    current_location = leak.get("location", "")
    new_location = input(f"Exact location [{current_location}]: ").strip()
    if new_location:
        leak["location"] = new_location

    # SEVERITY (can change; status stays untouched)
    current_sev = leak.get("severity", "MEDIUM").upper()
    severity_changed = False
    while True:
        new_sev = input(
            f"Severity [{current_sev}] (H/M/L or Enter to keep): "
        ).strip().upper()

        if not new_sev:
            break  # keep current

        if new_sev in ("H", "HIGH"):
            leak["severity"] = "HIGH"
            severity_changed = (leak["severity"] != current_sev)
            break
        elif new_sev in ("M", "MEDIUM"):
            leak["severity"] = "MEDIUM"
            severity_changed = (leak["severity"] != current_sev)
            break
        elif new_sev in ("L", "LOW"):
            leak["severity"] = "LOW"
            severity_changed = (leak["severity"] != current_sev)
            break
        else:
            print("Invalid severity. Use H/M/L or press Enter to keep.")

    # if severity changed, recalc due_date based on original date_found + SLA
    if severity_changed:
        try:
            found_date = datetime.fromisoformat(leak["date_found"]).date()
        except Exception:
            found_date = datetime.today().date()

        days_to_fix = SEVERITY_SLA_DAYS[leak["severity"]]
        new_due = found_date + timedelta(days=days_to_fix)
        leak["due_date"] = new_due.isoformat()
        print(
            f"Due date recalculated to {leak['due_date']} "
            f"based on {leak['severity']} Service Level Agreement (SLA)."
        )

    # WORK ORDER (WO)
    current_wo = leak.get("work_order", "")
    new_wo = input(f"Work order number (WO) [{current_wo}]: ").strip()
    if new_wo:
        leak["work_order"] = new_wo

    # NOTES
    current_notes = leak.get("notes", "")
    new_notes = input(f"Notes [{current_notes}]: ").strip()
    if new_notes:
        leak["notes"] = new_notes

    print(f"\nLeak #{leak_id} updated (status unchanged: {leak['status']}).")

def main_menu():
    leaks = load_leaks()
    while True:
        print("\n=== IFM Air Leak Tracker ===")
        print("1) Add new leak")
        print("2) List OPEN leaks")
        print("3) List ALL leaks")
        print("4) Mark leak as fixed")
        print("5) Update leak (edit)")
        print("6) Summary report")
        print("7) Export OPEN leaks report (CSV)")
        print("8) Save & Exit")

        choice = input("Choose option (1-8): ").strip()
        if choice == "1":
            add_leak(leaks)
        elif choice == "2":
            list_leaks(leaks, only_open=True)
        elif choice == "3":
            list_leaks(leaks, only_open=False)
        elif choice == "4":
            mark_fixed(leaks)
        elif choice == "5":
            update_leak(leaks)
        elif choice == "6":
            summary_report(leaks)
        elif choice == "7":
            export_open_leaks(leaks)
        elif choice == "8":
            save_leaks(leaks)
            print("Data saved. Goodbye.")
            break
        else:
            print("Invalid choice, try again.")




if __name__ == "__main__":
    main_menu()
