"""
One-time script: reads INVENTORY.csv and updates phone numbers
on existing inventory.json entries by matching serial number,
then by device_type + model + location as fallback.

Usage:
    python3 patch_phones.py /path/to/INVENTORY.csv
"""

import csv, io, json, sys, os

INVENTORY_FILE = 'data/inventory.json'

def normalize(val):
    return str(val).strip().lower() if val else ''

def load_csv(path):
    with open(path, 'r', encoding='utf-8-sig') as f:
        raw = f.read()
    # Find header row (first row containing 'Device Type')
    lines = raw.splitlines()
    header_idx = next(i for i, l in enumerate(lines) if 'device type' in l.lower())
    content = '\n'.join(lines[header_idx:])
    reader = csv.DictReader(io.StringIO(content))
    # Normalize header keys
    rows = []
    for row in reader:
        norm = {k.strip().lower(): v.strip() for k, v in row.items()}
        rows.append(norm)
    return rows

def main():
    if len(sys.argv) < 2:
        print("Usage: python3 patch_phones.py /path/to/INVENTORY.csv")
        sys.exit(1)

    csv_path = sys.argv[1]
    if not os.path.exists(csv_path):
        print(f"File not found: {csv_path}")
        sys.exit(1)

    rows = load_csv(csv_path)

    # Build lookup from CSV rows: serial → phone, and (type+model+loc) → phone
    serial_phone = {}
    fallback_phone = {}
    for row in rows:
        phone = row.get('phone', '').strip()
        if not phone:
            continue
        serial = normalize(row.get('computername/servicetag', ''))
        if serial and serial not in ('n/a', '', '-', 'n/a'):
            serial_phone[serial] = phone
        # fallback key
        key = (
            normalize(row.get('device type', '')),
            normalize(row.get('model', '')),
            normalize(row.get('location', ''))
        )
        if any(key):
            fallback_phone[key] = phone

    with open(INVENTORY_FILE, 'r') as f:
        inventory = json.load(f)

    updated = 0
    for device in inventory:
        if device.get('phone'):
            continue  # already has phone

        # Try serial match
        sn = normalize(device.get('serial_number', ''))
        if sn and sn not in ('n/a', '', '-'):
            if sn in serial_phone:
                device['phone'] = serial_phone[sn]
                updated += 1
                continue

        # Fallback: type + model + location
        key = (
            normalize(device.get('device_type', '')),
            normalize(device.get('model', '')),
            normalize(device.get('location_code', ''))
        )
        if key in fallback_phone:
            device['phone'] = fallback_phone[key]
            updated += 1

    with open(INVENTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(inventory, f, indent=2)

    print(f"Done. Updated {updated} devices with phone numbers.")
    no_phone = sum(1 for d in inventory if not d.get('phone'))
    print(f"Devices still without phone: {no_phone}")

if __name__ == '__main__':
    main()
