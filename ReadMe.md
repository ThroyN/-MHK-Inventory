# MHK Inventory & Ticket Management System

**Developer:** Throy Nicdao

---

## Overview

A web-based IT inventory and support ticket management system built with Python and Flask. Designed for real-world IT use at MHK, it centralizes device tracking and support requests across multiple office locations and islands.

---

## Technologies Used

- **Python 3** — application logic
- **Flask** — web framework and routing
- **SQLite** — persistent database (via `data/inventory.db`)
- **openpyxl** — Excel import/export
- **HTML / Jinja2** — templating

---

## Project Structure

```
app.py                  — Main Flask application (routes, logic, DB helpers)
templates/              — HTML templates for each page
data/
  inventory.db          — SQLite database (all data stored here)
```

---

## How to Run

1. Open Terminal and navigate to the project folder
2. Run the app:

```bash
python3 app.py
```

3. Open your browser to **http://localhost:5001**

---

## Features

### User Sessions

When you open the app you are prompted to select a user before accessing anything. This logs all inventory changes under your name.

- Select an existing user from the list
- Add a new user by name
- Delete a user
- Switch users at any time via the logout button

---

### Dashboard

The homepage gives a live snapshot of the system:

- Total number of active devices
- Device count grouped by location
- Total open and in-progress tickets
- Ticket breakdown by category
- Five most recently submitted tickets
- Filter the entire dashboard by island

---

### Inventory Management

#### Viewing Inventory

- Searchable table of all devices (search by model, serial number, user, or location)
- Filter by device type, location, or island
- Toggle between **Active**, **Archived**, and **All** devices

#### Adding a Device

Each device record stores:

| Field | Description |
|---|---|
| Device Type | Laptop, Desktop, iPhone, etc. |
| Model | Device model name |
| Serial Number / Service Tag | Unique identifier |
| Assigned User | Person the device is assigned to |
| Password | Device login password |
| Purchase / Warranty Date | Date of purchase or warranty end |
| Condition | New, Good, Fair, Poor, or Broken |
| Location Code | Office location code (e.g. 101, 102) |
| Island | Hawaii island the device is on |
| Phone Number | Associated phone number |
| Notes | Free-text notes |

#### Editing a Device

All fields can be edited. Every edit is logged to the change history with a before/after record of what changed.

#### Deleting a Device

Permanently removes the device. A history entry is saved so the deletion can be undone.

#### Archiving a Device

Marks a device as retired/inactive without deleting it. Archived devices:

- Are hidden from the default inventory view
- Can be viewed by switching the filter to "Archived"
- Can be restored to active status at any time

#### Mark as TBD

Used when a device is being reassigned. The original record is archived and a new active copy is created with the Assigned User, Password, and Phone set to **TBD**. Both actions are logged and can be undone together.

---

### Change History & Undo

Every action taken on a device is recorded in the history log:

- Added, Edited, Deleted, Archived, Restored, Mark as TBD

The history page shows:

- Timestamp and who performed the action
- Device details at the time of the action
- What fields changed (for edits — shows old value → new value)

**Undo** is available for all action types:

| Action | Undo behavior |
|---|---|
| Added | Deletes the device |
| Edited | Reverts all changed fields to their previous values |
| Deleted | Restores the device from snapshot |
| Archived | Restores device to active |
| Restored | Re-archives the device |
| Mark as TBD | Removes TBD copy and un-archives the original |

The history log can be searched by model, serial number, location, or assigned user, and filtered by action type.

---

### Import & Export

#### Import (CSV or Excel)

Bulk-import devices from a spreadsheet:

- **CSV** — uploaded and imported immediately
- **Excel (.xlsx)** — upload the file, then choose which sheet to import from

The importer is flexible with column names — it recognizes common aliases such as `Username`, `Service Tag`, `Serial #`, `Location`, etc.

Required columns: `device_type`, `location_code`

A **CSV template** can be downloaded from the import page as a starting point.

Rows with missing required fields are skipped and a summary is shown after import.

#### Export (Excel)

Exports inventory to a formatted `.xlsx` file. Choose to export:

- **Active** devices only
- **Archived** devices only
- **All** devices

The export matches the column layout of the original MHK spreadsheet format.

---

### Ticket Management

#### Submitting a Ticket

Each ticket captures:

| Field | Description |
|---|---|
| Title | Short description of the issue |
| Description | Full details |
| Category | Hardware, Software, Network/WiFi, Password Reset, Setup/Installation, Other |
| Priority | Low, Medium, High, Urgent |
| Submitted By | Name of the person submitting |

#### Managing Tickets

- View all tickets with filters for status, category, and priority
- Search tickets by title, description, or submitter
- Update ticket status: **Open → In Progress → Completed → Closed**
- View full ticket detail page
- Delete a ticket

Tickets are sorted newest-first by default.

---

### Location & Island Management

Locations and islands are managed directly from the Inventory page:

- **Add / remove location codes** — locations can be tied to an island
- **Add / remove islands** — used to filter inventory and the dashboard
- Removing a location or island that is still in use shows a warning and requires confirmation

---

### Device Type Management

The list of device types shown in dropdowns is fully customizable:

- Add new device types
- Remove existing ones (with a warning if any devices still use that type)
- Any device type imported from a spreadsheet that doesn't exist yet is added automatically

---

## Location Codes

| Code | Location |
|---|---|
| 101 | Hilo Office – 45 Mohouli St |
| 102 | Kona Office – 75 Kuakini Hwy |
| 103 | Honolulu Office – 1001 Bishop St |
| 104 | Kahului Office – 33 Lono Ave |
| 105 | Lihue Office – 4444 Rice St |
| 201 | Remote Office – Big Island |
| 202 | Remote Office – Oahu |
| 999 | Mobile / Field Device |

---

## Notes

- All data is stored in a single SQLite database file (`data/inventory.db`)
- The app runs on port **5001** by default
- Designed for single-office or small-team use
- No external server or internet connection required
