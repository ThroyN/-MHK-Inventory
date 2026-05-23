# MHK Inventory & Ticket Management System

**Developer:** Throy Nicdao

---

## What Is This App?

This is an internal tool built for MHK's IT team to keep track of all company devices and manage support requests in one place. Before this, inventory was scattered across spreadsheets and issues were handled through emails — things got lost, devices went unaccounted for, and there was no easy way to see what was going on at a glance.

This app fixes that. Everything lives in one place, it's easy to use, and nothing gets lost.

---

## How to Run It

Open Terminal, go to the project folder, and run:

```bash
python3 app.py
```

Then open your browser to **http://localhost:5001**

---

## Features

### Choosing a User

When you first open the app, it asks you to pick your name before you do anything. This is so the system knows who made what changes — every edit, addition, or deletion gets tagged with your name. You can add new users or remove old ones from the same screen.

---

### Dashboard

The dashboard is the first thing you see after logging in. It gives you a quick overview of everything without having to dig around:

- How many devices are currently active
- Which offices have the most devices
- How many support tickets are open right now
- What kinds of issues are being reported most
- The five most recent tickets so you can see what just came in

You can also filter the whole dashboard by island if you only want to see what's going on in, say, Oahu or the Big Island.

---

### Inventory

This is the main part of the app — a full record of every device the company owns.

#### Finding a device
You can search by model name, serial number, assigned user, or location. You can also filter down by device type (laptops only, phones only, etc.), by office location, or by island. This makes it easy to pull up exactly what you're looking for even if you have hundreds of devices.

#### Adding a device
When you add a new device, you fill in everything about it — the model, serial number, who it's assigned to, what condition it's in, where it's located, and any notes you want to attach. This becomes the permanent record for that device.

#### Editing a device
You can update any field on a device at any time. Every time you save a change, the app records exactly what was changed — so if someone updates the assigned user from "Dan" to "Tom," that's saved in the history forever.

#### Archiving a device
When a device is retired or taken out of service, you don't have to delete it. You can archive it instead. Archived devices disappear from the normal inventory view but are still in the system if you ever need to look back at them. You can restore an archived device to active status at any time.

#### Mark as TBD
This is useful when a device is being reassigned and you don't know who it's going to yet. Instead of leaving the old person's name on it, you can hit "Mark as TBD" — the system archives the original record to preserve the history, and creates a fresh active copy with the assigned user, password, and phone number all set to TBD. Once you know who it's going to, you just edit that copy.

---

### History & Undo

Every single change made in the inventory is logged — who did it, when, and what exactly changed. This is the audit trail.

If someone makes a mistake — accidentally deletes a device, archives the wrong one, or saves a bad edit — you can go into the history and hit **Undo** to reverse it. The app is smart about undoing: if you undo a deletion, it brings the full device back exactly as it was. If you undo an edit, it puts all the fields back to what they were before.

---

### Import & Export

#### Import
If you already have device records in a spreadsheet, you can bring them straight into the app without typing everything in one by one. The app accepts both CSV files and Excel (.xlsx) files. It's also flexible — if your spreadsheet has columns named "Username" or "Service Tag" instead of exactly what the app expects, it'll figure it out.

For Excel files with multiple sheets, it'll ask you which sheet to pull from before importing.

A template CSV is available to download if you want a clean starting point.

#### Export
You can export the full inventory to an Excel file at any time — just active devices, just archived ones, or everything. The export is formatted to match the original MHK spreadsheet layout, so it fits right into any existing workflows.

---

### Support Tickets

The ticket system is for logging and tracking IT issues that come in. Instead of getting a one-off email that disappears, every issue becomes a tracked ticket.

When someone submits a ticket they fill in:
- What the issue is (title and description)
- What kind of problem it is (hardware, software, network, password reset, etc.)
- How urgent it is (Low, Medium, High, or Urgent)
- Who submitted it

From there, the IT team can update the ticket's status as they work on it — Open, In Progress, Completed, or Closed. You can filter and search tickets to find what you're looking for, and each ticket has its own detail page so you can see the full picture in one place.

---

### Managing Locations, Islands, and Device Types

All the dropdown options in the app — office locations, islands, and device types — are fully customizable. You don't need to touch any code to add a new office location or a new device type. Just go to the inventory page, find the management panel, and add or remove options from there.

The app will warn you before you remove something that's still being used by devices, so you don't accidentally break anything.

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
