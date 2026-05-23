# Inventory & Ticket Management System  
**ITM 352 – Final Project**

**Student:** Throy Nicdao  
**Course:** ITM 352  
**Semester:** Fall 2025  

---

## Project Overview

The purpose of this project was to design and build a web-based Inventory and Ticket Management System using Python and Flask. The system is intended to solve a real-world problem I experience in my IT work environment, where inventory is often tracked in spreadsheets and technical issues are handled through scattered emails. This makes it easy to lose track of devices and support requests.

This application centralizes both inventory tracking and IT support tickets into one platform. Users can add, edit, and delete devices, as well as submit and manage support tickets. The goal of the project is to demonstrate how Management Information Systems tools can be used to improve operational efficiency and organization.

---

## Use of AI in This Project

AI tools were used as a support resource, not to build the entire project automatically. Specifically, AI was used to:

- Help generate the initial structure of the Flask application  
- Assist with route planning and JSON data handling logic  
- Suggest improvements for code organization and validation  

All implementation, debugging, testing, and customization were done manually. The project was adjusted to meet the course requirements and reflect real-world IT workflows. I reviewed and modified all AI-assisted code to ensure I understood how it worked and that it met the project’s goals.

---

## Technologies Used

- Python 3  
- Flask  
- HTML (Jinja2 Templates)  
- JSON (for data storage)  

No external database was used. All data is stored locally using JSON files.

---

## Project Structure

The application is organized as follows:

- `app.py` – Main Flask application containing routes and logic  
- `templates/` – HTML templates for each page  
- `data/` – JSON files used to store inventory and ticket data  
  - `inventory.json`  
  - `tickets.json`  

---

## Application Features

### Dashboard

The homepage displays a dashboard that provides:

- Total number of devices in inventory  
- Devices grouped by location  
- Total number of open and in-progress tickets  
- Tickets grouped by category  
- Five most recent tickets  

This gives a quick overview of system activity.

---

### Inventory Management

The inventory module allows users to:

- View all devices in a searchable and filterable list  
- Add new devices with details such as model, serial number, condition, and location  
- Edit existing device records  
- Delete devices from inventory  

A location code system is used to standardize office locations and reduce data entry errors.

---

### Ticket Management

The ticket module allows users to:

- Submit new support tickets  
- View all tickets with filters for status, category, and priority  
- View individual ticket details  
- Update ticket status (Open, In Progress, Completed, Closed)  
- Delete tickets  

Tickets are automatically timestamped when created and updated.

---

## Data Storage

The system uses JSON files for persistent storage:

- Inventory data is stored in `inventory.json`  
- Ticket data is stored in `tickets.json`  

Each record is assigned a unique ID. Data files are automatically created when the application starts if they do not already exist.

---

## How to Run the Application

1. Navigate to the project folder in Terminal  
2. Activate the virtual environment  
3. Run the application using Python  

The application runs locally on port **5001** and can be accessed through a web browser.

---

## Limitations

- No user authentication or role management  
- No database (JSON storage only)  
- Designed for small-scale use  
- No concurrent multi-user write protection  

