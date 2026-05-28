from flask import Flask, render_template, request, redirect, url_for, flash, Response, session
from datetime import datetime
from urllib.parse import quote
from werkzeug.security import generate_password_hash, check_password_hash
import json
import os
import csv
import io
import copy
import sqlite3
import tempfile
import openpyxl
from openpyxl.styles import Font

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY') or os.urandom(24)

# Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
DB_FILE  = os.path.join(DATA_DIR, 'inventory.db')

# Legacy JSON paths — kept only for one-time migration on first run
INVENTORY_FILE    = os.path.join(DATA_DIR, 'inventory.json')
TICKETS_FILE      = os.path.join(DATA_DIR, 'tickets.json')
HISTORY_FILE      = os.path.join(DATA_DIR, 'history.json')
USERS_FILE        = os.path.join(DATA_DIR, 'users.json')
DEVICE_TYPES_FILE = os.path.join(DATA_DIR, 'device_types.json')
LOCATIONS_FILE    = os.path.join(DATA_DIR, 'locations.json')
ISLANDS_FILE      = os.path.join(DATA_DIR, 'islands.json')

# Location codes mapping
LOCATION_CODES = {
    '101': 'Hilo Office – 45 Mohouli St',
    '102': 'Kona Office – 75 Kuakini Hwy',
    '103': 'Honolulu Office – 1001 Bishop St',
    '104': 'Kahului Office – 33 Lono Ave',
    '105': 'Lihue Office – 4444 Rice St',
    '201': 'Remote Office – Big Island',
    '202': 'Remote Office – Oahu',
    '999': 'Mobile/Field Device'
}

# Device types and ticket categories
DEVICE_TYPES = ['Laptop', 'Desktop', 'Laptop/Desktop', 'iPhone', 'Android Phone', 'Phone', 'Tablet', 'Monitor', 'Printer', 'Scanner', 'Router', 'Other']

DEVICE_TYPE_ICONS = {
    'Laptop': '💻',
    'Desktop': '🖥️',
    'Laptop/Desktop': '💻',
    'iPhone': '📱',
    'Android Phone': '📱',
    'Phone': '📱',
    'Tablet': '📱',
    'Monitor': '🖥️',
    'Printer': '🖨️',
    'Scanner': '🔍',
    'Router': '📡',
    'Other': '📦',
}

# Maps raw import values (lowercase) to canonical DEVICE_TYPES entries
DEVICE_TYPE_ALIASES = {
    'phone': 'Phone',
    'scanner': 'Scanner',
    'laptop-desktop': 'Laptop/Desktop',
    'laptop desktop': 'Laptop/Desktop',
    'laptop / desktop': 'Laptop/Desktop',
}
TICKET_CATEGORIES = ['Hardware Issue', 'Software Issue', 'Network/WiFi', 'Password Reset', 'Setup/Installation', 'Other']
TICKET_STATUSES = ['Open', 'In Progress', 'Completed', 'Closed']
PRIORITY_LEVELS = ['Low', 'Medium', 'High', 'Urgent']
DEVICE_CONDITIONS = ['New', 'Good', 'Fair', 'Poor', 'Broken']

@app.context_processor
def inject_globals():
    return {'device_type_icons': DEVICE_TYPE_ICONS}

@app.template_filter('urlencode')
def urlencode_filter(s):
    return quote(str(s))

# ── SQLite helpers ────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA foreign_keys=ON')
    return conn

def init_db():
    """Create all tables if they don't exist."""
    conn = get_db()
    conn.executescript('''
        CREATE TABLE IF NOT EXISTS inventory (
            id            INTEGER PRIMARY KEY,
            device_type   TEXT    DEFAULT '',
            model         TEXT    DEFAULT '',
            serial_number TEXT    DEFAULT '',
            assigned_user TEXT    DEFAULT '',
            password      TEXT    DEFAULT '',
            purchase_date TEXT    DEFAULT '',
            condition     TEXT    DEFAULT 'Good',
            location_code TEXT    DEFAULT '',
            island        TEXT    DEFAULT '',
            phone         TEXT    DEFAULT '',
            notes         TEXT    DEFAULT '',
            archived      INTEGER DEFAULT 0,
            archived_at   TEXT    DEFAULT '',
            added_at      TEXT    DEFAULT '',
            updated_at    TEXT    DEFAULT '',
            source        TEXT    DEFAULT 'added'
        );
        CREATE TABLE IF NOT EXISTS tickets (
            id           INTEGER PRIMARY KEY,
            title        TEXT DEFAULT '',
            description  TEXT DEFAULT '',
            category     TEXT DEFAULT '',
            priority     TEXT DEFAULT 'Medium',
            status       TEXT DEFAULT 'Open',
            submitted_by TEXT DEFAULT '',
            created_at   TEXT DEFAULT '',
            updated_at   TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS history (
            id              INTEGER PRIMARY KEY,
            timestamp       TEXT DEFAULT '',
            performed_by    TEXT DEFAULT '',
            action          TEXT DEFAULT '',
            device_id       INTEGER,
            device_type     TEXT DEFAULT '',
            model           TEXT DEFAULT '',
            serial_number   TEXT DEFAULT '',
            location_code   TEXT DEFAULT '',
            island          TEXT DEFAULT '',
            assigned_user   TEXT DEFAULT '',
            changes         TEXT DEFAULT '{}',
            device_snapshot TEXT DEFAULT '{}'
        );
        CREATE TABLE IF NOT EXISTS users (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            name          TEXT UNIQUE NOT NULL,
            password_hash TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS device_types (
            id   INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL
        );
        CREATE TABLE IF NOT EXISTS locations (
            id     INTEGER PRIMARY KEY AUTOINCREMENT,
            code   TEXT UNIQUE NOT NULL,
            island TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS islands (
            id   INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL
        );
    ''')
    # Add source column to existing databases that don't have it yet
    existing = [row[1] for row in conn.execute('PRAGMA table_info(inventory)').fetchall()]
    if 'source' not in existing:
        conn.execute("ALTER TABLE inventory ADD COLUMN source TEXT DEFAULT 'added'")
    # Add password_hash column to users if missing
    user_cols = [row[1] for row in conn.execute('PRAGMA table_info(users)').fetchall()]
    if 'password_hash' not in user_cols:
        conn.execute("ALTER TABLE users ADD COLUMN password_hash TEXT DEFAULT ''")
    conn.commit()
    conn.close()

def _migrate_json_to_db():
    """One-time migration: import existing JSON data into SQLite."""
    conn = get_db()

    def _json_load(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read().strip()
                return json.loads(content) if content else []
        except Exception:
            return []

    # Inventory
    if conn.execute('SELECT COUNT(*) FROM inventory').fetchone()[0] == 0:
        for d in _json_load(INVENTORY_FILE):
            conn.execute('''INSERT OR IGNORE INTO inventory
                (id,device_type,model,serial_number,assigned_user,password,
                 purchase_date,condition,location_code,island,phone,notes,
                 archived,archived_at,added_at,updated_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', (
                d.get('id'), d.get('device_type',''), d.get('model',''),
                d.get('serial_number',''), d.get('assigned_user',''),
                d.get('password',''), d.get('purchase_date',''),
                d.get('condition','Good'), d.get('location_code',''),
                d.get('island',''), d.get('phone',''), d.get('notes',''),
                1 if d.get('archived') else 0,
                d.get('archived_at',''), d.get('added_at',''), d.get('updated_at','')))

    # Tickets
    if conn.execute('SELECT COUNT(*) FROM tickets').fetchone()[0] == 0:
        for t in _json_load(TICKETS_FILE):
            conn.execute('''INSERT OR IGNORE INTO tickets
                (id,title,description,category,priority,status,submitted_by,created_at,updated_at)
                VALUES (?,?,?,?,?,?,?,?,?)''', (
                t.get('id'), t.get('title',''), t.get('description',''),
                t.get('category',''), t.get('priority','Medium'),
                t.get('status','Open'), t.get('submitted_by',''),
                t.get('created_at',''), t.get('updated_at','')))

    # History
    if conn.execute('SELECT COUNT(*) FROM history').fetchone()[0] == 0:
        for h in _json_load(HISTORY_FILE):
            conn.execute('''INSERT OR IGNORE INTO history
                (id,timestamp,performed_by,action,device_id,device_type,model,
                 serial_number,location_code,island,assigned_user,changes,device_snapshot)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''', (
                h.get('id'), h.get('timestamp',''), h.get('performed_by',''),
                h.get('action',''), h.get('device_id'),
                h.get('device_type',''), h.get('model',''),
                h.get('serial_number',''), h.get('location_code',''),
                h.get('island',''), h.get('assigned_user',''),
                json.dumps(h.get('changes',{})),
                json.dumps(h.get('device_snapshot',{}))))

    # Users
    if conn.execute('SELECT COUNT(*) FROM users').fetchone()[0] == 0:
        for name in _json_load(USERS_FILE):
            conn.execute('INSERT OR IGNORE INTO users (name) VALUES (?)', (name,))

    # Device types
    if conn.execute('SELECT COUNT(*) FROM device_types').fetchone()[0] == 0:
        types = _json_load(DEVICE_TYPES_FILE) or DEVICE_TYPES
        for t in types:
            conn.execute('INSERT OR IGNORE INTO device_types (name) VALUES (?)', (t,))

    # Locations
    if conn.execute('SELECT COUNT(*) FROM locations').fetchone()[0] == 0:
        for loc in _json_load(LOCATIONS_FILE):
            conn.execute('INSERT OR IGNORE INTO locations (code,island) VALUES (?,?)',
                         (loc.get('code',''), loc.get('island','')))

    # Islands
    if conn.execute('SELECT COUNT(*) FROM islands').fetchone()[0] == 0:
        for name in _json_load(ISLANDS_FILE):
            conn.execute('INSERT OR IGNORE INTO islands (name) VALUES (?)', (name,))

    conn.commit()
    conn.close()

# Initialize data storage
def init_data_storage():
    os.makedirs(DATA_DIR, exist_ok=True)
    init_db()
    _migrate_json_to_db()

# ── Data access functions ─────────────────────────────────────────────────────

def load_inventory():
    conn = get_db()
    rows = conn.execute('SELECT * FROM inventory').fetchall()
    conn.close()
    result = []
    for row in rows:
        d = dict(row)
        d['archived'] = bool(d.get('archived', 0))
        result.append(d)
    return result

def save_inventory(inventory):
    conn = get_db()
    try:
        conn.execute('BEGIN')
        conn.execute('DELETE FROM inventory')
        for d in inventory:
            conn.execute('''INSERT INTO inventory
                (id,device_type,model,serial_number,assigned_user,password,
                 purchase_date,condition,location_code,island,phone,notes,
                 archived,archived_at,added_at,updated_at,source)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', (
                d.get('id'), d.get('device_type',''), d.get('model',''),
                d.get('serial_number',''), d.get('assigned_user',''),
                d.get('password',''), d.get('purchase_date',''),
                d.get('condition','Good'), d.get('location_code',''),
                d.get('island',''), d.get('phone',''), d.get('notes',''),
                1 if d.get('archived') else 0,
                d.get('archived_at',''), d.get('added_at',''), d.get('updated_at',''),
                d.get('source','added')))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def load_tickets():
    conn = get_db()
    rows = conn.execute('SELECT * FROM tickets').fetchall()
    conn.close()
    return [dict(row) for row in rows]

def save_tickets(tickets):
    conn = get_db()
    try:
        conn.execute('BEGIN')
        conn.execute('DELETE FROM tickets')
        for t in tickets:
            conn.execute('''INSERT INTO tickets
                (id,title,description,category,priority,status,submitted_by,created_at,updated_at)
                VALUES (?,?,?,?,?,?,?,?,?)''', (
                t.get('id'), t.get('title',''), t.get('description',''),
                t.get('category',''), t.get('priority','Medium'),
                t.get('status','Open'), t.get('submitted_by',''),
                t.get('created_at',''), t.get('updated_at','')))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def generate_id(items):
    if not items:
        return 1
    return max(item.get('id', 0) for item in items) + 1

def load_history():
    conn = get_db()
    rows = conn.execute('SELECT * FROM history ORDER BY id DESC').fetchall()
    conn.close()
    result = []
    for row in rows:
        d = dict(row)
        d['changes'] = json.loads(d.get('changes') or '{}')
        d['device_snapshot'] = json.loads(d.get('device_snapshot') or '{}')
        result.append(d)
    return result

def save_history(history):
    conn = get_db()
    try:
        conn.execute('BEGIN')
        conn.execute('DELETE FROM history')
        for h in history:
            conn.execute('''INSERT INTO history
                (id,timestamp,performed_by,action,device_id,device_type,model,
                 serial_number,location_code,island,assigned_user,changes,device_snapshot)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''', (
                h.get('id'), h.get('timestamp',''), h.get('performed_by',''),
                h.get('action',''), h.get('device_id'),
                h.get('device_type',''), h.get('model',''),
                h.get('serial_number',''), h.get('location_code',''),
                h.get('island',''), h.get('assigned_user',''),
                json.dumps(h.get('changes',{})),
                json.dumps(h.get('device_snapshot',{}))))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def log_history(action, device, changes=None):
    conn = get_db()
    conn.execute('''INSERT INTO history
        (timestamp,performed_by,action,device_id,device_type,model,
         serial_number,location_code,island,assigned_user,changes,device_snapshot)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)''', (
        datetime.now().strftime('%m-%d-%Y %H:%M:%S'),
        session.get('current_user', 'Unknown') if session else 'System',
        action,
        device.get('id'),
        device.get('device_type', ''),
        device.get('model', ''),
        device.get('serial_number', ''),
        device.get('location_code', ''),
        device.get('island', ''),
        device.get('assigned_user', ''),
        json.dumps(changes or {}),
        json.dumps(copy.deepcopy(device))
    ))
    conn.commit()
    conn.close()

def load_users():
    conn = get_db()
    rows = conn.execute('SELECT name FROM users ORDER BY id').fetchall()
    conn.close()
    return [row['name'] for row in rows]


def load_device_types():
    conn = get_db()
    rows = conn.execute('SELECT name FROM device_types ORDER BY name').fetchall()
    conn.close()
    types = [row['name'] for row in rows]
    if not types:
        save_device_types(DEVICE_TYPES)
        return DEVICE_TYPES[:]
    return types

def save_device_types(types):
    conn = get_db()
    try:
        conn.execute('BEGIN')
        conn.execute('DELETE FROM device_types')
        for t in types:
            conn.execute('INSERT INTO device_types (name) VALUES (?)', (t,))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def load_locations():
    conn = get_db()
    rows = conn.execute('SELECT code, island FROM locations ORDER BY code').fetchall()
    conn.close()
    return [{'code': row['code'], 'island': row['island']} for row in rows]

def save_locations(locations):
    conn = get_db()
    try:
        conn.execute('BEGIN')
        conn.execute('DELETE FROM locations')
        for loc in locations:
            conn.execute('INSERT INTO locations (code, island) VALUES (?,?)',
                         (loc.get('code',''), loc.get('island','')))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def load_islands():
    conn = get_db()
    rows = conn.execute('SELECT name FROM islands ORDER BY name').fetchall()
    conn.close()
    return [row['name'] for row in rows]

def save_islands(islands):
    conn = get_db()
    try:
        conn.execute('BEGIN')
        conn.execute('DELETE FROM islands')
        for name in islands:
            conn.execute('INSERT INTO islands (name) VALUES (?)', (name,))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

@app.before_request
def require_user():
    allowed = {'select_user', 'add_user', 'delete_user', 'change_password', 'logout', 'static'}
    if request.endpoint not in allowed and 'current_user' not in session:
        return redirect(url_for('select_user'))

@app.route('/select-user', methods=['GET', 'POST'])
def select_user():
    users = load_users()
    conn = get_db()
    login_users = [r['name'] for r in conn.execute(
        "SELECT name FROM users WHERE password_hash IS NOT NULL AND password_hash != '' ORDER BY id"
    ).fetchall()]
    conn.close()
    if request.method == 'POST':
        chosen = request.form.get('user', '').strip()
        password = request.form.get('password', '')
        if not chosen or chosen not in login_users:
            flash('Please select a valid user.', 'error')
            return redirect(url_for('select_user'))
        conn = get_db()
        row = conn.execute('SELECT password_hash FROM users WHERE name = ?', (chosen,)).fetchone()
        conn.close()
        stored_hash = row['password_hash'] if row else ''
        if not check_password_hash(stored_hash, password):
            flash('Incorrect password.', 'error')
            return redirect(url_for('select_user'))
        session['current_user'] = chosen
        return redirect(url_for('index'))
    return render_template('select_user.html', users=users, login_users=login_users)

@app.route('/add-user', methods=['POST'])
def add_user():
    users = load_users()
    new_name = request.form.get('new_user', '').strip()
    password = request.form.get('new_password', '').strip()
    confirm = request.form.get('confirm_password', '').strip()
    if not new_name:
        flash('Please enter a name.', 'error')
    elif not password:
        flash('Please enter a password.', 'error')
    elif password != confirm:
        flash('Passwords do not match.', 'error')
    elif new_name in users:
        flash(f'"{new_name}" already exists.', 'error')
    else:
        conn = get_db()
        conn.execute('INSERT INTO users (name, password_hash) VALUES (?, ?)',
                     (new_name, generate_password_hash(password)))
        conn.commit()
        conn.close()
        flash(f'User "{new_name}" added.', 'success')
    return redirect(url_for('select_user'))

@app.route('/delete-user', methods=['POST'])
def delete_user():
    users = load_users()
    name = request.form.get('user', '').strip()
    password = request.form.get('password', '')
    if name not in users:
        flash('User not found.', 'error')
        return redirect(url_for('select_user'))
    conn = get_db()
    row = conn.execute('SELECT password_hash FROM users WHERE name = ?', (name,)).fetchone()
    stored_hash = row['password_hash'] if row else ''
    if stored_hash and not check_password_hash(stored_hash, password):
        conn.close()
        flash('Incorrect password. User not removed.', 'error')
        return redirect(url_for('select_user'))
    conn.execute('DELETE FROM users WHERE name = ?', (name,))
    conn.commit()
    conn.close()
    if session.get('current_user') == name:
        session.pop('current_user', None)
    flash(f'User "{name}" removed.', 'success')
    return redirect(url_for('select_user'))

@app.route('/change-password', methods=['POST'])
def change_password():
    name = request.form.get('user', '').strip()
    current = request.form.get('current_password', '')
    new_pw = request.form.get('new_password', '').strip()
    confirm = request.form.get('confirm_password', '').strip()
    if not name:
        flash('Please select a user.', 'error')
        return redirect(url_for('select_user'))
    conn = get_db()
    row = conn.execute('SELECT password_hash FROM users WHERE name = ?', (name,)).fetchone()
    if not row:
        conn.close()
        flash('User not found.', 'error')
        return redirect(url_for('select_user'))
    if not check_password_hash(row['password_hash'], current):
        conn.close()
        flash('Current password is incorrect.', 'error')
        return redirect(url_for('select_user'))
    if not new_pw:
        conn.close()
        flash('New password cannot be empty.', 'error')
        return redirect(url_for('select_user'))
    if new_pw == current:
        conn.close()
        flash('New password must be different from your current password.', 'error')
        return redirect(url_for('select_user'))
    if new_pw != confirm:
        conn.close()
        flash('New passwords do not match.', 'error')
        return redirect(url_for('select_user'))
    conn.execute('UPDATE users SET password_hash = ? WHERE name = ?',
                 (generate_password_hash(new_pw), name))
    conn.commit()
    conn.close()
    flash(f'Password updated for "{name}".', 'success')
    return redirect(url_for('select_user'))

@app.route('/logout')
def logout():
    session.pop('current_user', None)
    return redirect(url_for('select_user'))

# Routes
@app.route('/')
def index():
    """Dashboard homepage"""
    inventory = load_inventory()
    tickets = load_tickets()

    selected_island = request.args.get('island', '')
    archive_filter = request.args.get('archive_filter', 'active')

    # Build island list for filter dropdown
    all_islands = _get_all_islands(inventory)

    # Filter by archive status
    if archive_filter == 'archived':
        filtered = [d for d in inventory if d.get('archived')]
    elif archive_filter == 'all':
        filtered = inventory[:]
    else:
        filtered = [d for d in inventory if not d.get('archived')]

    # Filter by island if selected
    if selected_island:
        filtered = [d for d in filtered if d.get('island', '').strip() == selected_island]

    # Calculate statistics
    total_devices = len(filtered)

    # Devices by location (sorted alphabetically)
    devices_by_location_raw = {}
    for device in filtered:
        location = device.get('location_code', 'Unknown')
        devices_by_location_raw[location] = devices_by_location_raw.get(location, 0) + 1
    devices_by_location = dict(sorted(devices_by_location_raw.items(), key=lambda x: x[0].lower()))

    # Open tickets
    open_tickets = [t for t in tickets if t.get('status') in ['Open', 'In Progress']]
    total_open_tickets = len(open_tickets)

    # Tickets by category
    tickets_by_category = {}
    for ticket in tickets:
        category = ticket.get('category', 'Other')
        tickets_by_category[category] = tickets_by_category.get(category, 0) + 1

    # Recent tickets (last 5)
    recent_tickets = sorted(tickets, key=lambda x: x.get('created_at', ''), reverse=True)[:5]

    return render_template('index.html',
                         total_devices=total_devices,
                         devices_by_location=devices_by_location,
                         total_open_tickets=total_open_tickets,
                         all_islands=all_islands,
                         selected_island=selected_island,
                         archive_filter=archive_filter,
                         tickets_by_category=tickets_by_category,
                         recent_tickets=recent_tickets)

@app.route('/islands/add', methods=['POST'])
def add_island():
    islands = load_islands()
    name = request.form.get('new_island', '').strip()
    if not name:
        flash('Please enter an island name.', 'error')
    elif name in islands:
        flash(f'"{name}" already exists.', 'error')
    else:
        islands.append(name)
        save_islands(islands)
        flash(f'Island "{name}" added.', 'success')
    return redirect(url_for('inventory_list'))

@app.route('/islands/delete', methods=['POST'])
def delete_island():
    islands = load_islands()
    name = request.form.get('island', '').strip()
    confirm = request.form.get('confirm_delete', 'false')
    if name in islands:
        inventory = load_inventory()
        in_use = sum(1 for d in inventory if d.get('island', '') == name)
        if in_use > 0 and confirm != 'true':
            flash(f'WARNING: {in_use} device(s) are assigned to island "{name}". '
                  f'Removing it won\'t change those devices, but "{name}" will disappear from filters. '
                  f'Click Remove again to confirm.', 'error')
            return redirect(url_for('inventory_list'))
        islands.remove(name)
        save_islands(islands)
        flash(f'Island "{name}" removed.', 'success')
    else:
        flash('Island not found.', 'error')
    return redirect(url_for('inventory_list'))

@app.route('/locations/add', methods=['POST'])
def add_location():
    locations = load_locations()
    code = request.form.get('new_location', '').strip()
    island = request.form.get('location_island', '').strip()
    if not code:
        flash('Please enter a location name.', 'error')
    elif any(l['code'] == code for l in locations):
        flash(f'"{code}" already exists.', 'error')
    else:
        locations.append({'code': code, 'island': island})
        save_locations(locations)
        log_history('Location Added', {
            'id': None, 'device_type': '', 'model': f'Location "{code}" (Island: {island or "—"})',
            'serial_number': '', 'location_code': code, 'island': island, 'assigned_user': ''
        })
        flash(f'Location "{code}" added.', 'success')
    return redirect(url_for('inventory_list'))

@app.route('/locations/delete', methods=['POST'])
def delete_location():
    locations = load_locations()
    code = request.form.get('location', '').strip()
    if not any(l['code'] == code for l in locations):
        flash('Location not found.', 'error')
        return redirect(url_for('inventory_list'))
    inventory = load_inventory()
    in_use = sum(1 for d in inventory if d.get('location_code', '') == code)
    confirm = request.form.get('confirm_delete', 'false')
    if in_use > 0 and confirm != 'true':
        flash(f'WARNING: {in_use} device(s) use location "{code}". Submit again to confirm removal.', 'error')
        return redirect(url_for('inventory_list'))
    locations = [l for l in locations if l['code'] != code]
    save_locations(locations)
    log_history('Location Deleted', {
        'id': None, 'device_type': '', 'model': f'Location "{code}"',
        'serial_number': '', 'location_code': code, 'island': '', 'assigned_user': ''
    })
    flash(f'Location "{code}" removed.', 'success')
    return redirect(url_for('inventory_list'))

@app.route('/device-types/add', methods=['POST'])
def add_device_type():
    types = load_device_types()
    new_type = request.form.get('new_type', '').strip()
    if not new_type:
        flash('Please enter a device type name.', 'error')
    elif new_type in types:
        flash(f'"{new_type}" already exists.', 'error')
    else:
        types.append(new_type)
        save_device_types(types)
        flash(f'Device type "{new_type}" added.', 'success')
    return redirect(url_for('inventory_list'))

@app.route('/device-types/delete', methods=['POST'])
def delete_device_type():
    types = load_device_types()
    name = request.form.get('type', '').strip()
    confirm = request.form.get('confirm_delete', 'false')
    if name in types:
        inventory = load_inventory()
        in_use = sum(1 for d in inventory if d.get('device_type', '') == name)
        if in_use > 0 and confirm != 'true':
            flash(f'WARNING: {in_use} device(s) are currently using type "{name}". '
                  f'The type will be removed from the dropdown but those devices will keep their existing type label. '
                  f'Click Remove again to confirm.', 'error')
            return redirect(url_for('inventory_list') + '#deviceTypesModal')
        types.remove(name)
        save_device_types(types)
        flash(f'Device type "{name}" removed.', 'success')
    else:
        flash('Device type not found.', 'error')
    return redirect(url_for('inventory_list'))

@app.route('/inventory')
def inventory_list():
    """Display all inventory items"""
    inventory = load_inventory()

    # Get filter parameters
    search = request.args.get('search', '').lower()
    device_type = request.args.get('device_type', '')
    location = request.args.get('location', '')
    selected_island = request.args.get('island', '')
    archive_filter = request.args.get('archive_filter', 'active')  # active | archived | all

    # Filter by archive status first
    if archive_filter == 'archived':
        filtered_inventory = [item for item in inventory if item.get('archived')]
    elif archive_filter == 'all':
        filtered_inventory = inventory[:]
    else:
        filtered_inventory = [item for item in inventory if not item.get('archived')]

    if selected_island:
        filtered_inventory = [item for item in filtered_inventory if item.get('island', '').strip() == selected_island]

    if search:
        filtered_inventory = [
            item for item in filtered_inventory
            if search in item.get('model', '').lower() or
               search in item.get('serial_number', '').lower() or
               search in item.get('assigned_user', '').lower() or
               search in item.get('location_code', '').lower()
        ]

    if device_type:
        filtered_inventory = [item for item in filtered_inventory if item.get('device_type') == device_type]

    if location:
        filtered_inventory = [item for item in filtered_inventory if item.get('location_code') == location]

    # Add location names to items
    for item in filtered_inventory:
        item['location_name'] = LOCATION_CODES.get(item.get('location_code', ''), item.get('location_code', 'Unknown'))

    locations_by_island = _build_locations_by_island(inventory)
    all_islands = _get_all_islands(inventory)

    # Build device type list from saved types + any in inventory
    saved_types = load_device_types()
    all_types_in_inventory = sorted({item.get('device_type', '').strip() for item in inventory if item.get('device_type', '').strip()})
    device_types_for_filter = sorted(set(saved_types) | set(all_types_in_inventory))

    saved_locations = load_locations()
    saved_islands = sorted(load_islands())

    return render_template('inventory.html',
                         inventory=filtered_inventory,
                         full_inventory=inventory,
                         device_types=device_types_for_filter,
                         saved_device_types=sorted(saved_types),
                         saved_locations=saved_locations,
                         saved_islands=saved_islands,
                         location_codes=LOCATION_CODES,
                         locations_by_island=locations_by_island,
                         all_islands=all_islands,
                         search=search,
                         selected_device_type=device_type,
                         selected_location=location,
                         selected_island=selected_island,
                         archive_filter=archive_filter)

@app.route('/inventory/add', methods=['GET', 'POST'])
def add_device():
    """Add a new device to inventory"""
    if request.method == 'POST':
        # Validate required fields
        device_type = request.form.get('device_type')
        model = request.form.get('model')
        serial_number = request.form.get('serial_number')
        location_code = request.form.get('location_code')
        
        if not all([device_type, model, serial_number, location_code]):
            flash('Please fill in all required fields.', 'error')
            return redirect(url_for('add_device'))
        
        # Create new device
        inventory = load_inventory()
        
        new_device = {
            'id': generate_id(inventory),
            'device_type': device_type,
            'model': model,
            'serial_number': serial_number,
            'assigned_user': request.form.get('assigned_user', ''),
            'password': request.form.get('password', ''),
            'purchase_date': _normalize_date(request.form.get('purchase_date', '')),
            'condition': request.form.get('condition', 'Good'),
            'location_code': location_code,
            'island': _island_for_location(location_code),
            'phone': request.form.get('phone', ''),
            'notes': request.form.get('notes', ''),
            'archived': False,
            'added_at': datetime.now().strftime('%m-%d-%Y %H:%M:%S'),
            'updated_at': '',
            'source': 'added'
        }

        inventory.append(new_device)
        save_inventory(inventory)
        log_history('Added', new_device)
        flash(f'Device {model} added successfully!', 'success')
        return redirect(url_for('inventory_list'))
    
    inventory = load_inventory()
    all_islands = _get_all_islands(inventory)
    return render_template('add_device.html',
                         device_types=load_device_types(),
                         locations_by_island=_build_locations_by_island(inventory),
                         all_islands=all_islands,
                         location_codes=LOCATION_CODES,
                         conditions=DEVICE_CONDITIONS)

@app.route('/inventory/edit/<int:device_id>', methods=['GET', 'POST'])
def edit_device(device_id):
    """Edit an existing device"""
    inventory = load_inventory()
    device = next((item for item in inventory if item['id'] == device_id), None)
    
    if not device:
        flash('Device not found.', 'error')
        return redirect(url_for('inventory_list'))
    
    if request.method == 'POST':
        if not all([request.form.get('device_type'), request.form.get('model'),
                    request.form.get('serial_number'), request.form.get('location_code')]):
            flash('Please fill in all required fields.', 'error')
            return redirect(url_for('edit_device', device_id=device_id))
        # Track changes for history
        track_fields = ['device_type', 'model', 'serial_number', 'assigned_user', 'password', 'location_code', 'phone', 'purchase_date', 'condition', 'notes']
        changes = {f: {'from': device.get(f, ''), 'to': request.form.get(f, device.get(f, ''))} for f in track_fields if str(device.get(f, '')) != str(request.form.get(f, device.get(f, '')))}
        # Island is auto-derived — track separately since it has no form field
        old_island = device.get('island', '')
        new_island = _island_for_location(request.form.get('location_code', ''))
        if old_island != new_island:
            changes['island'] = {'from': old_island, 'to': new_island}
        # Update device
        device['device_type'] = request.form.get('device_type')
        device['model'] = request.form.get('model')
        device['serial_number'] = request.form.get('serial_number')
        device['assigned_user'] = request.form.get('assigned_user', '')
        device['password'] = request.form.get('password', '')
        device['purchase_date'] = _normalize_date(request.form.get('purchase_date', ''))
        condition = request.form.get('condition', 'Good')
        device['condition'] = condition if condition in DEVICE_CONDITIONS else 'Good'
        device['location_code'] = request.form.get('location_code')
        device['island'] = _island_for_location(request.form.get('location_code', ''))
        device['phone'] = request.form.get('phone', '')
        device['notes'] = request.form.get('notes', '')
        device['updated_at'] = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
        
        save_inventory(inventory)
        log_history('Edited', device, changes)
        flash(f'Device {device["model"]} updated successfully!', 'success')
        return redirect(url_for('inventory_list'))
    
    all_islands = _get_all_islands(inventory)
    device['location_name'] = LOCATION_CODES.get(device.get('location_code', ''), device.get('location_code', 'Unknown'))
    return render_template('edit_device.html',
                         device=device,
                         device_types=load_device_types(),
                         locations_by_island=_build_locations_by_island(inventory),
                         all_islands=all_islands,
                         location_codes=LOCATION_CODES,
                         conditions=DEVICE_CONDITIONS)

@app.route('/inventory/delete/<int:device_id>', methods=['POST'])
def delete_device(device_id):
    """Delete a device from inventory"""
    inventory = load_inventory()
    device = next((item for item in inventory if item['id'] == device_id), None)
    if device:
        log_history('Deleted', device)
    inventory = [item for item in inventory if item['id'] != device_id]
    save_inventory(inventory)

    flash('Device deleted successfully!', 'success')
    return redirect(url_for('inventory_list'))

@app.route('/inventory/archive/<int:device_id>', methods=['POST'])
def archive_device(device_id):
    """Archive a device (mark as predecessor/retired)"""
    inventory = load_inventory()
    device = next((item for item in inventory if item['id'] == device_id), None)
    if device:
        device['archived'] = True
        device['archived_at'] = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
        save_inventory(inventory)
        log_history('Archived', device)
        flash(f'Device {device["model"]} archived.', 'success')
    else:
        flash('Device not found.', 'error')
    return redirect(request.referrer or url_for('inventory_list'))

@app.route('/inventory/mark-tbd/<int:device_id>', methods=['POST'])
def mark_tbd(device_id):
    """Keep device active but set assigned user to TBD"""
    inventory = load_inventory()
    device = next((item for item in inventory if item['id'] == device_id), None)
    if device:
        # Archive the original to preserve history
        device['archived'] = True
        device['archived_at'] = datetime.now().strftime('%m-%d-%Y %H:%M:%S')

        # Create a fresh active copy with TBD fields
        new_device = copy.deepcopy(device)
        new_device['id'] = generate_id(inventory)
        new_device['assigned_user'] = 'TBD'
        new_device['password'] = 'TBD'
        new_device['phone'] = 'TBD'
        new_device['archived'] = False
        new_device.pop('archived_at', None)
        new_device['added_at'] = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
        new_device.pop('updated_at', None)

        inventory.append(new_device)
        save_inventory(inventory)
        log_history('Archived (TBD copy created)', device)
        log_history('Added (TBD copy)', new_device)
        flash(f'Original {device["model"]} archived. Active copy created with TBD fields.', 'success')
    else:
        flash('Device not found.', 'error')
    return redirect(request.referrer or url_for('inventory_list'))

@app.route('/inventory/unarchive/<int:device_id>', methods=['POST'])
def unarchive_device(device_id):
    """Restore an archived device to active inventory"""
    inventory = load_inventory()
    device = next((item for item in inventory if item['id'] == device_id), None)
    if device:
        device['archived'] = False
        device.pop('archived_at', None)
        save_inventory(inventory)
        log_history('Restored', device)
        flash(f'Device {device["model"]} restored to active inventory.', 'success')
    else:
        flash('Device not found.', 'error')
    return redirect(request.referrer or url_for('inventory_list'))

@app.route('/inventory/history/undo/<int:history_id>', methods=['POST'])
def undo_history(history_id):
    history = load_history()
    inventory = load_inventory()
    entry = next((e for e in history if e['id'] == history_id), None)

    if not entry:
        flash('History entry not found.', 'error')
        return redirect(url_for('inventory_history'))

    action = entry.get('action', '')
    snapshot = entry.get('device_snapshot', {})
    device_id = entry.get('device_id')
    undone_by = session.get('current_user', 'Unknown')

    # --- TBD pair: undo both archive + TBD copy together ---
    if action in ('Archived (TBD copy created)', 'Added (TBD copy)'):
        ts = entry.get('timestamp')
        user = entry.get('performed_by')
        pair = [e for e in history if e.get('timestamp') == ts and
                e.get('performed_by') == user and
                e.get('action') in ('Archived (TBD copy created)', 'Added (TBD copy)')]

        archive_entry = next((e for e in pair if e['action'] == 'Archived (TBD copy created)'), None)
        tbd_entry     = next((e for e in pair if e['action'] == 'Added (TBD copy)'), None)

        if archive_entry and tbd_entry:
            # Restore the original (un-archive it)
            original = next((d for d in inventory if d['id'] == archive_entry['device_id']), None)
            if original:
                original['archived'] = False
                original.pop('archived_at', None)
            # Delete the TBD copy
            inventory = [d for d in inventory if d['id'] != tbd_entry['device_id']]
            save_inventory(inventory)
            # Remove both history entries and log the undo
            history = [e for e in history if e['id'] not in (archive_entry['id'], tbd_entry['id'])]
            undo_entry = {
                'id': generate_id(history),
                'timestamp': datetime.now().strftime('%m-%d-%Y %H:%M:%S'),
                'performed_by': undone_by,
                'action': 'Undone: Archive+TBD',
                'device_id': archive_entry['device_id'],
                'device_type': archive_entry.get('device_type', ''),
                'model': archive_entry.get('model', ''),
                'serial_number': archive_entry.get('serial_number', ''),
                'location_code': archive_entry.get('location_code', ''),
                'island': archive_entry.get('island', ''),
                'assigned_user': archive_entry.get('assigned_user', ''),
                'changes': {},
                'device_snapshot': {}
            }
            history.insert(0, undo_entry)
            save_history(history)
            flash(f'Undone: archive and TBD copy for {archive_entry.get("model")} reversed.', 'success')
        else:
            flash('Could not find matching TBD pair to undo.', 'error')

    # --- Added: delete the device ---
    elif action == 'Added':
        inventory = [d for d in inventory if d['id'] != device_id]
        save_inventory(inventory)
        history = [e for e in history if e['id'] != history_id]
        history.insert(0, {'id': generate_id(history), 'timestamp': datetime.now().strftime('%m-%d-%Y %H:%M:%S'),
            'performed_by': undone_by, 'action': f'Undone: Added',
            'device_id': device_id, 'device_type': entry.get('device_type',''), 'model': entry.get('model',''),
            'serial_number': entry.get('serial_number',''), 'location_code': entry.get('location_code',''),
            'island': entry.get('island',''), 'assigned_user': entry.get('assigned_user',''),
            'changes': {}, 'device_snapshot': {}})
        save_history(history)
        flash(f'Undone: removed added device {entry.get("model")}.', 'success')

    # --- Edited: revert fields to their previous values ---
    elif action == 'Edited':
        device = next((d for d in inventory if d['id'] == device_id), None)
        if device and entry.get('changes'):
            for field, diff in entry['changes'].items():
                if 'from' in diff:
                    device[field] = diff['from']
            save_inventory(inventory)
            history = [e for e in history if e['id'] != history_id]
            history.insert(0, {'id': generate_id(history), 'timestamp': datetime.now().strftime('%m-%d-%Y %H:%M:%S'),
                'performed_by': undone_by, 'action': 'Undone: Edited',
                'device_id': device_id, 'device_type': device.get('device_type',''), 'model': device.get('model',''),
                'serial_number': device.get('serial_number',''), 'location_code': device.get('location_code',''),
                'island': device.get('island',''), 'assigned_user': device.get('assigned_user',''),
                'changes': {}, 'device_snapshot': {}})
            save_history(history)
            flash(f'Undone: reverted edit on {entry.get("model")}.', 'success')
        else:
            flash('No changes to revert or device not found.', 'error')

    # --- Deleted: restore from snapshot ---
    elif action == 'Deleted':
        if snapshot:
            snapshot['id'] = generate_id(inventory)
            snapshot['archived'] = False
            inventory.append(snapshot)
            save_inventory(inventory)
            history = [e for e in history if e['id'] != history_id]
            history.insert(0, {'id': generate_id(history), 'timestamp': datetime.now().strftime('%m-%d-%Y %H:%M:%S'),
                'performed_by': undone_by, 'action': 'Undone: Deleted',
                'device_id': snapshot['id'], 'device_type': snapshot.get('device_type',''), 'model': snapshot.get('model',''),
                'serial_number': snapshot.get('serial_number',''), 'location_code': snapshot.get('location_code',''),
                'island': snapshot.get('island',''), 'assigned_user': snapshot.get('assigned_user',''),
                'changes': {}, 'device_snapshot': {}})
            save_history(history)
            flash(f'Undone: restored deleted device {entry.get("model")}.', 'success')
        else:
            flash('No snapshot available — this entry was created before undo support was added.', 'error')

    # --- Archived: restore to active ---
    elif action == 'Archived':
        device = next((d for d in inventory if d['id'] == device_id), None)
        if device:
            device['archived'] = False
            device.pop('archived_at', None)
            save_inventory(inventory)
            history.insert(0, {'id': generate_id(history), 'timestamp': datetime.now().strftime('%m-%d-%Y %H:%M:%S'),
                'performed_by': undone_by, 'action': 'Undone: Archived',
                'device_id': device_id, 'device_type': entry.get('device_type',''), 'model': entry.get('model',''),
                'serial_number': entry.get('serial_number',''), 'location_code': entry.get('location_code',''),
                'island': entry.get('island',''), 'assigned_user': entry.get('assigned_user',''),
                'changes': {}, 'device_snapshot': {}})
            save_history(history)
            flash(f'Undone: {entry.get("model")} restored to active inventory.', 'success')
        else:
            flash('Device not found in inventory.', 'error')

    # --- Restored: re-archive ---
    elif action == 'Restored':
        device = next((d for d in inventory if d['id'] == device_id), None)
        if device:
            device['archived'] = True
            device['archived_at'] = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
            save_inventory(inventory)
            history = [e for e in history if e['id'] != history_id]
            history.insert(0, {'id': generate_id(history), 'timestamp': datetime.now().strftime('%m-%d-%Y %H:%M:%S'),
                'performed_by': undone_by, 'action': 'Undone: Restored',
                'device_id': device_id, 'device_type': entry.get('device_type',''), 'model': entry.get('model',''),
                'serial_number': entry.get('serial_number',''), 'location_code': entry.get('location_code',''),
                'island': entry.get('island',''), 'assigned_user': entry.get('assigned_user',''),
                'changes': {}, 'device_snapshot': {}})
            save_history(history)
            flash(f'Undone: {entry.get("model")} re-archived.', 'success')
        else:
            flash('Device not found.', 'error')

    else:
        flash(f'Undo not supported for action: {action}', 'error')

    return redirect(url_for('inventory_history'))


@app.route('/inventory/history')
def inventory_history():
    """Display inventory change history"""
    history = load_history()
    search = request.args.get('search', '').lower()
    action_filter = request.args.get('action', '')
    if search:
        history = [e for e in history if search in e.get('model', '').lower() or
                   search in e.get('serial_number', '').lower() or
                   search in e.get('location_code', '').lower() or
                   search in e.get('assigned_user', '').lower()]
    if action_filter:
        history = [e for e in history if e.get('action', '') == action_filter]
    all_actions = sorted({e.get('action', '') for e in load_history()})
    return render_template('inventory_history.html',
                           history=history,
                           all_actions=all_actions,
                           search=search,
                           selected_action=action_filter)

@app.route('/tickets')
def tickets_list():
    """Display all tickets"""
    tickets = load_tickets()
    
    # Get filter parameters
    search = request.args.get('search', '').lower()
    status = request.args.get('status', '')
    category = request.args.get('category', '')
    priority = request.args.get('priority', '')
    
    # Filter tickets
    filtered_tickets = tickets
    
    if search:
        filtered_tickets = [
            ticket for ticket in filtered_tickets
            if search in ticket.get('title', '').lower() or
               search in ticket.get('description', '').lower() or
               search in ticket.get('submitted_by', '').lower()
        ]
    
    if status:
        filtered_tickets = [ticket for ticket in filtered_tickets if ticket.get('status') == status]
    
    if category:
        filtered_tickets = [ticket for ticket in filtered_tickets if ticket.get('category') == category]
    
    if priority:
        filtered_tickets = [ticket for ticket in filtered_tickets if ticket.get('priority') == priority]
    
    # Sort by created date (newest first)
    filtered_tickets = sorted(filtered_tickets, key=lambda x: x.get('created_at', ''), reverse=True)
    
    return render_template('tickets.html',
                         tickets=filtered_tickets,
                         ticket_statuses=TICKET_STATUSES,
                         ticket_categories=TICKET_CATEGORIES,
                         priority_levels=PRIORITY_LEVELS,
                         search=search,
                         selected_status=status,
                         selected_category=category,
                         selected_priority=priority)

@app.route('/tickets/add', methods=['GET', 'POST'])
def add_ticket():
    """Submit a new support ticket"""
    if request.method == 'POST':
        # Validate required fields
        title = request.form.get('title')
        category = request.form.get('category')
        submitted_by = request.form.get('submitted_by')
        
        if not all([title, category, submitted_by]):
            flash('Please fill in all required fields.', 'error')
            return redirect(url_for('add_ticket'))
        
        # Create new ticket
        tickets = load_tickets()
        
        new_ticket = {
            'id': generate_id(tickets),
            'title': title,
            'description': request.form.get('description', ''),
            'category': category,
            'priority': request.form.get('priority', 'Medium'),
            'status': 'Open',
            'submitted_by': submitted_by,
            'created_at': datetime.now().strftime('%m-%d-%Y %H:%M:%S'),
            'updated_at': datetime.now().strftime('%m-%d-%Y %H:%M:%S')
        }
        
        tickets.append(new_ticket)
        save_tickets(tickets)
        
        flash(f'Ticket #{new_ticket["id"]} created successfully!', 'success')
        return redirect(url_for('tickets_list'))
    
    return render_template('add_ticket.html',
                         ticket_categories=TICKET_CATEGORIES,
                         priority_levels=PRIORITY_LEVELS)

@app.route('/tickets/view/<int:ticket_id>')
def view_ticket(ticket_id):
    """View ticket details"""
    tickets = load_tickets()
    ticket = next((t for t in tickets if t['id'] == ticket_id), None)
    
    if not ticket:
        flash('Ticket not found.', 'error')
        return redirect(url_for('tickets_list'))
    
    return render_template('view_ticket.html',
                         ticket=ticket,
                         ticket_statuses=TICKET_STATUSES)

@app.route('/tickets/update/<int:ticket_id>', methods=['POST'])
def update_ticket_status(ticket_id):
    """Update ticket status"""
    tickets = load_tickets()
    ticket = next((t for t in tickets if t['id'] == ticket_id), None)
    
    if ticket:
        new_status = request.form.get('status', '').strip()
        if new_status not in TICKET_STATUSES:
            flash('Invalid status value.', 'error')
            return redirect(url_for('view_ticket', ticket_id=ticket_id))
        ticket['status'] = new_status
        ticket['updated_at'] = datetime.now().strftime('%m-%d-%Y %H:%M:%S')
        save_tickets(tickets)
        flash(f'Ticket #{ticket_id} status updated to {new_status}!', 'success')
    else:
        flash('Ticket not found.', 'error')
    
    return redirect(url_for('view_ticket', ticket_id=ticket_id))

@app.route('/tickets/delete/<int:ticket_id>', methods=['POST'])
def delete_ticket(ticket_id):
    """Delete a ticket"""
    tickets = load_tickets()
    tickets = [t for t in tickets if t['id'] != ticket_id]
    save_tickets(tickets)
    
    flash('Ticket deleted successfully!', 'success')
    return redirect(url_for('tickets_list'))

CSV_COLUMNS = ['device_type', 'model', 'serial_number', 'assigned_user', 'password', 'purchase_date', 'condition', 'location_code', 'island', 'phone', 'notes']

def _normalize_date(date_str):
    """Parse common date formats and return MM-DD-YYYY. Returns original string if unparseable."""
    if not date_str or not date_str.strip():
        return date_str
    s = date_str.strip()
    for fmt in (
        '%m-%d-%Y',    # already correct: 03-07-2030
        '%m/%d/%y',    # 3/7/30  →  2-digit year (00-68 = 2000-2068)
        '%m/%d/%Y',    # 3/7/2030
        '%Y-%m-%d',    # 2030-03-07  (ISO / Excel)
        '%m-%d-%y',    # 03-07-30
        '%B %d, %Y',   # January 15, 2030
        '%b %d, %Y',   # Jan 15, 2030
        '%B %d %Y',    # January 15 2030
        '%b %d %Y',    # Jan 15 2030
    ):
        try:
            return datetime.strptime(s, fmt).strftime('%m-%d-%Y')
        except ValueError:
            continue
    return s  # unrecognised — keep as-is

def _island_for_location(location_code):
    """Return the island for a given location code.
    Checks locations table first, then falls back to existing inventory data."""
    conn = get_db()
    row = conn.execute('SELECT island FROM locations WHERE code = ?', (location_code,)).fetchone()
    if row and row['island']:
        conn.close()
        return row['island']
    inv_row = conn.execute(
        "SELECT island FROM inventory WHERE location_code = ? AND island != '' LIMIT 1",
        (location_code,)
    ).fetchone()
    conn.close()
    return inv_row['island'] if inv_row else ''

def _build_locations_by_island(inventory=None):
    if inventory is None:
        inventory = load_inventory()
    loc_to_island = {}
    # Start with manually stored locations as base
    for entry in load_locations():
        code = entry.get('code', '').strip()
        island = entry.get('island', '').strip()
        if code:
            loc_to_island[code] = island or 'Other'
    # Inventory data always overrides stored locations when it has a non-empty island
    for item in inventory:
        loc = item.get('location_code', '').strip()
        island = item.get('island', '').strip()
        if loc:
            if island:
                # Non-empty island from inventory always wins
                loc_to_island[loc] = island
            elif loc not in loc_to_island:
                # Only fall back to 'Other' if not already mapped
                loc_to_island[loc] = 'Other'
    result = {}
    for loc, island in sorted(loc_to_island.items()):
        result.setdefault(island, []).append(loc)
    return dict(sorted(result.items()))

def _get_all_islands(inventory=None):
    """Build full island list from stored islands + inventory data."""
    if inventory is None:
        inventory = load_inventory()
    islands = set(load_islands())
    for item in inventory:
        isl = item.get('island', '').strip()
        if isl:
            islands.add(isl)
    return sorted(islands)
CSV_REQUIRED = {'device_type', 'location_code'}

COLUMN_ALIASES = {
    'device type': 'device_type',
    'devicetype': 'device_type',
    'type': 'device_type',
    'computername/servicetag': 'serial_number',
    'computername': 'serial_number',
    'service tag': 'serial_number',
    'servicetag': 'serial_number',
    'serial': 'serial_number',
    'serial #': 'serial_number',
    'serial no': 'serial_number',
    'location': 'location_code',
    'location code': 'location_code',
    'username': 'assigned_user',
    'user': 'assigned_user',
    'user name': 'assigned_user',
    'assigned user': 'assigned_user',
    'notes:': 'notes',
    'note': 'notes',
    'purchase date': 'purchase_date',
    'warranty end date': 'purchase_date',
    'warranty': 'purchase_date',
}

def _normalize_col(name):
    name = str(name).strip().lower()
    return COLUMN_ALIASES.get(name, name)

def _xlsx_cell_str(val):
    if val is None:
        return ''
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    if hasattr(val, 'strftime'):
        return val.strftime('%m-%d-%Y')
    return str(val).strip()

def _process_rows(rows):
    """Import a list of dicts (with CSV_COLUMNS keys) into inventory. Returns (imported, skipped, skip_reasons)."""
    inventory = load_inventory()
    saved_types = load_device_types()
    imported, skipped = 0, 0
    skip_reasons = {}

    for row in rows:
        missing = [f for f in CSV_REQUIRED if not str(row.get(f, '')).strip()]
        if missing:
            reason = f'missing: {", ".join(missing)}'
            skip_reasons[reason] = skip_reasons.get(reason, 0) + 1
            skipped += 1
            continue

        # Normalize device type: check saved types first, then aliases, then import as-is
        device_type = str(row.get('device_type', '')).strip()
        matched_type = next((dt for dt in saved_types if dt.lower() == device_type.lower()), None)
        if not matched_type:
            matched_type = DEVICE_TYPE_ALIASES.get(device_type.lower())
        device_type = matched_type or device_type  # fall back to raw value

        # If the resolved type isn't in saved types, add it automatically
        if device_type and device_type not in saved_types:
            saved_types.append(device_type)
            save_device_types(saved_types)

        condition = str(row.get('condition', '')).strip()
        if condition not in DEVICE_CONDITIONS:
            condition = 'Good'

        new_device = {
            'id': generate_id(inventory),
            'device_type': device_type,
            'model': str(row.get('model', '')).strip() or '-',
            'serial_number': str(row.get('serial_number', '')).strip() or 'n/a',
            'assigned_user': (str(row.get('assigned_user', '')).strip() or
                              str(row.get('contact', '')).strip()),
            'password': str(row.get('password', '')).strip(),
            'purchase_date': _normalize_date(str(row.get('purchase_date', '')).strip()),
            'condition': condition,
            'location_code': str(row.get('location_code', '')).strip(),
            'island': str(row.get('island', '')).strip() or _island_for_location(str(row.get('location_code', '')).strip()),
            'phone': str(row.get('phone', '')).strip(),
            'notes': str(row.get('notes', '')).strip(),
            'archived': False,
            'added_at': datetime.now().strftime('%m-%d-%Y %H:%M:%S'),
            'updated_at': '',
            'source': 'imported'
        }
        inventory.append(new_device)
        imported += 1

    save_inventory(inventory)
    return imported, skipped, skip_reasons


@app.route('/inventory/import', methods=['GET', 'POST'])
def import_inventory():
    if request.method == 'POST':
        # Phase 2: user selected a sheet from an xlsx
        if request.form.get('sheet_name'):
            sheet_name = request.form.get('sheet_name')
            temp_path = session.pop('xlsx_temp_path', '')

            if not temp_path or not os.path.exists(temp_path):
                flash('Upload session expired. Please upload the file again.', 'error')
                return redirect(url_for('import_inventory'))

            try:
                wb = openpyxl.load_workbook(temp_path)
                ws = wb[sheet_name]
                headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
                header_map = {_normalize_col(h): i for i, h in enumerate(headers) if h is not None}

                rows = []
                for row in ws.iter_rows(min_row=2):
                    if all(cell.value is None for cell in row):
                        continue
                    rows.append({
                        col: (_xlsx_cell_str(row[header_map[col]].value) if col in header_map else '')
                        for col in CSV_COLUMNS + ['contact']
                    })
                wb.close()
            except Exception as e:
                flash(f'Failed to read file: {e}', 'error')
                return redirect(url_for('import_inventory'))
            finally:
                if os.path.exists(temp_path):
                    os.remove(temp_path)

            imported, skipped, skip_reasons = _process_rows(rows)

            if imported:
                log_history('Imported', {
                    'id': None, 'device_type': '', 'model': f'{imported} device(s) from "{sheet_name}"',
                    'serial_number': '', 'location_code': '', 'island': '', 'assigned_user': ''
                })
                msg = f'Successfully imported {imported} device(s) from "{sheet_name}".'
                if skipped:
                    reasons = '; '.join(f'{count}x {reason}' for reason, count in skip_reasons.items())
                    msg += f' Skipped {skipped} row(s): {reasons}.'
                flash(msg, 'success')
            else:
                detected = ', '.join(header_map.keys()) or 'none'
                reasons = '; '.join(f'{count}x {reason}' for reason, count in skip_reasons.items())
                flash(f'No devices imported from "{sheet_name}". {skipped} row(s) skipped — {reasons}. Columns detected: [{detected}].', 'error')

            return redirect(url_for('inventory_list'))

        # Phase 1: file upload
        file = request.files.get('csv_file')
        if not file or not file.filename:
            flash('Please select a file.', 'error')
            return redirect(url_for('import_inventory'))

        filename = file.filename.lower()

        if filename.endswith('.csv'):
            raw = file.stream.read()
            try:
                content = raw.decode('utf-8-sig')
            except UnicodeDecodeError:
                content = raw.decode('latin-1')
            stream = io.StringIO(content)
            rows = [{_normalize_col(k): v for k, v in row.items()} for row in csv.DictReader(stream)]
            imported, skipped, skip_reasons = _process_rows(rows)

            if imported:
                log_history('Imported', {
                    'id': None, 'device_type': '', 'model': f'{imported} device(s) from "{file.filename}"',
                    'serial_number': '', 'location_code': '', 'island': '', 'assigned_user': ''
                })
                msg = f'Successfully imported {imported} device(s).'
                if skipped:
                    reasons = '; '.join(f'{count}x {reason}' for reason, count in skip_reasons.items())
                    msg += f' Skipped {skipped} row(s): {reasons}.'
                flash(msg, 'success')
            else:
                reasons = '; '.join(f'{count}x {reason}' for reason, count in skip_reasons.items())
                flash(f'No devices imported. {skipped} row(s) skipped — {reasons}.', 'error')

            return redirect(url_for('inventory_list'))

        elif filename.endswith('.xlsx'):
            temp_fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
            os.close(temp_fd)
            file.save(temp_path)

            try:
                wb = openpyxl.load_workbook(temp_path, read_only=True)
                sheets = wb.sheetnames
                wb.close()
            except Exception as e:
                os.remove(temp_path)
                flash(f'Failed to read Excel file: {e}', 'error')
                return redirect(url_for('import_inventory'))

            session['xlsx_temp_path'] = temp_path
            return render_template('import_inventory.html', sheets=sheets)

        else:
            flash('Please upload a .csv or .xlsx file.', 'error')
            return redirect(url_for('import_inventory'))

    return render_template('import_inventory.html')


@app.route('/inventory/import/template')
def download_csv_template():
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=CSV_COLUMNS)
    writer.writeheader()
    writer.writerow({
        'device_type': 'Laptop',
        'model': 'Dell XPS 15',
        'serial_number': 'SN123456',
        'assigned_user': 'Jane Doe',
        'purchase_date': '2024-01-15',
        'condition': 'Good',
        'location_code': '101',
        'notes': 'Example row — delete before importing'
    })
    output.seek(0)
    return Response(
        output.getvalue(),
        mimetype='text/csv',
        headers={'Content-Disposition': 'attachment; filename=inventory_template.csv'}
    )


@app.route('/inventory/export')
def export_inventory():
    filter_type = request.args.get('filter', 'active')
    inventory = load_inventory()

    if filter_type == 'archived':
        devices = [d for d in inventory if d.get('archived')]
        filename = 'inventory_archived.xlsx'
    elif filter_type == 'all':
        devices = inventory
        filename = 'inventory_all.xlsx'
    else:
        devices = [d for d in inventory if not d.get('archived')]
        filename = 'inventory_active.xlsx'

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = filter_type.capitalize()

    headers = [
        'ComputerName/ServiceTag', 'Location', 'Device Type', 'Model',
        'Username', 'Password', 'Island', 'Location Address w/zip',
        'Phone', 'Contact', 'Warranty End Date', 'Memory', 'CPU',
        'Operating Sys', 'NOTES:', 'Archived'
    ]
    ws.append(headers)

    # Bold header row
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for d in devices:
        ws.append([
            d.get('serial_number', ''),
            d.get('location_code', ''),
            d.get('device_type', ''),
            d.get('model', ''),
            d.get('assigned_user', ''),
            d.get('password', ''),
            d.get('island', ''),
            '',                               # Location Address w/zip — not stored
            d.get('phone', ''),
            '',                               # Contact — not stored separately
            d.get('purchase_date', ''),
            '',                               # Memory — not stored
            '',                               # CPU — not stored
            '',                               # Operating Sys — not stored
            d.get('notes', ''),
            'Yes' if d.get('archived') else 'No',
        ])

    # Auto-fit column widths
    for col in ws.columns:
        max_len = max((len(str(cell.value)) if cell.value else 0) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return Response(
        output.getvalue(),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename={filename}'}
    )


if __name__ == '__main__':
    init_data_storage()
    app.run(debug=True, host='0.0.0.0', port=5002)

