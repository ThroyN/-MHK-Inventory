"""
Daily backup of inventory.db to OneDrive.
Run via Windows Task Scheduler — see README or ask Claude for setup steps.
"""
import sqlite3
import os
from datetime import datetime, timedelta

DB_FILE     = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'inventory.db')
BACKUP_DIR  = r'C:\Users\MGR\Mhk.Inventory.Backups\OneDrive - Mental Health Kokua\MHK_Shared\Mhk.Inventory.DataBase'
KEEP_DAYS   = 7

def run_backup():
    if not os.path.exists(DB_FILE):
        raise RuntimeError(f'Database not found at {DB_FILE}')

    os.makedirs(BACKUP_DIR, exist_ok=True)

    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    dest = os.path.join(BACKUP_DIR, f'inventory_{timestamp}.db')

    # SQLite online backup — safe even while app is running
    src  = sqlite3.connect(DB_FILE)
    dst  = sqlite3.connect(dest)
    src.backup(dst)
    dst.close()
    src.close()

    print(f'Backup saved: {dest}')

    # Remove backups older than KEEP_DAYS
    cutoff = datetime.now() - timedelta(days=KEEP_DAYS)
    for fname in os.listdir(BACKUP_DIR):
        if not fname.startswith('inventory_') or not fname.endswith('.db'):
            continue
        fpath = os.path.join(BACKUP_DIR, fname)
        if datetime.fromtimestamp(os.path.getmtime(fpath)) < cutoff:
            os.remove(fpath)
            print(f'Removed old backup: {fname}')

if __name__ == '__main__':
    run_backup()
