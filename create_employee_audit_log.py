import sqlite3
from datetime import datetime

conn = sqlite3.connect("hr_payroll.db")
cur = conn.cursor()

cur.execute("""
CREATE TABLE IF NOT EXISTS employee_audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id INTEGER NOT NULL,
    action TEXT NOT NULL,
    field_name TEXT,
    old_value TEXT,
    new_value TEXT,
    changed_by TEXT NOT NULL,
    changed_at TEXT NOT NULL,
    FOREIGN KEY (employee_id) REFERENCES employees(id)
);
""")

conn.commit()
conn.close()

print("employee_audit_log table created successfully.")
