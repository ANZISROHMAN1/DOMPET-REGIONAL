import sqlite3
import os

DB_NAME = "klaim.db"

def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS reimbursements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama TEXT NOT NULL,
            kegiatan TEXT NOT NULL,
            nominal REAL NOT NULL,
            bank TEXT NOT NULL,
            rekening TEXT NOT NULL,
            bukti_path TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'Pending',
            tanggal TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            bukti_tf_path TEXT DEFAULT NULL
        )
    ''')
    
    # Attempt to add column to existing databases
    try:
        conn.execute("ALTER TABLE reimbursements ADD COLUMN bukti_tf_path TEXT DEFAULT NULL")
    except sqlite3.OperationalError:
        pass # Column already exists
        
    conn.commit()
    conn.close()

if __name__ == '__main__':
    init_db()
    print("Database initialized successfully.")
