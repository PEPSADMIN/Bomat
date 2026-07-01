"""
PEPS BOM Automation Tool - Database Layer
Version: 2.1
Date: 22 May 2026

SQLite database schema and all database functions.
Tables: products, runs, run_items, replace_snapshots, settings (NEW)

CHANGELOG v2.1:
- Added settings table for RAMCO connection configuration
- Added password encryption support for settings
- Settings stored as single-row table (max 1 record)
"""

import sqlite3
import json
import base64
from datetime import datetime
from typing import List, Dict, Optional, Tuple
from cryptography.fernet import Fernet
import os

# Import config
try:
    import config_v2_1 as config
except ImportError:
    import config_v2_0 as config


class PasswordEncryption:
    """Simple password encryption/decryption using Fernet"""
    
    def __init__(self):
        # Generate or load encryption key
        key_file = os.path.join(os.path.dirname(__file__), '.encryption_key')
        
        if os.path.exists(key_file):
            with open(key_file, 'rb') as f:
                self.key = f.read()
        else:
            self.key = Fernet.generate_key()
            with open(key_file, 'wb') as f:
                f.write(self.key)
        
        self.cipher = Fernet(self.key)
    
    def encrypt(self, password: str) -> str:
        """Encrypt password and return base64 string"""
        if not password:
            return ''
        encrypted_bytes = self.cipher.encrypt(password.encode())
        return base64.b64encode(encrypted_bytes).decode()
    
    def decrypt(self, encrypted_password: str) -> str:
        """Decrypt password from base64 string"""
        if not encrypted_password:
            return ''
        try:
            encrypted_bytes = base64.b64decode(encrypted_password.encode())
            decrypted_bytes = self.cipher.decrypt(encrypted_bytes)
            return decrypted_bytes.decode()
        except Exception:
            return ''


# Global encryption instance
password_encryptor = PasswordEncryption()


class Database:
    """Database manager for BOM Tool"""
    
    def __init__(self, db_path: str = None):
        """Initialize database connection"""
        self.db_path = db_path or config.DATABASE_PATH
        self.init_db()
    
    def get_connection(self):
        """Get database connection with row factory"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn
    
    def init_db(self):
        """Initialize database schema - creates all tables"""
        conn = self.get_connection()
        cursor = conn.cursor()

        # WAL mode: allows concurrent reads while writing, eliminates most lock errors
        cursor.execute('PRAGMA journal_mode=WAL')
        cursor.execute('PRAGMA synchronous=NORMAL')
        
        # ====================================================================
        # PRODUCTS TABLE - Metadata for all scanned xlsm files
        # ====================================================================
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS products (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                name         TEXT    NOT NULL,
                family       TEXT    NOT NULL,
                filepath     TEXT    NOT NULL UNIQUE,
                category     TEXT,
                heights      TEXT,
                lengths      TEXT,
                widths       TEXT,
                colours      TEXT,
                departments  TEXT    DEFAULT '[]',
                sections     TEXT    DEFAULT '[]',
                wh_codes     TEXT    DEFAULT '[]',
                sku_count    INTEGER DEFAULT 0,
                component_count INTEGER DEFAULT 0,
                status       TEXT    DEFAULT 'ok',
                last_scanned TEXT,
                created_at   TEXT    NOT NULL
            )
        ''')
        
        # ====================================================================
        # RUNS TABLE - History of all BOM generation runs
        # ====================================================================
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS runs (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                run_label       TEXT    NOT NULL,
                status          TEXT    DEFAULT 'running',
                product_count   INTEGER DEFAULT 0,
                rows_generated  INTEGER DEFAULT 0,
                run_by          TEXT    DEFAULT 'Admin',
                created_at      TEXT    NOT NULL,
                completed_at    TEXT,
                mode            TEXT    DEFAULT 'FULL',
                output_mode     TEXT    DEFAULT 'ZIP',
                output_filename TEXT,
                approval_ref    TEXT,
                warnings        INTEGER DEFAULT 0,
                error_message   TEXT
            )
        ''')
        
        # ====================================================================
        # RUN_ITEMS TABLE - Details of products included in each run
        # ====================================================================
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS run_items (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id      INTEGER NOT NULL,
                product_id  INTEGER NOT NULL,
                rows_count  INTEGER DEFAULT 0,
                status      TEXT    DEFAULT 'pending',
                FOREIGN KEY (run_id) REFERENCES runs(id),
                FOREIGN KEY (product_id) REFERENCES products(id)
            )
        ''')
        
        # ====================================================================
        # REPLACE_SNAPSHOTS TABLE - Global material replacement history
        # ====================================================================
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS replace_snapshots (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at    TEXT    NOT NULL,
                old_code      TEXT    NOT NULL,
                new_code      TEXT    NOT NULL,
                reason        TEXT,
                approval_ref  TEXT,
                run_by        TEXT,
                files_changed INTEGER DEFAULT 0,
                rows_changed  INTEGER DEFAULT 0,
                affected_data TEXT,
                status        TEXT    DEFAULT 'active'
            )
        ''')
        
        # ====================================================================
        # SETTINGS TABLE (NEW v2.1) - RAMCO Connection Configuration
        # Single-row table storing global settings for Excel generation
        # ====================================================================
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS settings (
                id                     INTEGER PRIMARY KEY CHECK (id = 1),
                url                    TEXT    NOT NULL DEFAULT 'https://peps.ramcoes.com/RVW/extui/vwrt/LaunchPanel.htm',
                user_id                TEXT    NOT NULL DEFAULT 'PEPS54',
                password               TEXT    NOT NULL DEFAULT '',
                role_name              TEXT    NOT NULL DEFAULT 'ERP-ADMIN',
                organisation_unit      TEXT    NOT NULL DEFAULT '14- HOSUR BARE COIR - FACTORY UNIT',
                language_id            TEXT    DEFAULT '1',
                number_of_records      TEXT    DEFAULT '285',
                bpc                    TEXT    DEFAULT 'Y',
                bpc_name               TEXT    DEFAULT 'Discrete Production',
                enable                 TEXT    DEFAULT '1',
                version                TEXT    DEFAULT 'DotNet',
                user_interface_type    TEXT    DEFAULT 'EXTUI',
                excel_start_row        TEXT    DEFAULT '5',
                excel_start_col        TEXT    DEFAULT '2',
                skip_blank_values      TEXT    DEFAULT '1',
                mdcf_version           TEXT    DEFAULT 'Service',
                created_at             TEXT    NOT NULL,
                updated_at             TEXT    NOT NULL
            )
        ''')
        
        # Initialize default settings if table is empty
        cursor.execute('SELECT COUNT(*) as count FROM settings')
        if cursor.fetchone()['count'] == 0:
            cursor.execute('''
                INSERT INTO settings (
                    id, url, user_id, password, role_name, organisation_unit,
                    language_id, number_of_records, bpc, bpc_name, enable,
                    version, user_interface_type, excel_start_row, excel_start_col,
                    skip_blank_values, mdcf_version, created_at, updated_at
                ) VALUES (
                    1,
                    'https://peps.ramcoes.com/RVW/extui/vwrt/LaunchPanel.htm',
                    'PEPS54',
                    '',
                    'ERP-ADMIN',
                    '14- HOSUR BARE COIR - FACTORY UNIT',
                    '1', '285', 'Y', 'Discrete Production', '1',
                    'DotNet', 'EXTUI', '5', '2', '1', 'Service',
                    ?, ?
                )
            ''', (datetime.now().isoformat(), datetime.now().isoformat()))
            print("✓ Default settings initialized")
        
        # ====================================================================
        # ROLES TABLE - Custom hierarchy roles with tab mappings
        # ====================================================================
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS roles (
                id                 INTEGER PRIMARY KEY AUTOINCREMENT,
                name               TEXT    NOT NULL UNIQUE,
                display_name       TEXT    NOT NULL,
                description        TEXT    DEFAULT '',
                allowed_tabs       TEXT    NOT NULL DEFAULT '[]',
                can_approve        INTEGER NOT NULL DEFAULT 0,
                can_create_users   INTEGER NOT NULL DEFAULT 0,
                can_manage_users   INTEGER NOT NULL DEFAULT 0,
                can_manage_roles   INTEGER NOT NULL DEFAULT 0,
                is_system          INTEGER NOT NULL DEFAULT 0,
                created_at         TEXT    NOT NULL
            )
        ''')

        # ====================================================================
        # USERS TABLE - Login accounts and tab permissions
        # ====================================================================
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                username      TEXT    NOT NULL UNIQUE COLLATE NOCASE,
                password_hash TEXT    NOT NULL,
                role          TEXT    NOT NULL DEFAULT 'user',
                allowed_tabs  TEXT    NOT NULL DEFAULT '[]',
                is_active     INTEGER NOT NULL DEFAULT 1,
                created_at    TEXT    NOT NULL,
                last_login    TEXT
            )
        ''')

        # Migrate existing DBs: add columns if absent
        existing_cols = {r[1] for r in cursor.execute("PRAGMA table_info(products)").fetchall()}
        for col, default in [('departments', "DEFAULT '[]'"), ('sections', "DEFAULT '[]'"),
                             ('wh_codes', "DEFAULT '[]'"), ('item_codes', "DEFAULT '[]'"),
                             ('product_group', "DEFAULT ''")]:
            if col not in existing_cols:
                cursor.execute(f"ALTER TABLE products ADD COLUMN {col} TEXT {default}")

        # ====================================================================
        # APPROVALS TABLE — change requests pending admin/developer review
        # ====================================================================
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS approvals (
                id                 INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id            INTEGER NOT NULL,
                username           TEXT    NOT NULL,
                action_type        TEXT    NOT NULL,
                action_summary     TEXT    NOT NULL,
                action_data        TEXT    NOT NULL DEFAULT '{}',
                status             TEXT    NOT NULL DEFAULT 'pending',
                reviewer_id        INTEGER,
                reviewer_name      TEXT,
                rejection_category TEXT,
                rejection_reason   TEXT,
                created_at         TEXT    NOT NULL,
                reviewed_at        TEXT
            )
        ''')

        # Migrate existing DBs: two-stage Sub-Admin -> Admin audit chain.
        # When a sub_admin approves, the action executes immediately (same
        # as before) but the request also gets flagged for Admin/Developer
        # audit. Admin can acknowledge it or overturn it (which rolls back
        # the action and notifies both the sub_admin and original user).
        existing_appr_cols = {r[1] for r in cursor.execute("PRAGMA table_info(approvals)").fetchall()}
        for col, coldef in [
            ('reviewer_role',       "TEXT"),
            ('result_ref',          "TEXT"),
            ('admin_audit_status',  "TEXT"),
            ('admin_reviewer_id',   "INTEGER"),
            ('admin_reviewer_name', "TEXT"),
            ('admin_remarks',       "TEXT"),
            ('admin_reviewed_at',   "TEXT"),
        ]:
            if col not in existing_appr_cols:
                cursor.execute(f"ALTER TABLE approvals ADD COLUMN {col} {coldef}")

        # ====================================================================
        # NOTIFICATIONS TABLE — per-user inbox for approval outcomes
        # ====================================================================
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS notifications (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id      INTEGER NOT NULL,
                message      TEXT    NOT NULL,
                notif_type   TEXT    NOT NULL DEFAULT 'info',
                is_read      INTEGER NOT NULL DEFAULT 0,
                approval_id  INTEGER,
                created_at   TEXT    NOT NULL
            )
        ''')

        # Create indexes for performance
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_products_family ON products(family)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_runs_created ON runs(created_at)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_snapshots_created ON replace_snapshots(created_at)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_approvals_status ON approvals(status)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_notifs_user ON notifications(user_id, is_read)')

        conn.commit()
        conn.close()
        print(f"✓ Database initialized: {self.db_path}")
    
    # ========================================================================
    # PRODUCTS FUNCTIONS (UNCHANGED FROM v2.0)
    # ========================================================================
    
    def add_product(self, name: str, family: str, filepath: str, **kwargs) -> int:
        """Register a new product in the database"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            INSERT INTO products (
                name, family, filepath, category, heights, lengths, widths,
                colours, sku_count, component_count, status, last_scanned, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            name,
            family,
            filepath,
            kwargs.get('category', ''),
            json.dumps(kwargs.get('heights', [])),
            json.dumps(kwargs.get('lengths', [])),
            json.dumps(kwargs.get('widths', [])),
            json.dumps(kwargs.get('colours', [])),
            kwargs.get('sku_count', 0),
            kwargs.get('component_count', 0),
            kwargs.get('status', 'ok'),
            datetime.now().isoformat(),
            datetime.now().isoformat()
        ))
        
        product_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return product_id
    
    def update_product(self, product_id: int, **kwargs):
        """Update product metadata"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        updates = []
        values = []
        
        for key, value in kwargs.items():
            if key in ['heights', 'lengths', 'widths', 'colours']:
                value = json.dumps(value)
            updates.append(f"{key} = ?")
            values.append(value)
        
        values.append(datetime.now().isoformat())
        values.append(product_id)
        
        query = f"UPDATE products SET {', '.join(updates)}, last_scanned = ? WHERE id = ?"
        cursor.execute(query, values)
        
        conn.commit()
        conn.close()
    
    def get_product_by_id(self, product_id: int) -> Optional[Dict]:
        """Get product by ID"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT * FROM products WHERE id = ?', (product_id,))
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return dict(row)
        return None
    
    def get_product_by_filepath(self, filepath: str) -> Optional[Dict]:
        """Get product by filepath"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT * FROM products WHERE filepath = ?', (filepath,))
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return dict(row)
        return None
    
    def get_all_products(self) -> List[Dict]:
        """Get all products"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT * FROM products ORDER BY family, name')
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]
    
    def get_products_by_family(self, family: str) -> List[Dict]:
        """Get all products in a family"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT * FROM products WHERE family = ? ORDER BY name', (family,))
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]
    
    def delete_all_products(self):
        """Delete all products (for rescan)"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('DELETE FROM products')
        conn.commit()
        conn.close()
    
    # ========================================================================
    # RUNS FUNCTIONS (UNCHANGED FROM v2.0)
    # ========================================================================
    
    def create_run(self, product_ids: list, mode: str = 'FULL',
                   output_mode: str = 'ZIP', run_by: str = 'Admin',
                   approval_ref: str = '') -> int:
        """Create a new BOM generation run"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Generate run label
        if len(product_ids) == 1:
            product = self.get_product_by_id(product_ids[0])
            run_label = f"Run #{datetime.now().strftime('%Y%m%d%H%M')} - {product['name'] if product else 'Unknown'}"
        else:
            run_label = f"Run #{datetime.now().strftime('%Y%m%d%H%M')} - {len(product_ids)} products"
        
        cursor.execute('''
            INSERT INTO runs (
                run_label, status, product_count, run_by, created_at,
                mode, output_mode, approval_ref
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            run_label,
            'running',
            len(product_ids),
            run_by,
            datetime.now().isoformat(),
            mode,
            output_mode,
            approval_ref
        ))
        
        run_id = cursor.lastrowid
        
        # Add run items
        for product_id in product_ids:
            cursor.execute('''
                INSERT INTO run_items (run_id, product_id, status)
                VALUES (?, ?, ?)
            ''', (run_id, product_id, 'pending'))
        
        conn.commit()
        conn.close()
        return run_id
    
    def complete_run(self, run_id: int, output_filename: str = None, 
                     product_count: int = None, rows_generated: int = None,
                     warnings: int = 0):
        """Mark run as completed"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        updates = {
            'status': 'OK',
            'completed_at': datetime.now().isoformat(),
            'warnings': warnings
        }
        
        if output_filename:
            updates['output_filename'] = output_filename
        if product_count is not None:
            updates['product_count'] = product_count
        if rows_generated is not None:
            updates['rows_generated'] = rows_generated
        
        set_clause = ', '.join([f"{k} = ?" for k in updates.keys()])
        values = list(updates.values()) + [run_id]
        
        cursor.execute(f'UPDATE runs SET {set_clause} WHERE id = ?', values)
        
        conn.commit()
        conn.close()
    
    def fail_run(self, run_id: int, error_message: str):
        """Mark run as failed"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            UPDATE runs SET status = ?, completed_at = ?, error_message = ?
            WHERE id = ?
        ''', ('ERROR', datetime.now().isoformat(), error_message, run_id))
        
        conn.commit()
        conn.close()
    
    def get_run_history(self, limit: int = 100, from_date: str = '', 
                        to_date: str = '') -> List[Dict]:
        """Get run history with optional date filtering"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        query = 'SELECT * FROM runs'
        params = []
        
        # Add date filtering if provided
        if from_date and to_date:
            query += ' WHERE date(created_at) BETWEEN ? AND ?'
            params.extend([from_date, to_date])
        elif from_date:
            query += ' WHERE date(created_at) >= ?'
            params.append(from_date)
        elif to_date:
            query += ' WHERE date(created_at) <= ?'
            params.append(to_date)
        
        query += ' ORDER BY created_at DESC LIMIT ?'
        params.append(limit)
        
        cursor.execute(query, params)
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]
    
    def get_run_by_id(self, run_id: int) -> Optional[Dict]:
        """Get run by ID"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT * FROM runs WHERE id = ?', (run_id,))
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return dict(row)
        return None
    
    # ========================================================================
    # REPLACE SNAPSHOTS FUNCTIONS (UNCHANGED FROM v2.0)
    # ========================================================================
    
    def create_replace_snapshot(self, old_code: str, new_code: str,
                                reason: str, approval_ref: str, run_by: str,
                                files_changed: int, rows_changed: int,
                                affected_data: str) -> int:
        """Create a global replace snapshot"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            INSERT INTO replace_snapshots (
                created_at, old_code, new_code, reason, approval_ref,
                run_by, files_changed, rows_changed, affected_data, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            datetime.now().isoformat(),
            old_code,
            new_code,
            reason,
            approval_ref,
            run_by,
            files_changed,
            rows_changed,
            affected_data,
            'active'
        ))
        
        snapshot_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return snapshot_id
    
    def get_replace_snapshots(self, limit: int = 50, active_only: bool = True) -> List[Dict]:
        """Get replace snapshots — by default only returns active (rollback-able) ones."""
        conn = self.get_connection()
        cursor = conn.cursor()
        if active_only:
            cursor.execute('''
                SELECT * FROM replace_snapshots
                WHERE status = 'active'
                ORDER BY created_at DESC LIMIT ?
            ''', (limit,))
        else:
            cursor.execute('''
                SELECT * FROM replace_snapshots
                ORDER BY created_at DESC LIMIT ?
            ''', (limit,))
        
        rows = cursor.fetchall()
        conn.close()
        
        return [dict(row) for row in rows]
    
    def get_replace_snapshot(self, snapshot_id: int) -> Optional[Dict]:
        """Get snapshot by ID"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT * FROM replace_snapshots WHERE id = ?', (snapshot_id,))
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return dict(row)
        return None
    
    def mark_snapshot_rolled_back(self, snapshot_id: int):
        """Mark snapshot as rolled back"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('''
            UPDATE replace_snapshots SET status = ? WHERE id = ?
        ''', ('rolled_back', snapshot_id))
        
        conn.commit()
        conn.close()
    
    # ========================================================================
    # SETTINGS FUNCTIONS (NEW v2.1)
    # ========================================================================
    
    def get_settings(self) -> Dict:
        """Get current settings (decrypts password)"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT * FROM settings WHERE id = 1')
        row = cursor.fetchone()
        conn.close()
        
        if row:
            settings = dict(row)
            # Decrypt password
            settings['password'] = password_encryptor.decrypt(settings['password'])
            return settings
        
        # Return defaults if no settings found
        return {
            'url': 'https://peps.ramcoes.com/RVW/extui/vwrt/LaunchPanel.htm',
            'user_id': 'PEPS54',
            'password': '',
            'role_name': 'ERP-ADMIN',
            'organisation_unit': '14- HOSUR BARE COIR - FACTORY UNIT',
            'language_id': '1',
            'number_of_records': '285',
            'bpc': 'Y',
            'bpc_name': 'Discrete Production',
            'enable': '1',
            'version': 'DotNet',
            'user_interface_type': 'EXTUI',
            'excel_start_row': '5',
            'excel_start_col': '2',
            'skip_blank_values': '1',
            'mdcf_version': 'Service'
        }
    
    def save_settings(self, settings: Dict) -> bool:
        """Save settings (encrypts password)"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Encrypt password before saving
        encrypted_password = password_encryptor.encrypt(settings.get('password', ''))
        
        try:
            cursor.execute('''
                INSERT OR REPLACE INTO settings (
                    id, url, user_id, password, role_name, organisation_unit,
                    language_id, number_of_records, bpc, bpc_name, enable,
                    version, user_interface_type, excel_start_row, excel_start_col,
                    skip_blank_values, mdcf_version, created_at, updated_at
                ) VALUES (
                    1, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 
                    COALESCE((SELECT created_at FROM settings WHERE id = 1), ?), 
                    ?
                )
            ''', (
                settings.get('url', ''),
                settings.get('user_id', ''),
                encrypted_password,
                settings.get('role_name', ''),
                settings.get('organisation_unit', ''),
                settings.get('language_id', '1'),
                settings.get('number_of_records', '285'),
                settings.get('bpc', 'Y'),
                settings.get('bpc_name', 'Discrete Production'),
                settings.get('enable', '1'),
                settings.get('version', 'DotNet'),
                settings.get('user_interface_type', 'EXTUI'),
                settings.get('excel_start_row', '5'),
                settings.get('excel_start_col', '2'),
                settings.get('skip_blank_values', '1'),
                settings.get('mdcf_version', 'Service'),
                datetime.now().isoformat(),
                datetime.now().isoformat()
            ))
            
            conn.commit()
            conn.close()
            return True
        
        except Exception as e:
            import traceback
            print(f"Error saving settings: {e}")
            traceback.print_exc()
            try: conn.rollback()
            except: pass
            conn.close()
            return False
    
    def update_setting(self, key: str, value) -> None:
        """Update a single setting column without touching the rest of the row."""
        _allowed = {
            'number_of_records', 'url', 'user_id', 'role_name', 'organisation_unit',
            'language_id', 'bpc', 'bpc_name', 'enable', 'version',
            'user_interface_type', 'excel_start_row', 'excel_start_col',
            'skip_blank_values', 'mdcf_version',
        }
        if key not in _allowed:
            raise ValueError(f'Unknown setting key: {key}')
        conn = self.get_connection()
        conn.execute(f'UPDATE settings SET {key} = ?, updated_at = ? WHERE id = 1',
                     (str(value), datetime.now().isoformat()))
        conn.commit()
        conn.close()

    # ========================================================================
    # USER / AUTH FUNCTIONS
    # ========================================================================

    def create_user(self, username: str, password_hash: str,
                    role: str = 'user', allowed_tabs: list = None) -> int:
        """Insert a new user. Returns new user id."""
        tabs = json.dumps(allowed_tabs or [])
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(
            'INSERT INTO users (username, password_hash, role, allowed_tabs, created_at) VALUES (?,?,?,?,?)',
            (username, password_hash, role, tabs, datetime.now().isoformat())
        )
        new_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return new_id

    def get_user_by_username(self, username: str) -> Optional[Dict]:
        conn = self.get_connection()
        row = conn.execute(
            'SELECT * FROM users WHERE username = ? COLLATE NOCASE', (username,)
        ).fetchone()
        conn.close()
        if not row:
            return None
        u = dict(row)
        u['allowed_tabs'] = json.loads(u['allowed_tabs'] or '[]')
        return u

    def get_user_by_id(self, user_id: int) -> Optional[Dict]:
        conn = self.get_connection()
        row = conn.execute('SELECT * FROM users WHERE id = ?', (user_id,)).fetchone()
        conn.close()
        if not row:
            return None
        u = dict(row)
        u['allowed_tabs'] = json.loads(u['allowed_tabs'] or '[]')
        return u

    def get_all_users(self) -> List[Dict]:
        conn = self.get_connection()
        rows = conn.execute('SELECT * FROM users ORDER BY created_at').fetchall()
        conn.close()
        result = []
        for row in rows:
            u = dict(row)
            u['allowed_tabs'] = json.loads(u['allowed_tabs'] or '[]')
            result.append(u)
        return result

    def update_user_password(self, user_id: int, password_hash: str) -> None:
        conn = self.get_connection()
        conn.execute('UPDATE users SET password_hash = ? WHERE id = ?', (password_hash, user_id))
        conn.commit()
        conn.close()

    def update_user_tabs(self, user_id: int, allowed_tabs: list) -> None:
        conn = self.get_connection()
        conn.execute('UPDATE users SET allowed_tabs = ? WHERE id = ?',
                     (json.dumps(allowed_tabs), user_id))
        conn.commit()
        conn.close()

    def update_user_role(self, user_id: int, role: str) -> None:
        conn = self.get_connection()
        conn.execute('UPDATE users SET role = ? WHERE id = ?', (role, user_id))
        conn.commit()
        conn.close()

    def toggle_user_active(self, user_id: int) -> bool:
        """Flip is_active. Returns new state."""
        conn = self.get_connection()
        current = conn.execute('SELECT is_active FROM users WHERE id = ?', (user_id,)).fetchone()
        new_state = 0 if current and current['is_active'] else 1
        conn.execute('UPDATE users SET is_active = ? WHERE id = ?', (new_state, user_id))
        conn.commit()
        conn.close()
        return bool(new_state)

    def update_last_login(self, user_id: int) -> None:
        conn = self.get_connection()
        conn.execute('UPDATE users SET last_login = ? WHERE id = ?',
                     (datetime.now().isoformat(), user_id))
        conn.commit()
        conn.close()

    # ========================================================================
    # ROLES FUNCTIONS
    # ========================================================================

    def seed_system_roles(self, all_tabs: list) -> None:
        """Insert built-in roles if they don't exist yet."""
        conn = self.get_connection()
        cursor = conn.cursor()
        system_roles = [
            ('developer', 'Developer', 'Full system access + developer tools',
             all_tabs, 1, 1, 1, 1),
            ('admin', 'Admin', 'Full access including user management and approvals',
             all_tabs, 1, 1, 1, 0),
            ('sub_admin', 'Sub Admin',
             'Can approve/reject requests but cannot create users or access CMS/Ramco MDCF',
             [t for t in all_tabs if t not in ('cms', 'settings')], 1, 0, 0, 0),
            ('user', 'User', 'Standard user — tabs assigned individually',
             [], 0, 0, 0, 0),
        ]
        for name, disp, desc, tabs, appr, create, manage, manage_roles in system_roles:
            existing = cursor.execute('SELECT id FROM roles WHERE name=?', (name,)).fetchone()
            if not existing:
                cursor.execute('''
                    INSERT INTO roles (name, display_name, description, allowed_tabs,
                        can_approve, can_create_users, can_manage_users, can_manage_roles,
                        is_system, created_at)
                    VALUES (?,?,?,?,?,?,?,?,1,?)
                ''', (name, disp, desc, json.dumps(tabs),
                      appr, create, manage, manage_roles,
                      datetime.now().isoformat()))
            else:
                # Keep system role tabs in sync with ALL_TABS additions
                if name in ('developer', 'admin'):
                    cursor.execute('UPDATE roles SET allowed_tabs=? WHERE name=?',
                                   (json.dumps(tabs), name))
        conn.commit()
        conn.close()

    def get_all_roles(self) -> List[Dict]:
        conn = self.get_connection()
        rows = conn.execute('SELECT * FROM roles ORDER BY is_system DESC, id').fetchall()
        conn.close()
        result = []
        for r in rows:
            d = dict(r)
            d['allowed_tabs'] = json.loads(d['allowed_tabs'] or '[]')
            result.append(d)
        return result

    def get_role_by_name(self, name: str) -> Optional[Dict]:
        conn = self.get_connection()
        row = conn.execute('SELECT * FROM roles WHERE name=?', (name,)).fetchone()
        conn.close()
        if not row: return None
        d = dict(row)
        d['allowed_tabs'] = json.loads(d['allowed_tabs'] or '[]')
        return d

    def get_role_by_id(self, role_id: int) -> Optional[Dict]:
        conn = self.get_connection()
        row = conn.execute('SELECT * FROM roles WHERE id=?', (role_id,)).fetchone()
        conn.close()
        if not row: return None
        d = dict(row)
        d['allowed_tabs'] = json.loads(d['allowed_tabs'] or '[]')
        return d

    def create_role(self, name: str, display_name: str, description: str,
                    allowed_tabs: list, can_approve: int, can_create_users: int,
                    can_manage_users: int, can_manage_roles: int) -> int:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO roles (name, display_name, description, allowed_tabs,
                can_approve, can_create_users, can_manage_users, can_manage_roles,
                is_system, created_at)
            VALUES (?,?,?,?,?,?,?,?,0,?)
        ''', (name, display_name, description, json.dumps(allowed_tabs),
              can_approve, can_create_users, can_manage_users, can_manage_roles,
              datetime.now().isoformat()))
        new_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return new_id

    def update_role(self, role_id: int, display_name: str, description: str,
                    allowed_tabs: list, can_approve: int, can_create_users: int,
                    can_manage_users: int, can_manage_roles: int) -> None:
        conn = self.get_connection()
        conn.execute('''
            UPDATE roles SET display_name=?, description=?, allowed_tabs=?,
                can_approve=?, can_create_users=?, can_manage_users=?, can_manage_roles=?
            WHERE id=? AND is_system=0
        ''', (display_name, description, json.dumps(allowed_tabs),
              can_approve, can_create_users, can_manage_users, can_manage_roles, role_id))
        conn.commit()
        conn.close()

    def delete_role(self, role_id: int) -> bool:
        conn = self.get_connection()
        row = conn.execute('SELECT is_system FROM roles WHERE id=?', (role_id,)).fetchone()
        if not row or row['is_system']:
            conn.close()
            return False
        conn.execute('DELETE FROM roles WHERE id=?', (role_id,))
        conn.commit()
        conn.close()
        return True

    # ========================================================================
    # APPROVALS FUNCTIONS
    # ========================================================================

    def create_approval(self, user_id: int, username: str, action_type: str,
                        action_summary: str, action_data: dict) -> int:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO approvals (user_id, username, action_type, action_summary, action_data, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (user_id, username, action_type, action_summary,
              json.dumps(action_data), datetime.now().isoformat()))
        new_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return new_id

    def get_pending_approvals(self) -> List[Dict]:
        conn = self.get_connection()
        rows = conn.execute(
            "SELECT * FROM approvals WHERE status='pending' ORDER BY created_at DESC"
        ).fetchall()
        conn.close()
        result = []
        for r in rows:
            d = dict(r)
            d['action_data'] = json.loads(d['action_data'] or '{}')
            result.append(d)
        return result

    def get_approvals_by_user(self, user_id: int) -> List[Dict]:
        conn = self.get_connection()
        rows = conn.execute(
            'SELECT * FROM approvals WHERE user_id=? ORDER BY created_at DESC', (user_id,)
        ).fetchall()
        conn.close()
        result = []
        for r in rows:
            d = dict(r)
            d['action_data'] = json.loads(d['action_data'] or '{}')
            result.append(d)
        return result

    def get_approval_by_id(self, approval_id: int) -> Optional[Dict]:
        conn = self.get_connection()
        row = conn.execute('SELECT * FROM approvals WHERE id=?', (approval_id,)).fetchone()
        conn.close()
        if not row:
            return None
        d = dict(row)
        d['action_data'] = json.loads(d['action_data'] or '{}')
        return d

    def resolve_approval(self, approval_id: int, status: str, reviewer_id: int,
                         reviewer_name: str, rejection_category: str = '',
                         rejection_reason: str = '', reviewer_role: str = '') -> None:
        conn = self.get_connection()
        conn.execute('''
            UPDATE approvals
            SET status=?, reviewer_id=?, reviewer_name=?,
                rejection_category=?, rejection_reason=?, reviewed_at=?,
                reviewer_role=?
            WHERE id=?
        ''', (status, reviewer_id, reviewer_name,
              rejection_category, rejection_reason,
              datetime.now().isoformat(), reviewer_role, approval_id))
        conn.commit()
        conn.close()

    def flag_admin_audit(self, approval_id: int) -> None:
        """Mark an approval (just approved by a sub_admin) as needing Admin/Developer audit."""
        conn = self.get_connection()
        conn.execute("UPDATE approvals SET admin_audit_status='pending_audit' WHERE id=?", (approval_id,))
        conn.commit()
        conn.close()

    def set_approval_result_ref(self, approval_id: int, ref: str) -> None:
        """Store a pointer to whatever the approved action created (e.g. a run id
        or a new product's filename), so an Admin overturn can find and undo it."""
        conn = self.get_connection()
        conn.execute("UPDATE approvals SET result_ref=? WHERE id=?", (ref, approval_id))
        conn.commit()
        conn.close()

    def get_pending_admin_audits(self) -> List[Dict]:
        """Approvals a sub_admin approved that are still awaiting Admin/Developer audit."""
        conn = self.get_connection()
        rows = conn.execute(
            "SELECT * FROM approvals WHERE admin_audit_status='pending_audit' ORDER BY reviewed_at DESC"
        ).fetchall()
        conn.close()
        result = []
        for r in rows:
            d = dict(r)
            d['action_data'] = json.loads(d['action_data'] or '{}')
            result.append(d)
        return result

    def resolve_admin_audit(self, approval_id: int, audit_status: str, admin_id: int,
                            admin_name: str, remarks: str = '') -> None:
        """audit_status: 'acknowledged' (no change) or 'overturned' (caller handles rollback).
        On overturn, also flips the main status so history/UI shows it as overturned, not approved."""
        conn = self.get_connection()
        if audit_status == 'overturned':
            conn.execute('''
                UPDATE approvals
                SET admin_audit_status=?, admin_reviewer_id=?, admin_reviewer_name=?,
                    admin_remarks=?, admin_reviewed_at=?, status='overturned'
                WHERE id=?
            ''', (audit_status, admin_id, admin_name, remarks, datetime.now().isoformat(), approval_id))
        else:
            conn.execute('''
                UPDATE approvals
                SET admin_audit_status=?, admin_reviewer_id=?, admin_reviewer_name=?,
                    admin_remarks=?, admin_reviewed_at=?
                WHERE id=?
            ''', (audit_status, admin_id, admin_name, remarks, datetime.now().isoformat(), approval_id))
        conn.commit()
        conn.close()

    def get_all_approvals(self, limit: int = 200) -> List[Dict]:
        conn = self.get_connection()
        rows = conn.execute(
            'SELECT * FROM approvals ORDER BY created_at DESC LIMIT ?', (limit,)
        ).fetchall()
        conn.close()
        result = []
        for r in rows:
            d = dict(r)
            d['action_data'] = json.loads(d['action_data'] or '{}')
            result.append(d)
        return result

    # ========================================================================
    # NOTIFICATIONS FUNCTIONS
    # ========================================================================

    def create_notification(self, user_id: int, message: str,
                            notif_type: str = 'info', approval_id: int = None) -> int:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO notifications (user_id, message, notif_type, approval_id, created_at)
            VALUES (?, ?, ?, ?, ?)
        ''', (user_id, message, notif_type, approval_id, datetime.now().isoformat()))
        new_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return new_id

    def get_user_notifications(self, user_id: int, limit: int = 30) -> List[Dict]:
        conn = self.get_connection()
        rows = conn.execute('''
            SELECT * FROM notifications WHERE user_id=?
            ORDER BY created_at DESC LIMIT ?
        ''', (user_id, limit)).fetchall()
        conn.close()
        return [dict(r) for r in rows]

    def get_unread_count(self, user_id: int) -> int:
        conn = self.get_connection()
        row = conn.execute(
            'SELECT COUNT(*) as c FROM notifications WHERE user_id=? AND is_read=0', (user_id,)
        ).fetchone()
        conn.close()
        return row['c'] if row else 0

    def get_pending_count(self) -> int:
        conn = self.get_connection()
        row = conn.execute(
            "SELECT COUNT(*) as c FROM approvals WHERE status='pending'"
        ).fetchone()
        conn.close()
        return row['c'] if row else 0

    def mark_notifications_read(self, user_id: int) -> None:
        conn = self.get_connection()
        conn.execute('UPDATE notifications SET is_read=1 WHERE user_id=?', (user_id,))
        conn.commit()
        conn.close()

    # ========================================================================
    # UTILITY FUNCTIONS
    # ========================================================================

    def get_stats(self) -> Dict:
        """Get database statistics"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        stats = {}
        
        cursor.execute('SELECT COUNT(*) as count FROM products')
        stats['total_products'] = cursor.fetchone()['count']
        
        cursor.execute('SELECT COUNT(*) as count FROM runs')
        stats['total_runs'] = cursor.fetchone()['count']
        
        cursor.execute('SELECT COUNT(*) as count FROM replace_snapshots WHERE status = ?', ('active',))
        stats['active_replacements'] = cursor.fetchone()['count']
        
        cursor.execute('SELECT SUM(sku_count) as total FROM products')
        result = cursor.fetchone()
        stats['total_skus'] = result['total'] if result['total'] else 0
        
        conn.close()
        return stats


# Singleton database instance
db = Database()
