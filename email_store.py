"""Persistent storage for employee email addresses.

The storage is a small SQLite database that keeps track of previously
entered/known email addresses so they can be re-used across sessions
without repeatedly editing CSV files.  It also offers compatibility
helpers to import/export the legacy ``emails.csv`` file.
"""
from __future__ import annotations

import csv
import os
import sqlite3
import threading
from contextlib import contextmanager
from typing import Dict, Iterable, Optional


def _ensure_parent_dir(path: str) -> None:
    parent = os.path.dirname(os.path.abspath(path))
    if parent and not os.path.exists(parent):
        os.makedirs(parent, exist_ok=True)


def _clean(value: Optional[str]) -> Optional[str]:
    if value is None:
        return None
    value = str(value).strip()
    return value or None


def _clean_lower(value: Optional[str]) -> Optional[str]:
    val = _clean(value)
    return val.lower() if val else None


def _clean_email(value: Optional[str]) -> Optional[str]:
    val = _clean(value)
    return val.lower() if val else None


class EmailStore:
    """Tiny helper around SQLite to remember employee email addresses."""

    def __init__(self, db_path: str):
        self.db_path = os.path.abspath(db_path)
        _ensure_parent_dir(self.db_path)
        self._lock = threading.RLock()
        self._init_db()

    @contextmanager
    def _connect(self):
        with self._lock:
            conn = sqlite3.connect(self.db_path)
            try:
                conn.execute("PRAGMA journal_mode=WAL;")
            except Exception:
                pass
            try:
                yield conn
            finally:
                conn.close()

    def _init_db(self) -> None:
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS email_addresses (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    seq TEXT UNIQUE,
                    account_no TEXT UNIQUE,
                    name_key TEXT UNIQUE,
                    display_name TEXT,
                    email TEXT NOT NULL,
                    last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            conn.commit()

    # ------------------------------------------------------------------
    # Public helpers
    # ------------------------------------------------------------------
    def import_from_csv(self, csv_path: str) -> None:
        """Merge entries from a legacy emails.csv file into the store."""
        csv_path = os.path.abspath(csv_path)
        if not os.path.exists(csv_path):
            return
        with open(csv_path, "r", encoding="utf-8", newline="") as fh:
            reader = csv.DictReader(fh)
            for row in reader:
                email = _clean_email(row.get("email"))
                if not email:
                    continue
                seq = _clean(row.get("seq"))
                account_no = _clean(row.get("account_no"))
                name = _clean_lower(row.get("name"))
                display_name = _clean(row.get("name"))
                self._upsert(seq=seq, account_no=account_no, name_key=name, display_name=display_name, email=email)

    def export_to_csv(self, csv_path: str) -> None:
        """Write all known email addresses to a CSV file (legacy compatibility)."""
        csv_path = os.path.abspath(csv_path)
        _ensure_parent_dir(csv_path)
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                SELECT seq, account_no, COALESCE(display_name, name_key) AS name, email
                FROM email_addresses
                ORDER BY last_updated DESC
                """
            )
            rows = cur.fetchall()
        fieldnames = ["seq", "account_no", "name", "email"]
        with open(csv_path, "w", encoding="utf-8", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=fieldnames)
            writer.writeheader()
            for seq, account, name, email in rows:
                writer.writerow({
                    "seq": seq or "",
                    "account_no": account or "",
                    "name": name or "",
                    "email": email or "",
                })

    def lookup(self, *, seq: Optional[str] = None, account_no: Optional[str] = None, name: Optional[str] = None) -> Optional[str]:
        seq = _clean(seq)
        account_no = _clean(account_no)
        name_key = _clean_lower(name)
        with self._connect() as conn:
            cur = conn.cursor()
            if seq:
                cur.execute("SELECT email FROM email_addresses WHERE seq = ? ORDER BY last_updated DESC LIMIT 1", (seq,))
                row = cur.fetchone()
                if row:
                    return row[0]
            if account_no:
                cur.execute("SELECT email FROM email_addresses WHERE account_no = ? ORDER BY last_updated DESC LIMIT 1", (account_no,))
                row = cur.fetchone()
                if row:
                    return row[0]
            if name_key:
                cur.execute("SELECT email FROM email_addresses WHERE name_key = ? ORDER BY last_updated DESC LIMIT 1", (name_key,))
                row = cur.fetchone()
                if row:
                    return row[0]
        return None

    def apply_to_employees(self, employees: Iterable[Dict]) -> None:
        for emp in employees:
            if not isinstance(emp, dict):
                continue
            if _clean_email(emp.get("email")):
                continue
            email = self.lookup(
                seq=emp.get("seq"),
                account_no=emp.get("account_no"),
                name=emp.get("name"),
            )
            if email:
                emp["email"] = email

    def remember_from_employee(self, employee: Dict, email: Optional[str]) -> None:
        clean_email = _clean_email(email)
        if not clean_email:
            return
        seq = _clean(employee.get("seq"))
        account_no = _clean(employee.get("account_no"))
        name_key = _clean_lower(employee.get("name"))
        display_name = _clean(employee.get("name"))
        self._upsert(seq=seq, account_no=account_no, name_key=name_key, display_name=display_name, email=clean_email)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _upsert(self, *, seq: Optional[str], account_no: Optional[str], name_key: Optional[str], display_name: Optional[str], email: str) -> None:
        if not email:
            return
        with self._connect() as conn:
            cur = conn.cursor()
            record_id: Optional[int] = None
            for column, value in (("seq", seq), ("account_no", account_no), ("name_key", name_key)):
                if not value:
                    continue
                cur.execute(f"SELECT id FROM email_addresses WHERE {column} = ?", (value,))
                row = cur.fetchone()
                if row:
                    record_id = int(row[0])
                    break

            if record_id is None and (seq or account_no or name_key):
                cur.execute(
                    """
                    INSERT INTO email_addresses (seq, account_no, name_key, display_name, email)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (seq, account_no, name_key, display_name, email),
                )
                record_id = cur.lastrowid

            if record_id is None:
                cur.execute(
                    """
                    INSERT INTO email_addresses (display_name, email)
                    VALUES (?, ?)
                    """,
                    (display_name or email, email),
                )
                record_id = cur.lastrowid

            cur.execute(
                """
                UPDATE email_addresses
                SET seq = COALESCE(?, seq),
                    account_no = COALESCE(?, account_no),
                    name_key = COALESCE(?, name_key),
                    display_name = COALESCE(?, display_name),
                    email = ?,
                    last_updated = CURRENT_TIMESTAMP
                WHERE id = ?
                """,
                (seq, account_no, name_key, display_name, email, record_id),
            )
            conn.commit()

    def count(self) -> int:
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM email_addresses")
            row = cur.fetchone()
            return int(row[0] or 0) if row else 0

    def to_dict(self) -> Dict[str, str]:
        data: Dict[str, str] = {}
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("SELECT seq, email FROM email_addresses WHERE seq IS NOT NULL AND seq != ''")
            for seq, email in cur.fetchall():
                if seq and email:
                    data[str(seq)] = email
        return data


__all__ = ["EmailStore"]
