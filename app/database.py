import os
import sqlite3
from datetime import datetime


DB_DIR = "data"
DB_PATH = os.path.join(DB_DIR, "history.db")
LOG_DIR = "logs"


class HistoryDB:
    def __init__(self, db_path=DB_PATH):
        os.makedirs(DB_DIR, exist_ok=True)
        os.makedirs(LOG_DIR, exist_ok=True)
        self.db_path = db_path
        self._init_db()

    def _connect(self):
        return sqlite3.connect(self.db_path)

    def _init_db(self):
        conn = self._connect()
        cur = conn.cursor()

        cur.execute("""
            CREATE TABLE IF NOT EXISTS sends (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                nome TEXT,
                documento TEXT,
                telefone TEXT,
                mensagem TEXT,
                status TEXT NOT NULL,
                error TEXT
            )
        """)

        conn.commit()
        conn.close()

    def save_send(self, nome, documento, telefone, mensagem, status, error=""):
        conn = self._connect()
        cur = conn.cursor()

        cur.execute("""
            INSERT INTO sends (created_at, nome, documento, telefone, mensagem, status, error)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            nome,
            documento,
            telefone,
            mensagem,
            status,
            error
        ))

        conn.commit()
        conn.close()

        self._append_log_file(nome, documento, telefone, status, error)

    def _append_log_file(self, nome, documento, telefone, status, error=""):
        log_file = os.path.join(LOG_DIR, f"log_{datetime.now().strftime('%Y%m%d')}.log")
        line = (
            f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] "
            f"nome={nome} | documento={documento} | telefone={telefone} | "
            f"status={status} | erro={error}\n"
        )
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(line)

    def was_already_sent(self, telefone, documento):
        conn = self._connect()
        cur = conn.cursor()

        cur.execute("""
            SELECT COUNT(*)
            FROM sends
            WHERE telefone = ? AND documento = ? AND status = 'enviado'
        """, (telefone, documento))

        count = cur.fetchone()[0]
        conn.close()
        return count > 0

    def get_recent_history(self, limit=100):
        conn = self._connect()
        cur = conn.cursor()

        cur.execute("""
            SELECT created_at, nome, documento, telefone, status, error
            FROM sends
            ORDER BY id DESC
            LIMIT ?
        """, (limit,))

        rows = cur.fetchall()
        conn.close()
        return rows