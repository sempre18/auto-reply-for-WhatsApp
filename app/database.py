import os
import sqlite3
from datetime import datetime, timedelta
from typing import Any

DB_DIR = "data"
DB_PATH = os.path.join(DB_DIR, "history.db")
LOG_DIR = "logs"


class HistoryDB:
    """Persistência de histórico, sessões e relatórios."""

    def __init__(self, db_path: str = DB_PATH):
        os.makedirs(DB_DIR, exist_ok=True)
        os.makedirs(LOG_DIR, exist_ok=True)
        self.db_path = db_path
        self._init_db()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self) -> None:
        with self._connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS sends (
                    id                INTEGER PRIMARY KEY AUTOINCREMENT,
                    session_id        TEXT,
                    created_at        TEXT NOT NULL,
                    nome              TEXT,
                    documento         TEXT,
                    telefone          TEXT,
                    mensagem          TEXT,
                    template_id       TEXT,
                    status            TEXT NOT NULL,
                    error             TEXT,
                    delay_used        REAL DEFAULT 0.0,
                    typing_delay      REAL DEFAULT 0.0,
                    preparation_status TEXT,
                    placeholders_left TEXT,
                    context_json      TEXT,
                    simulation        INTEGER DEFAULT 0
                );

                CREATE INDEX IF NOT EXISTS idx_sends_created_at
                    ON sends(created_at);

                CREATE INDEX IF NOT EXISTS idx_sends_telefone
                    ON sends(telefone);

                CREATE INDEX IF NOT EXISTS idx_sends_session_id
                    ON sends(session_id);

                CREATE TABLE IF NOT EXISTS session_stats (
                    id              INTEGER PRIMARY KEY AUTOINCREMENT,
                    session_id      TEXT UNIQUE,
                    session_start   TEXT NOT NULL,
                    session_end     TEXT,
                    total_sent      INTEGER DEFAULT 0,
                    total_errors    INTEGER DEFAULT 0,
                    total_skipped   INTEGER DEFAULT 0,
                    total_simulated INTEGER DEFAULT 0,
                    avg_delay       REAL DEFAULT 0.0
                );
                """
            )

    def save_send(
        self,
        nome: str,
        documento: str,
        telefone: str,
        mensagem: str,
        status: str,
        error: str = "",
        template_id: str = "",
        delay_used: float = 0.0,
        session_id: str = "",
        typing_delay: float = 0.0,
        preparation_status: str = "",
        placeholders_left: str = "",
        context_json: str = "",
        simulation: bool = False,
    ) -> None:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO sends (
                    session_id, created_at, nome, documento, telefone, mensagem,
                    template_id, status, error, delay_used, typing_delay,
                    preparation_status, placeholders_left, context_json, simulation
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    session_id,
                    now,
                    nome,
                    documento,
                    telefone,
                    mensagem,
                    template_id,
                    status,
                    error,
                    delay_used,
                    typing_delay,
                    preparation_status,
                    placeholders_left,
                    context_json,
                    int(simulation),
                ),
            )
        self._append_log_file(
            timestamp=now,
            nome=nome,
            documento=documento,
            telefone=telefone,
            status=status,
            error=error,
            delay_used=delay_used,
            template_id=template_id,
            session_id=session_id,
        )

    def _append_log_file(
        self,
        timestamp: str,
        nome: str,
        documento: str,
        telefone: str,
        status: str,
        error: str,
        delay_used: float,
        template_id: str = "",
        session_id: str = "",
    ) -> None:
        log_file = os.path.join(LOG_DIR, f"log_{datetime.now().strftime('%Y%m%d')}.log")
        line = (
            f"[{timestamp}] session={session_id or '-'} | status={status:<12} | "
            f"template={template_id or '-':<20} | nome={str(nome):<30} | "
            f"documento={str(documento):<15} | telefone={str(telefone):<15} | "
            f"delay={delay_used:.1f}s | erro={error}\n"
        )
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(line)

    def save_session(
        self,
        session_id: str,
        session_start: str,
        total_sent: int,
        total_errors: int,
        total_skipped: int,
        total_simulated: int,
        avg_delay: float,
    ) -> None:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO session_stats (
                    session_id, session_start, session_end, total_sent,
                    total_errors, total_skipped, total_simulated, avg_delay
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(session_id) DO UPDATE SET
                    session_end=excluded.session_end,
                    total_sent=excluded.total_sent,
                    total_errors=excluded.total_errors,
                    total_skipped=excluded.total_skipped,
                    total_simulated=excluded.total_simulated,
                    avg_delay=excluded.avg_delay
                """,
                (
                    session_id,
                    session_start,
                    now,
                    total_sent,
                    total_errors,
                    total_skipped,
                    total_simulated,
                    avg_delay,
                ),
            )

    def count_sent_last_hour(self) -> int:
        one_hour_ago = (datetime.now() - timedelta(hours=1)).strftime("%Y-%m-%d %H:%M:%S")
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT COUNT(*) AS n
                FROM sends
                WHERE status='enviado' AND created_at >= ?
                """,
                (one_hour_ago,),
            ).fetchone()
        return int(row["n"]) if row else 0

    def count_sent_last_minutes(self, minutes: int = 10) -> int:
        since = (datetime.now() - timedelta(minutes=minutes)).strftime("%Y-%m-%d %H:%M:%S")
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT COUNT(*) AS n
                FROM sends
                WHERE status='enviado' AND created_at >= ?
                """,
                (since,),
            ).fetchone()
        return int(row["n"]) if row else 0

    def get_recent_history(self, limit: int = 300) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT
                    created_at, session_id, nome, documento, telefone,
                    template_id, status, error, delay_used, typing_delay,
                    preparation_status, placeholders_left, simulation
                FROM sends
                ORDER BY id DESC
                LIMIT ?
                """,
                (limit,),
            ).fetchall()
        return [dict(r) for r in rows]

    def get_metrics(self) -> dict[str, Any]:
        with self._connect() as conn:
            row = conn.execute(
                """
                SELECT
                    COUNT(*) as total,
                    SUM(CASE WHEN status='enviado' THEN 1 ELSE 0 END) as enviados,
                    SUM(CASE WHEN status='erro' THEN 1 ELSE 0 END) as erros,
                    SUM(CASE WHEN status='ignorado' THEN 1 ELSE 0 END) as ignorados,
                    SUM(CASE WHEN simulation=1 THEN 1 ELSE 0 END) as simulados,
                    AVG(CASE WHEN delay_used > 0 THEN delay_used END) as avg_delay
                FROM sends
                """
            ).fetchone()
        return dict(row) if row else {}

    def get_session_rows(self, session_id: str) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT
                    session_id, created_at, nome, documento, telefone, mensagem,
                    template_id, status, error, delay_used, typing_delay,
                    preparation_status, placeholders_left, context_json, simulation
                FROM sends
                WHERE session_id = ?
                ORDER BY id ASC
                """,
                (session_id,),
            ).fetchall()
        return [dict(r) for r in rows]