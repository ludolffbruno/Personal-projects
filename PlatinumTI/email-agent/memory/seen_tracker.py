"""
memory/seen_tracker.py
Rastreador de e-mails já processados.
Usa SQLite local para garantir que o agente NÃO reprocesse e-mails antigos.
"""
import sqlite3
from pathlib import Path
from datetime import datetime


DB_PATH = Path(__file__).parent.parent / "seen_emails.db"


def _get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(str(DB_PATH))
    conn.execute("""
        CREATE TABLE IF NOT EXISTS seen_emails (
            message_id TEXT PRIMARY KEY,
            subject TEXT,
            sender TEXT,
            seen_at TEXT,
            is_quote INTEGER,
            draft_created INTEGER
        )
    """)
    conn.commit()
    return conn


def is_seen(message_id: str) -> bool:
    """Verifica se este e-mail já foi processado antes."""
    with _get_conn() as conn:
        row = conn.execute(
            "SELECT 1 FROM seen_emails WHERE message_id = ?", (message_id,)
        ).fetchone()
        return row is not None


def mark_as_seen(
    message_id: str,
    subject: str = "",
    sender: str = "",
    is_quote: bool = False,
    draft_created: bool = False,
):
    """Registra este e-mail como processado."""
    with _get_conn() as conn:
        conn.execute(
            """INSERT OR IGNORE INTO seen_emails
               (message_id, subject, sender, seen_at, is_quote, draft_created)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (
                message_id,
                subject,
                sender,
                datetime.now().isoformat(),
                int(is_quote),
                int(draft_created),
            ),
        )
        conn.commit()


def get_stats() -> dict:
    """Retorna estatísticas do tracker."""
    with _get_conn() as conn:
        total = conn.execute("SELECT COUNT(*) FROM seen_emails").fetchone()[0]
        quotes = conn.execute(
            "SELECT COUNT(*) FROM seen_emails WHERE is_quote = 1"
        ).fetchone()[0]
        drafted = conn.execute(
            "SELECT COUNT(*) FROM seen_emails WHERE draft_created = 1"
        ).fetchone()[0]
        return {"total_seen": total, "quotes": quotes, "drafts_created": drafted}
