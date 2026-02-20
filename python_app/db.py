from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DATA_DIR / "contabilidad.db"


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    conn = get_conn()
    try:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS periods (
              period_key TEXT PRIMARY KEY,
              year INTEGER NOT NULL,
              month INTEGER NOT NULL,
              filename TEXT,
              exchange_rate REAL NOT NULL,
              uploaded_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS period_rows (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              period_key TEXT NOT NULL,
              row_id TEXT NOT NULL,
              sort_order INTEGER NOT NULL,
              code TEXT NOT NULL,
              name TEXT,
              sid REAL NOT NULL,
              sia REAL NOT NULL,
              cargos REAL NOT NULL,
              abonos REAL NOT NULL,
              sfd REAL NOT NULL,
              sfa REAL NOT NULL,
              is_new INTEGER NOT NULL DEFAULT 0,
              exclude_from_analysis INTEGER NOT NULL DEFAULT 0,
              UNIQUE(period_key, row_id)
            );

            CREATE TABLE IF NOT EXISTS period_manual_mappings (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              period_key TEXT NOT NULL,
              row_id TEXT NOT NULL,
              pgc TEXT,
              pgc_name TEXT,
              grupo TEXT,
              subgrupo TEXT,
              UNIQUE(period_key, row_id)
            );
            """
        )
        conn.commit()
    finally:
        conn.close()


def build_period_key(year: int, month: int) -> str:
    return f"{year}-{str(month).zfill(2)}"


def list_periods() -> list[dict[str, Any]]:
    conn = get_conn()
    try:
        cur = conn.execute(
            """
            SELECT
              p.period_key,
              p.year,
              p.month,
              p.filename,
              p.exchange_rate,
              p.uploaded_at,
              (SELECT COUNT(1) FROM period_rows r WHERE r.period_key = p.period_key) AS row_count
            FROM periods p
            ORDER BY p.year DESC, p.month DESC
            """
        )
        return [dict(r) for r in cur.fetchall()]
    finally:
        conn.close()


def save_period_data(
    *,
    year: int,
    month: int,
    filename: str,
    exchange_rate: float,
    rows: list[dict[str, Any]],
    manual_mappings: dict[str, dict[str, Any]],
    uploaded_at: str,
) -> None:
    period_key = build_period_key(year, month)
    conn = get_conn()
    try:
        conn.execute("BEGIN")
        conn.execute(
            """
            INSERT INTO periods(period_key, year, month, filename, exchange_rate, uploaded_at)
            VALUES(?, ?, ?, ?, ?, ?)
            ON CONFLICT(period_key) DO UPDATE SET
              filename = excluded.filename,
              exchange_rate = excluded.exchange_rate,
              uploaded_at = excluded.uploaded_at
            """,
            (period_key, year, month, filename, exchange_rate, uploaded_at),
        )
        conn.execute("DELETE FROM period_rows WHERE period_key = ?", (period_key,))
        conn.execute("DELETE FROM period_manual_mappings WHERE period_key = ?", (period_key,))

        for idx, row in enumerate(rows):
            conn.execute(
                """
                INSERT INTO period_rows(
                  period_key, row_id, sort_order, code, name, sid, sia, cargos, abonos, sfd, sfa, is_new, exclude_from_analysis
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    period_key,
                    row.get("_rowId"),
                    idx,
                    row.get("code"),
                    row.get("name"),
                    float(row.get("sid", 0) or 0),
                    float(row.get("sia", 0) or 0),
                    float(row.get("cargos", 0) or 0),
                    float(row.get("abonos", 0) or 0),
                    float(row.get("sfd", 0) or 0),
                    float(row.get("sfa", 0) or 0),
                    1 if row.get("_isNew") else 0,
                    1 if row.get("_excludeFromAnalysis") else 0,
                ),
            )

        for row_id, mapping in (manual_mappings or {}).items():
            conn.execute(
                """
                INSERT INTO period_manual_mappings(period_key, row_id, pgc, pgc_name, grupo, subgrupo)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    period_key,
                    row_id,
                    mapping.get("pgc"),
                    mapping.get("pgcName"),
                    mapping.get("grupo"),
                    mapping.get("subgrupo"),
                ),
            )

        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def load_period_data(year: int, month: int) -> dict[str, Any] | None:
    period_key = build_period_key(year, month)
    conn = get_conn()
    try:
        period = conn.execute("SELECT * FROM periods WHERE period_key = ?", (period_key,)).fetchone()
        if not period:
            return None

        rows_cur = conn.execute(
            """
            SELECT row_id, code, name, sid, sia, cargos, abonos, sfd, sfa, is_new, exclude_from_analysis
            FROM period_rows
            WHERE period_key = ?
            ORDER BY sort_order ASC
            """,
            (period_key,),
        )
        rows = []
        for r in rows_cur.fetchall():
            rows.append(
                {
                    "_rowId": r["row_id"],
                    "_isNew": bool(r["is_new"]),
                    "_excludeFromAnalysis": bool(r["exclude_from_analysis"]),
                    "code": r["code"],
                    "name": r["name"],
                    "sid": float(r["sid"] or 0),
                    "sia": float(r["sia"] or 0),
                    "cargos": float(r["cargos"] or 0),
                    "abonos": float(r["abonos"] or 0),
                    "sfd": float(r["sfd"] or 0),
                    "sfa": float(r["sfa"] or 0),
                }
            )

        mapping_cur = conn.execute(
            "SELECT row_id, pgc, pgc_name, grupo, subgrupo FROM period_manual_mappings WHERE period_key = ?",
            (period_key,),
        )
        manual_mappings: dict[str, dict[str, Any]] = {}
        for m in mapping_cur.fetchall():
            manual_mappings[m["row_id"]] = {
                "pgc": m["pgc"] or "",
                "pgcName": m["pgc_name"] or "",
                "grupo": m["grupo"] or "Sin clasificar",
                "subgrupo": m["subgrupo"] or "Sin clasificar",
            }

        return {
            "period": dict(period),
            "rows": rows,
            "manualMappings": manual_mappings,
        }
    finally:
        conn.close()
