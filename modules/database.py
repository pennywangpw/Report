import sqlite3
import os
from pathlib import Path

DB_PATH = Path(__file__).parent.parent / "data" / "financial.db"


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """初始化資料庫與 financial_master 資料表"""
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS financial_master (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            site_id         TEXT    NOT NULL,
            year_month      TEXT    NOT NULL,
            data_type       TEXT    NOT NULL DEFAULT 'Actual',
            revenue         REAL    DEFAULT 0,
            cogs            REAL    DEFAULT 0,
            opex            REAL    DEFAULT 0,
            gross_profit    REAL    GENERATED ALWAYS AS (revenue - cogs) STORED,
            created_at      TEXT    DEFAULT (datetime('now','localtime')),
            updated_at      TEXT    DEFAULT (datetime('now','localtime')),
            UNIQUE (site_id, year_month, data_type)
        )
    """)

    conn.commit()
    conn.close()


def upsert_records(records: list[dict]) -> dict:
    """
    批次 Upsert：相同 (site_id, year_month, data_type) 則覆蓋更新。
    回傳 {'inserted': int, 'updated': int, 'failed': int, 'errors': list}
    """
    conn = get_connection()
    cursor = conn.cursor()

    inserted = updated = failed = 0
    errors = []

    for rec in records:
        try:
            # 先查是否已存在
            cursor.execute(
                "SELECT id FROM financial_master WHERE site_id=? AND year_month=? AND data_type=?",
                (rec["site_id"], rec["year_month"], rec["data_type"]),
            )
            existing = cursor.fetchone()

            if existing:
                cursor.execute(
                    """
                    UPDATE financial_master
                    SET revenue=?, cogs=?, opex=?, updated_at=datetime('now','localtime')
                    WHERE site_id=? AND year_month=? AND data_type=?
                    """,
                    (
                        rec["revenue"], rec["cogs"], rec["opex"],
                        rec["site_id"], rec["year_month"], rec["data_type"],
                    ),
                )
                updated += 1
            else:
                cursor.execute(
                    """
                    INSERT INTO financial_master (site_id, year_month, data_type, revenue, cogs, opex)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        rec["site_id"], rec["year_month"], rec["data_type"],
                        rec["revenue"], rec["cogs"], rec["opex"],
                    ),
                )
                inserted += 1

        except Exception as e:
            failed += 1
            errors.append(f"Row {rec}: {e}")

    conn.commit()
    conn.close()
    return {"inserted": inserted, "updated": updated, "failed": failed, "errors": errors}


def get_available_options() -> list[str]:
    """回傳 DB 中所有可用的 'site_id / year_quarter' 選項（供模組 B 使用）"""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT DISTINCT site_id, year_month, data_type FROM financial_master ORDER BY site_id, year_month"
    )
    rows = cursor.fetchall()
    conn.close()

    options = []
    for row in rows:
        ym = row["year_month"]          # 格式如 2025-Q1 或 2025-01
        options.append(f"{row['site_id']} / {ym} / {row['data_type']}")
    return options
