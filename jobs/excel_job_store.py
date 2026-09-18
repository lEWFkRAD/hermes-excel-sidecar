"""Durable, single-lease store for resumable Excel jobs."""
from __future__ import annotations

import json
import sqlite3
import uuid
from datetime import datetime, timezone
from pathlib import Path

STATUSES = {"queued", "running", "paused", "failed", "completed", "needs_review"}


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


class ExcelJobStore:
    def __init__(self, database: str | Path):
        self.database = str(database)
        with self._connect() as db:
            db.execute("""CREATE TABLE IF NOT EXISTS excel_jobs (
                job_id TEXT PRIMARY KEY, workbook_id TEXT NOT NULL, sheet TEXT NOT NULL,
                input_range TEXT NOT NULL, output_ranges TEXT NOT NULL, cursor INTEGER NOT NULL,
                batch_size INTEGER NOT NULL, status TEXT NOT NULL, release TEXT NOT NULL,
                session_id TEXT NOT NULL, last_checkpoint TEXT NOT NULL, error TEXT,
                lease_id TEXT, lease_until REAL
            )""")

    def _connect(self):
        db = sqlite3.connect(self.database, timeout=10, isolation_level=None)
        db.row_factory = sqlite3.Row
        return db

    def create(self, *, workbook_id, sheet, input_range, output_ranges, cursor, batch_size, release, session_id, job_id=None):
        if not all(isinstance(v, str) and v for v in (workbook_id, sheet, input_range, release, session_id)):
            raise ValueError("job identity fields are required")
        if not isinstance(output_ranges, list) or not output_ranges or not all(isinstance(v, str) and v for v in output_ranges):
            raise ValueError("output_ranges must be a non-empty list")
        if not isinstance(cursor, int) or cursor < 1 or not isinstance(batch_size, int) or not 1 <= batch_size <= 250:
            raise ValueError("invalid cursor or batch size")
        job_id = job_id or f"job_{uuid.uuid4().hex}"
        with self._connect() as db:
            db.execute("INSERT INTO excel_jobs VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, NULL)",
                       (job_id, workbook_id, sheet, input_range, json.dumps(output_ranges), cursor, batch_size, "queued", release, session_id, _now(), None))
        return self.get(job_id)

    def get(self, job_id):
        with self._connect() as db:
            row = db.execute("SELECT * FROM excel_jobs WHERE job_id = ?", (job_id,)).fetchone()
        if row is None:
            return None
        result = dict(row)
        result["output_ranges"] = json.loads(result["output_ranges"])
        return result

    def checkpoint(self, job_id, *, cursor, status="running", error=None, lease_id=None):
        if status not in STATUSES or not isinstance(cursor, int) or cursor < 1:
            raise ValueError("invalid checkpoint")
        with self._connect() as db:
            db.execute("BEGIN IMMEDIATE")
            current = db.execute("SELECT lease_id FROM excel_jobs WHERE job_id = ?", (job_id,)).fetchone()
            if current is None or (lease_id is not None and current["lease_id"] != lease_id):
                db.rollback()
                raise ValueError("job lease mismatch")
            db.execute("UPDATE excel_jobs SET cursor=?, status=?, error=?, last_checkpoint=? WHERE job_id=?", (cursor, status, error, _now(), job_id))
            db.commit()
        return self.get(job_id)

    def lease(self, job_id, *, seconds=300):
        if seconds <= 0:
            raise ValueError("lease duration must be positive")
        now = datetime.now(timezone.utc).timestamp()
        lease_id = f"lease_{uuid.uuid4().hex}"
        with self._connect() as db:
            db.execute("BEGIN IMMEDIATE")
            row = db.execute("SELECT lease_id, lease_until, status FROM excel_jobs WHERE job_id=?", (job_id,)).fetchone()
            if row is None or (row["lease_id"] and (row["lease_until"] or 0) > now):
                db.rollback()
                raise ValueError("job is already leased or missing")
            db.execute("UPDATE excel_jobs SET lease_id=?, lease_until=?, status='running', last_checkpoint=? WHERE job_id=?", (lease_id, now + seconds, _now(), job_id))
            db.commit()
        return lease_id

    def recover_stale(self):
        now = datetime.now(timezone.utc).timestamp()
        with self._connect() as db:
            result = db.execute("UPDATE excel_jobs SET lease_id=NULL, lease_until=NULL, status=CASE WHEN status='running' THEN 'paused' ELSE status END WHERE lease_until IS NOT NULL AND lease_until <= ?", (now,))
        return result.rowcount
