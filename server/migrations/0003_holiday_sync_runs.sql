-- One row per public-holiday sync run, so its outcome can be seen at
-- GET /api/holidays without digging through the Worker's logs.
CREATE TABLE IF NOT EXISTS holiday_sync_runs (
  id     INTEGER PRIMARY KEY AUTOINCREMENT,
  ran_at TEXT NOT NULL DEFAULT (datetime('now')),
  ok     INTEGER NOT NULL,
  result TEXT NOT NULL
);
