-- Public holidays already put on the Fergus calendar, so the monthly sync
-- adds each one only once. A row is written before the Fergus call (with
-- event_id NULL) and filled in after, so a run that dies part-way never
-- leads to a duplicate event.
CREATE TABLE IF NOT EXISTS holiday_events (
  holiday_date TEXT NOT NULL,
  name         TEXT NOT NULL,
  event_id     TEXT,
  created_at   TEXT NOT NULL DEFAULT (datetime('now')),
  PRIMARY KEY (holiday_date, name)
);
