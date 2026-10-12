-- Which emails have been turned into Fergus jobs, keyed by internetMessageId.
-- `data` holds the progress record as JSON (customerId, siteId, jobId, ...).
-- `lock_until` stops two people creating a job from the same email at once.
CREATE TABLE IF NOT EXISTS actioned_emails (
  message_id TEXT PRIMARY KEY,
  data       TEXT NOT NULL DEFAULT '{}',
  lock_until INTEGER,
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);
