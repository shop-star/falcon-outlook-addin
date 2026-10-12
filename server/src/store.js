/* Remembers which emails have already been turned into Fergus jobs, keyed by
 * the email's internetMessageId (the same for every recipient and every
 * copy of the message, unlike Outlook's per-mailbox item ID).
 *
 * Each record also tracks progress through the create steps (customer →
 * site → job → note). If a later step fails, a retry reuses what was already
 * created instead of making a second customer or site.
 *
 * Stored in Cloudflare D1 (see migrations/). A Worker runs as many copies
 * at once, so the "someone's already creating this job" lock lives in the
 * database too: acquire() is a single conditional UPDATE, which D1 runs
 * atomically.
 */

const LOCK_MS = 2 * 60 * 1000; // longer than any create takes; frees itself if a Worker dies mid-way

export class JobStore {
  constructor(db, now = () => Date.now()) {
    this.db = db;
    this.now = now;
  }

  async get(messageId) {
    const row = await this.db.prepare("SELECT data FROM actioned_emails WHERE message_id = ?").bind(messageId).first();
    // acquire() can create a row before anything's been recorded in it.
    const data = row ? JSON.parse(row.data) : {};
    return Object.keys(data).length ? data : null;
  }

  async update(messageId, changes) {
    const now = new Date(this.now()).toISOString();
    const existing = (await this.get(messageId)) || { messageId, createdAt: now };
    const record = { ...existing, ...changes, updatedAt: now };
    await this.db
      .prepare(
        "INSERT INTO actioned_emails (message_id, data, updated_at) VALUES (?1, ?2, ?3) " +
          "ON CONFLICT (message_id) DO UPDATE SET data = ?2, updated_at = ?3"
      )
      .bind(messageId, JSON.stringify(record), now)
      .run();
    return record;
  }

  // True if this caller now holds the lock for messageId.
  async acquire(messageId) {
    const now = this.now();
    await this.db.prepare("INSERT OR IGNORE INTO actioned_emails (message_id) VALUES (?)").bind(messageId).run();
    const result = await this.db
      .prepare("UPDATE actioned_emails SET lock_until = ?1 WHERE message_id = ?2 AND (lock_until IS NULL OR lock_until < ?3)")
      .bind(now + LOCK_MS, messageId, now)
      .run();
    return result.meta.changes === 1;
  }

  async release(messageId) {
    await this.db.prepare("UPDATE actioned_emails SET lock_until = NULL WHERE message_id = ?").bind(messageId).run();
  }
}
