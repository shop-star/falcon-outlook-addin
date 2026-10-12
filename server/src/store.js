/* Remembers which emails have already been turned into Fergus jobs, keyed by
 * the email's internetMessageId (the same for every recipient and every
 * copy of the message, unlike Outlook's per-mailbox item ID).
 *
 * Each record also tracks progress through the create steps (customer →
 * site → job → note). If a later step fails, a retry reuses what was already
 * created instead of making a second customer or site.
 *
 * A JSON file is plenty for one office's email volume. Writes go to a temp
 * file then rename, so a crash mid-write can't corrupt it.
 */
import { mkdir, readFile, rename, writeFile } from "node:fs/promises";
import path from "node:path";

export class JobStore {
  constructor(dataDir) {
    this.file = path.join(dataDir, "actioned-emails.json");
    this.records = null;
    this.writing = Promise.resolve();
  }

  async load() {
    if (this.records) return;
    await mkdir(path.dirname(this.file), { recursive: true });
    try {
      this.records = JSON.parse(await readFile(this.file, "utf8"));
    } catch (err) {
      if (err.code !== "ENOENT") throw err;
      this.records = {};
    }
  }

  async get(messageId) {
    await this.load();
    return this.records[messageId] || null;
  }

  async update(messageId, changes) {
    await this.load();
    const now = new Date().toISOString();
    const existing = this.records[messageId] || { messageId, createdAt: now };
    this.records[messageId] = { ...existing, ...changes, updatedAt: now };
    await this.flush();
    return this.records[messageId];
  }

  flush() {
    // Chained so overlapping updates are written in order, never interleaved.
    // An earlier failed write mustn't block every later one.
    this.writing = this.writing.catch(() => {}).then(async () => {
      const tmp = `${this.file}.tmp`;
      await writeFile(tmp, JSON.stringify(this.records, null, 2));
      await rename(tmp, this.file);
    });
    return this.writing;
  }
}
