/* Keeps New Zealand public holidays on the Fergus calendar a year ahead.
 *
 * Run monthly by the Worker's cron trigger (wrangler.jsonc). Each run looks
 * at the next twelve months, and creates a Fergus calendar event for every
 * holiday not already recorded in D1 (table holiday_events). The first run
 * adds about a year's worth; after that each run adds whatever has newly
 * come into range, usually one or two.
 */
import { holidaysBetween } from "./holidays.js";

const TIME_ZONE = "Pacific/Auckland";

export class HolidaySync {
  constructor({ fergus, db, region = "", now = () => new Date() }) {
    this.fergus = fergus;
    this.db = db;
    this.region = region;
    this.now = now;
  }

  async run() {
    const from = nzToday(this.now());
    const to = `${Number(from.slice(0, 4)) + 1}${from.slice(4)}`;
    const created = [];
    const skipped = [];

    for (const holiday of holidaysBetween(from, to, this.region)) {
      // Claim the holiday first: if the row's already there, it's been done
      // (or a previous run died mid-way and needs a person to check Fergus).
      const claim = await this.db
        .prepare("INSERT OR IGNORE INTO holiday_events (holiday_date, name) VALUES (?, ?)")
        .bind(holiday.date, holiday.name)
        .run();
      if (claim.meta.changes !== 1) {
        skipped.push(holiday);
        continue;
      }

      let event;
      try {
        event = await this.fergus.createCalendarEvent({ title: holiday.name, date: holiday.date });
      } catch (err) {
        // Free the claim so next month's run tries again, then stop: a bad
        // token or a rejected event would fail the same way for the rest.
        await this.db.prepare("DELETE FROM holiday_events WHERE holiday_date = ? AND name = ?").bind(holiday.date, holiday.name).run();
        err.message = `Adding ${holiday.name} (${holiday.date}) to the Fergus calendar failed: ${err.message}`;
        throw err;
      }
      await this.db
        .prepare("UPDATE holiday_events SET event_id = ? WHERE holiday_date = ? AND name = ?")
        .bind(event.id == null ? null : String(event.id), holiday.date, holiday.name)
        .run();
      created.push({ ...holiday, eventId: event.id });
    }

    return { from, to, created, skipped };
  }
}

// Today's date in New Zealand, "YYYY-MM-DD".
export function nzToday(now) {
  return new Intl.DateTimeFormat("en-CA", { timeZone: TIME_ZONE, year: "numeric", month: "2-digit", day: "2-digit" }).format(now);
}
