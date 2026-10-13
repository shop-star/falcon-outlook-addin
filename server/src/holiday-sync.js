/* Keeps New Zealand public holidays on the Fergus calendar a year ahead.
 *
 * Run monthly by the Worker's cron trigger (wrangler.jsonc). Each run looks
 * at the next twelve months, and creates a Fergus calendar event for every
 * holiday not already recorded in D1 (table holiday_events): the national
 * holidays and every region's anniversary day. The first run adds about a
 * year's worth; after that each run adds whatever has newly come into range.
 *
 * Fergus events have start and end times rather than an all-day flag, so a
 * holiday is entered as the working day, 7am to 5pm New Zealand time.
 */
import { holidaysBetween } from "./holidays.js";

const TIME_ZONE = "Pacific/Auckland";
const DAY_START = 7;
const DAY_END = 17;

export class HolidaySync {
  // `userIds`: the first owns each event, the rest are linked to it.
  constructor({ fergus, db, userIds, eventType, now = () => new Date() }) {
    this.fergus = fergus;
    this.db = db;
    this.userIds = userIds;
    this.eventType = eventType;
    this.now = now;
  }

  async run() {
    if (!this.userIds.length) {
      throw new Error("FERGUS_HOLIDAY_USER_IDS isn't set, so there's no Fergus user to put the public holidays on. See the README.");
    }
    const [userId, ...linkedUserIds] = this.userIds;
    const from = nzToday(this.now());
    const to = `${Number(from.slice(0, 4)) + 1}${from.slice(4)}`;
    const created = [];
    const skipped = [];

    for (const holiday of holidaysBetween(from, to)) {
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
        event = await this.fergus.createCalendarEvent({
          userId,
          linkedUserIds,
          eventType: this.eventType,
          title: holiday.name,
          description: holiday.region
            ? `Regional anniversary day: a public holiday in ${holiday.region} only. Added automatically.`
            : "New Zealand public holiday. Added automatically.",
          startTime: nzTimeToIso(holiday.date, DAY_START),
          endTime: nzTimeToIso(holiday.date, DAY_END),
        });
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

// The instant it's `hour`:00 on `date` ("YYYY-MM-DD") in New Zealand, as an
// ISO string. Daylight saving changes at 2–3am, never during the day.
export function nzTimeToIso(date, hour) {
  const [y, m, d] = date.split("-").map(Number);
  const asUtc = Date.UTC(y, m - 1, d, hour);
  const parts = Object.fromEntries(
    new Intl.DateTimeFormat("en-US", { timeZone: TIME_ZONE, hourCycle: "h23", year: "numeric", month: "numeric", day: "numeric", hour: "numeric" })
      .formatToParts(new Date(asUtc))
      .map((p) => [p.type, Number(p.value)])
  );
  const offset = Date.UTC(parts.year, parts.month - 1, parts.day, parts.hour) - asUtc;
  return new Date(asUtc - offset).toISOString();
}
