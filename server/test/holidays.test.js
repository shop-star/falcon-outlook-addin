import assert from "node:assert/strict";
import { test } from "node:test";
import { FergusClient } from "../src/fergus.js";
import { HolidaySync, nzToday } from "../src/holiday-sync.js";
import { holidaysBetween, nzPublicHolidays } from "../src/holidays.js";
import { testD1 } from "./d1.js";
import { fakeFergus } from "./fakes.js";

const dates = (holidays) => Object.fromEntries(holidays.map((h) => [h.name, h.date]));

test("2027 national holidays, checked against the government's list", () => {
  assert.deepEqual(dates(nzPublicHolidays(2027)), {
    "New Year's Day": "2027-01-01",
    "Day after New Year's Day": "2027-01-04", // Sat 2nd → Mon
    "Waitangi Day": "2027-02-08", // Sat → Mon
    "Good Friday": "2027-03-26",
    "Easter Monday": "2027-03-29",
    "ANZAC Day": "2027-04-26", // Sun → Mon
    "King's Birthday": "2027-06-07",
    Matariki: "2027-06-25",
    "Labour Day": "2027-10-25",
    "Christmas Day": "2027-12-27", // Sat → Mon
    "Boxing Day": "2027-12-28", // Sun → Tue
  });
});

test("Christmas on a Sunday is taken on the Tuesday", () => {
  const d = dates(nzPublicHolidays(2022));
  assert.equal(d["Christmas Day"], "2022-12-27");
  assert.equal(d["Boxing Day"], "2022-12-26");
  assert.equal(d["Good Friday"], "2022-04-15");
});

test("regional anniversary days", () => {
  assert.equal(dates(nzPublicHolidays(2027, "auckland"))["Auckland Anniversary Day"], "2027-02-01");
  assert.equal(dates(nzPublicHolidays(2026, "wellington"))["Wellington Anniversary Day"], "2026-01-19");
  assert.equal(dates(nzPublicHolidays(2026, "canterbury"))["Canterbury Anniversary Day (Show Day)"], "2026-11-13");
  assert.equal(dates(nzPublicHolidays(2026, "hawkes-bay"))["Hawke's Bay Anniversary Day"], "2026-10-23");
  assert.equal(nzPublicHolidays(2026).length, 11);
  assert.throws(() => nzPublicHolidays(2026, "atlantis"), /HOLIDAY_REGION/);
});

test("holidaysBetween spans a year boundary", () => {
  const names = holidaysBetween("2026-12-01", "2027-01-31").map((h) => `${h.date} ${h.name}`);
  assert.deepEqual(names, ["2026-12-25 Christmas Day", "2026-12-28 Boxing Day", "2027-01-01 New Year's Day", "2027-01-04 Day after New Year's Day"]);
});

test("nzToday uses New Zealand's date, not UTC's", () => {
  assert.equal(nzToday(new Date("2026-10-13T20:00:00Z")), "2026-10-14");
});

function syncSetup(now, failures) {
  const fake = fakeFergus({ failures });
  const fergus = new FergusClient({ apiToken: "t", baseUrl: "https://fergus.test", fetchImpl: fake.fetchImpl, sleep: async () => {} });
  const db = testD1();
  return { fake, db, sync: (at = now) => new HolidaySync({ fergus, db, region: "auckland", now: () => at }).run() };
}

test("sync adds the next year's holidays once, then only new ones", async () => {
  const { fake, sync } = syncSetup(new Date("2026-10-13T00:00:00Z"));

  const first = await sync();
  assert.equal(first.from, "2026-10-13");
  assert.equal(first.to, "2027-10-13");
  assert.equal(first.created[0].date, "2026-10-26"); // Labour Day 2026
  assert.equal(first.created.at(-1).date, "2027-06-25"); // Matariki; Labour Day 2027 is the 25th, out of range
  assert.equal(first.created.length, 12);
  const event = fake.calls.find((c) => c.path === "/calendarEvents").body;
  assert.equal(event.title, "Labour Day");
  assert.equal(event.startDate, "2026-10-26");
  assert.equal(event.isAllDay, true);

  const again = await sync();
  assert.equal(again.created.length, 0);

  const nextMonth = await sync(new Date("2026-11-01T19:00:00Z"));
  assert.deepEqual(nextMonth.created.map((h) => h.name), ["Labour Day"]); // 2027's
  assert.equal(fake.calls.filter((c) => c.path === "/calendarEvents").length, 13);
});

test("a Fergus failure stops the run and is retried next time", async () => {
  const { fake, sync } = syncSetup(new Date("2026-10-13T00:00:00Z"), {
    "POST /calendarEvents": [null, { status: 422, body: { message: "eventType is invalid" } }],
  });

  await assert.rejects(sync(), /Christmas Day.*eventType is invalid/);
  const retry = await sync();
  assert.equal(retry.created[0].name, "Christmas Day");
  assert.equal(retry.skipped.length, 1); // Labour Day, added by the first run
  assert.equal(fake.calls.filter((c) => c.path === "/calendarEvents").length, 13);
});
