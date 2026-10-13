import assert from "node:assert/strict";
import { test } from "node:test";
import { FergusClient } from "../src/fergus.js";
import { HolidaySync, nzTimeToIso, nzToday } from "../src/holiday-sync.js";
import { holidaysBetween, nzPublicHolidays } from "../src/holidays.js";
import { createHandler, createScheduled } from "../src/worker.js";
import { testD1 } from "./d1.js";
import { fakeFergus } from "./fakes.js";

const dates = (holidays) => Object.fromEntries(holidays.map((h) => [h.name, h.date]));
const national = (year) => nzPublicHolidays(year).filter((h) => !h.region);

test("2027 national holidays, checked against the government's list", () => {
  assert.deepEqual(dates(national(2027)), {
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
  const d = dates(national(2022));
  assert.equal(d["Christmas Day"], "2022-12-27");
  assert.equal(d["Boxing Day"], "2022-12-26");
  assert.equal(d["Good Friday"], "2022-04-15");
});

test("every regional anniversary day is included", () => {
  const regional = nzPublicHolidays(2026).filter((h) => h.region);
  assert.equal(regional.length, 12);
  const d = dates(regional);
  assert.equal(d["Wellington Anniversary Day"], "2026-01-19");
  assert.equal(d["Auckland Anniversary Day"], "2026-01-26");
  assert.equal(d["Otago Anniversary Day"], "2026-03-23");
  assert.equal(d["Southland Anniversary Day"], "2026-04-07");
  assert.equal(d["South Canterbury Anniversary Day"], "2026-09-28");
  assert.equal(d["Hawke's Bay Anniversary Day"], "2026-10-23");
  assert.equal(d["Marlborough Anniversary Day"], "2026-11-02");
  assert.equal(d["Canterbury Anniversary Day (Show Day)"], "2026-11-13");
  assert.equal(d["Westland Anniversary Day"], "2026-11-30");
  assert.equal(nzPublicHolidays(2027).find((h) => h.name === "Auckland Anniversary Day").date, "2027-02-01");
});

test("holidaysBetween spans a year boundary", () => {
  const names = holidaysBetween("2026-12-01", "2027-01-31").map((h) => `${h.date} ${h.name}`);
  assert.deepEqual(names, [
    "2026-12-25 Christmas Day",
    "2026-12-28 Boxing Day",
    "2027-01-01 New Year's Day",
    "2027-01-04 Day after New Year's Day",
    "2027-01-25 Wellington Anniversary Day",
  ]);
});

test("New Zealand dates and times, across daylight saving", () => {
  assert.equal(nzToday(new Date("2026-10-13T20:00:00Z")), "2026-10-14");
  assert.equal(nzTimeToIso("2026-12-25", 7), "2026-12-24T18:00:00.000Z"); // NZDT, +13
  assert.equal(nzTimeToIso("2027-06-25", 17), "2027-06-25T05:00:00.000Z"); // NZST, +12
});

function syncSetup(now, failures) {
  const fake = fakeFergus({ failures });
  const fergus = new FergusClient({ apiToken: "t", baseUrl: "https://fergus.test", fetchImpl: fake.fetchImpl, sleep: async () => {} });
  const db = testD1();
  return { fake, db, sync: (at = now) => new HolidaySync({ fergus, db, now: () => at }).run() };
}

test("sync adds the next year's holidays once, then only new ones", async () => {
  const { fake, sync } = syncSetup(new Date("2026-10-13T00:00:00Z"));

  const first = await sync();
  assert.equal(first.from, "2026-10-13");
  assert.equal(first.to, "2027-10-13");
  assert.deepEqual(first.created.slice(0, 3).map((h) => h.date), ["2026-10-23", "2026-10-26", "2026-11-02"]);
  assert.equal(first.created.at(-1).name, "South Canterbury Anniversary Day"); // 2027-09-27
  assert.equal(first.created.length, 23); // 11 national + 12 regional
  const [hawkesBay, labourDay] = fake.calls.filter((c) => c.path === "/calendarEvents").map((c) => c.body);
  assert.deepEqual(labourDay, {
    startTime: "2026-10-25T18:00:00.000Z",
    endTime: "2026-10-26T04:00:00.000Z",
    eventTitle: "Labour Day",
    eventType: "OTHER",
    description: "New Zealand public holiday. Added automatically.",
    frequency: "NEVER",
  });
  assert.equal(hawkesBay.eventTitle, "Hawke's Bay Anniversary Day");
  assert.match(hawkesBay.description, /public holiday in Hawke's Bay only/);

  const again = await sync();
  assert.equal(again.created.length, 0);

  const nextMonth = await sync(new Date("2026-11-01T19:00:00Z"));
  assert.deepEqual(nextMonth.created.map((h) => h.name), ["Hawke's Bay Anniversary Day", "Labour Day", "Marlborough Anniversary Day"]); // 2027's
  assert.equal(fake.calls.filter((c) => c.path === "/calendarEvents").length, 26);
});

test("a Fergus failure stops the run and is retried next time", async () => {
  const { fake, sync } = syncSetup(new Date("2026-10-13T00:00:00Z"), {
    "POST /calendarEvents": [null, { status: 422, body: { message: "eventType is invalid" } }],
  });

  await assert.rejects(sync(), /Labour Day.*eventType is invalid/);
  const retry = await sync();
  assert.equal(retry.created[0].name, "Labour Day");
  assert.equal(retry.skipped.length, 1); // Hawke's Bay, added by the first run
  assert.equal(fake.calls.filter((c) => c.path === "/calendarEvents").length, 24);
});

test("each scheduled run's outcome shows at GET /api/holidays", async () => {
  const fake = fakeFergus({ failures: { "POST /calendarEvents": [{ status: 404, body: { message: "Route not found" } }] } });
  const fergus = new FergusClient({ apiToken: "t", baseUrl: "https://fergus.test", fetchImpl: fake.fetchImpl, sleep: async () => {} });
  const env = { DB: testD1() };
  const scheduled = createScheduled(() => ({ fergus }));
  const handle = createHandler(() => ({ config: { allowedOrigins: [] }, fergus }));
  const status = async () => (await handle(new Request("https://worker.test/api/holidays"), env)).json();

  const error = console.error;
  console.error = () => {};
  try {
    await assert.rejects(scheduled({}, env), /Route not found/);
  } finally {
    console.error = error;
  }
  let body = await status();
  assert.equal(body.runs[0].ok, false);
  assert.match(body.runs[0].error, /couldn't find a record.*Route not found/);
  assert.deepEqual(body.added, []);

  const log = console.log;
  console.log = () => {};
  try {
    await scheduled({}, env);
  } finally {
    console.log = log;
  }
  body = await status();
  assert.equal(body.runs[0].ok, true);
  assert.equal(body.runs[0].created.length, body.added.length);
  assert.equal(body.runs[1].ok, false);

  // Run from a browser: everything's already added, so nothing new.
  console.log = () => {};
  try {
    const res = await handle(new Request("https://worker.test/api/holidays/run"), env);
    assert.equal(res.status, 200);
    const run = await res.json();
    assert.equal(run.ok, true);
    assert.deepEqual(run.created, []);
    assert.equal(run.skipped, body.added.length);
  } finally {
    console.log = log;
  }
  assert.equal((await status()).runs.length, 3);
});
