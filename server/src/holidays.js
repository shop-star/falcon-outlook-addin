/* New Zealand public holidays, worked out from the rules in the Holidays Act
 * 2003 rather than fetched from a web service, so the monthly sync has
 * nothing extra to break.
 *
 * Dates are the days the holiday is *observed* — Waitangi Day on a Saturday
 * is taken on the Monday, Christmas on a Sunday on the Tuesday, and so on.
 * Each holiday is { date: "YYYY-MM-DD", name }.
 */

// Matariki moves each year; its dates are fixed in law up to 2052 (Te Kāhui
// o Matariki Public Holiday Act 2022, Schedule 1).
const MATARIKI = {
  2022: "06-24", 2023: "07-14", 2024: "06-28", 2025: "06-20", 2026: "07-10",
  2027: "06-25", 2028: "07-14", 2029: "07-06", 2030: "06-21", 2031: "07-11",
  2032: "07-02", 2033: "06-24", 2034: "07-07", 2035: "06-29", 2036: "07-18",
  2037: "07-10", 2038: "06-25", 2039: "07-15", 2040: "07-06", 2041: "07-19",
  2042: "07-11", 2043: "07-03", 2044: "06-24", 2045: "07-07", 2046: "06-29",
  2047: "07-19", 2048: "07-03", 2049: "06-25", 2050: "07-15", 2051: "06-30",
  2052: "06-21",
};

// Regional anniversary days, as the government observes them. Keys are what
// HOLIDAY_REGION accepts.
const ANNIVERSARIES = {
  auckland: { name: "Auckland Anniversary Day", date: (y) => mondayNearest(day(y, 1, 29)) },
  northland: { name: "Auckland Anniversary Day", date: (y) => mondayNearest(day(y, 1, 29)) },
  wellington: { name: "Wellington Anniversary Day", date: (y) => mondayNearest(day(y, 1, 22)) },
  nelson: { name: "Nelson Anniversary Day", date: (y) => mondayNearest(day(y, 2, 1)) },
  taranaki: { name: "Taranaki Anniversary Day", date: (y) => nthWeekday(y, 3, MON, 2) },
  otago: {
    name: "Otago Anniversary Day",
    // Moves to the Tuesday when it would land on Easter Monday.
    date: (y) => {
      const d = mondayNearest(day(y, 3, 23));
      return sameDay(d, addDays(easterSunday(y), 1)) ? addDays(d, 1) : d;
    },
  },
  southland: { name: "Southland Anniversary Day", date: (y) => addDays(easterSunday(y), 2) },
  "hawkes-bay": { name: "Hawke's Bay Anniversary Day", date: (y) => addDays(labourDay(y), -3) },
  marlborough: { name: "Marlborough Anniversary Day", date: (y) => addDays(labourDay(y), 7) },
  canterbury: { name: "Canterbury Anniversary Day (Show Day)", date: (y) => addDays(nthWeekday(y, 11, TUE, 1), 10) },
  "chatham-islands": { name: "Chatham Islands Anniversary Day", date: (y) => mondayNearest(day(y, 11, 30)) },
  westland: { name: "Westland Anniversary Day", date: (y) => mondayNearest(day(y, 12, 1)) },
};

export const HOLIDAY_REGIONS = Object.keys(ANNIVERSARIES);

const SUN = 0;
const MON = 1;
const TUE = 2;
const FRI = 5;
const SAT = 6;

// Every public holiday observed in `year`, in date order. `region` adds that
// region's anniversary day; leave it empty for the national holidays only.
export function nzPublicHolidays(year, region = "") {
  if (region && !ANNIVERSARIES[region]) {
    throw new Error(`Unknown HOLIDAY_REGION "${region}". Use one of: ${HOLIDAY_REGIONS.join(", ")}`);
  }

  const easter = easterSunday(year);
  const [newYear, dayAfterNewYear] = mondayisePair(day(year, 1, 1));
  const [christmas, boxingDay] = mondayisePair(day(year, 12, 25));
  const holidays = [
    { date: newYear, name: "New Year's Day" },
    { date: dayAfterNewYear, name: "Day after New Year's Day" },
    { date: mondayise(day(year, 2, 6)), name: "Waitangi Day" },
    { date: addDays(easter, -2), name: "Good Friday" },
    { date: addDays(easter, 1), name: "Easter Monday" },
    { date: mondayise(day(year, 4, 25)), name: "ANZAC Day" },
    { date: nthWeekday(year, 6, MON, 1), name: "King's Birthday" },
    { date: labourDay(year), name: "Labour Day" },
    { date: christmas, name: "Christmas Day" },
    { date: boxingDay, name: "Boxing Day" },
  ];
  if (MATARIKI[year]) {
    holidays.push({ date: new Date(`${year}-${MATARIKI[year]}T00:00:00Z`), name: "Matariki" });
  } else {
    console.warn(`Matariki's date for ${year} isn't set in law yet; it's been left out.`);
  }
  if (region) {
    const anniversary = ANNIVERSARIES[region];
    holidays.push({ date: anniversary.date(year), name: anniversary.name });
  }

  return holidays
    .map((h) => ({ date: isoDate(h.date), name: h.name }))
    .sort((a, b) => a.date.localeCompare(b.date));
}

// Holidays observed from `from` to `to` inclusive, both "YYYY-MM-DD".
export function holidaysBetween(from, to, region = "") {
  const firstYear = Number(from.slice(0, 4));
  const lastYear = Number(to.slice(0, 4));
  const all = [];
  for (let year = firstYear; year <= lastYear; year++) all.push(...nzPublicHolidays(year, region));
  return all.filter((h) => h.date >= from && h.date <= to);
}

// ---------------------------------------------------------------------------
// Date helpers. Calendar dates are held as midnight-UTC Dates so no time zone
// or daylight-saving change can shift a day.

function day(year, month, date) {
  return new Date(Date.UTC(year, month - 1, date));
}

function addDays(d, n) {
  return new Date(d.getTime() + n * 86400000);
}

function sameDay(a, b) {
  return a.getTime() === b.getTime();
}

export function isoDate(d) {
  return d.toISOString().slice(0, 10);
}

// Waitangi and ANZAC Day: a weekend date is taken on the following Monday.
function mondayise(d) {
  const wd = d.getUTCDay();
  if (wd === SAT) return addDays(d, 2);
  if (wd === SUN) return addDays(d, 1);
  return d;
}

// New Year / Christmas and the day after: two consecutive holidays, each
// moved to the next free weekday when it lands on a weekend.
function mondayisePair(first) {
  const second = addDays(first, 1);
  switch (first.getUTCDay()) {
    case FRI: // Fri, Sat → Fri, Mon
      return [first, addDays(first, 3)];
    case SAT: // Sat, Sun → Mon, Tue
      return [addDays(first, 2), addDays(first, 3)];
    case SUN: // Sun, Mon → Tue, Mon
      return [addDays(first, 2), second];
    default:
      return [first, second];
  }
}

// Fri–Sun go forward to the next Monday, Tue–Thu back to the previous one.
function mondayNearest(d) {
  const offset = [1, 0, -1, -2, -3, 3, 2][d.getUTCDay()];
  return addDays(d, offset);
}

// The nth `weekday` (0 = Sunday) of month `month` (1-12).
function nthWeekday(year, month, weekday, n) {
  const first = day(year, month, 1);
  const offset = (weekday - first.getUTCDay() + 7) % 7;
  return addDays(first, offset + (n - 1) * 7);
}

// Fourth Monday in October.
function labourDay(year) {
  return nthWeekday(year, 10, MON, 4);
}

// Western (Gregorian) Easter Sunday — the "anonymous Gregorian" algorithm.
function easterSunday(year) {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const date = ((h + l - 7 * m + 114) % 31) + 1;
  return day(year, month, date);
}
