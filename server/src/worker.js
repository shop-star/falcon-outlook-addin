/* Cloudflare Worker: the HTTP API for the "Create Fergus job" task pane.
 *
 *   GET  /api/health    – liveness check, no auth
 *   POST /api/analyze   – { email } → extraction + match candidates, or the existing job
 *   POST /api/jobs      – { email, customer, site, jobType, group, title, description } → the job
 *   GET  /api/holidays      – recent public-holiday sync runs and the holidays added, no auth
 *   GET  /api/holidays/run  – run the public-holiday sync now and show the result, no auth
 *
 * It also runs on a monthly cron trigger (wrangler.jsonc) to keep New
 * Zealand public holidays on the Fergus calendar a year ahead
 * (holiday-sync.js).
 *
 * Settings come from the Worker's environment (wrangler.jsonc vars plus
 * secrets set in the Cloudflare dashboard); the actioned-emails record is
 * the D1 database bound as DB.
 */
import Anthropic from "@anthropic-ai/sdk";
import { createAuthenticator } from "./auth.js";
import { loadConfig } from "./config.js";
import { AppError } from "./errors.js";
import { FergusClient } from "./fergus.js";
import { HolidaySync, holidayStatus, recordRun } from "./holiday-sync.js";
import { JobService } from "./jobs.js";
import { JobStore } from "./store.js";

const MAX_REQUEST_BYTES = 2 * 1024 * 1024;

// Built once per Worker instance and reused across requests, so Microsoft's
// signing keys are fetched once rather than on every call.
let cached = null;

function setup(env) {
  if (cached?.env === env) return cached;
  const config = loadConfig(env);
  const fergus = new FergusClient({ apiToken: config.fergus.apiToken, baseUrl: config.fergus.baseUrl });
  cached = {
    env,
    config,
    fergus,
    authenticate: createAuthenticator(config.auth),
    service: new JobService({
      config,
      fergus,
      anthropic: new Anthropic({ apiKey: config.anthropic.apiKey }),
      store: new JobStore(env.DB),
    }),
  };
  return cached;
}

export function createHandler(getApp) {
  const routes = {
    "POST /api/analyze": (service, body) => service.analyze(body.email),
    "POST /api/jobs": (service, body, user) => service.create(body, user),
  };

  return async function handle(request, env) {
    const url = new URL(request.url);
    const headers = new Headers();

    let app;
    try {
      app = getApp(env);
    } catch (err) {
      // Missing settings: say so plainly rather than a bare 500.
      console.error(err);
      return json(500, { error: { code: "MISCONFIGURED", message: "The Fergus job server isn't fully set up yet. Tell IT." } }, headers);
    }

    const origin = request.headers.get("Origin");
    if (origin && app.config.allowedOrigins.includes(origin)) {
      headers.set("Access-Control-Allow-Origin", origin);
      headers.set("Vary", "Origin");
      headers.set("Access-Control-Allow-Headers", "Authorization, Content-Type");
      headers.set("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
      headers.set("Access-Control-Max-Age", "600");
    }

    try {
      if (request.method === "OPTIONS") return new Response(null, { status: 204, headers });
      if (request.method === "GET" && url.pathname === "/api/health") return json(200, { ok: true }, headers);
      if (request.method === "GET" && url.pathname === "/api/holidays") return json(200, await holidayStatus(env.DB), headers);
      // A GET so it can be run from a browser. No sign-in: a run only adds
      // holidays not already added, so running it again does no harm.
      if (request.method === "GET" && url.pathname === "/api/holidays/run") {
        try {
          return json(200, await syncHolidays(() => app.fergus, env.DB), headers);
        } catch (err) {
          return json(502, { ok: false, error: err.message, details: err.details }, headers);
        }
      }

      const handler = routes[`${request.method} ${url.pathname}`];
      if (!handler) throw new AppError(404, "NOT_FOUND", "Not found.");

      const user = await app.authenticate(request);
      const body = await readJson(request);
      return json(200, await handler(app.service, body, user), headers);
    } catch (err) {
      if (err instanceof AppError) {
        if (err.retryAfter) headers.set("Retry-After", String(err.retryAfter));
        return json(err.status, err.toJSON(), headers);
      }
      console.error(err);
      return json(500, { error: { code: "INTERNAL", message: "Something went wrong on the server. Try again, and tell IT if it keeps happening." } }, headers);
    }
  };
}

function json(status, body, headers) {
  headers.set("Content-Type", "application/json; charset=utf-8");
  return new Response(JSON.stringify(body), { status, headers });
}

async function readJson(request) {
  const text = await request.text();
  if (text.length > MAX_REQUEST_BYTES) throw new AppError(413, "EMAIL_TOO_LARGE", "This email is too large to send to the server.");
  try {
    return JSON.parse(text || "{}");
  } catch {
    throw new AppError(400, "BAD_REQUEST", "The request wasn't valid JSON.");
  }
}

export function createScheduled(getApp) {
  return async function scheduled(controller, env) {
    await syncHolidays(() => getApp(env).fergus, env.DB);
  };
}

// Runs the sync and records the outcome for GET /api/holidays, including a
// failure to set up the Fergus client.
async function syncHolidays(getFergus, db) {
  try {
    const result = await new HolidaySync({ fergus: getFergus(), db }).run();
    const added = result.created.map((h) => `${h.date} ${h.name}`).join("; ") || "nothing new";
    console.log(`Public holidays ${result.from} to ${result.to}: added ${added}.`);
    const summary = { ok: true, ...result, skipped: result.skipped.length };
    await recordRun(db, true, summary);
    return summary;
  } catch (err) {
    console.error(err);
    await recordRun(db, false, { error: err.message, details: err.details });
    throw err;
  }
}

export default { fetch: createHandler(setup), scheduled: createScheduled(setup) };
