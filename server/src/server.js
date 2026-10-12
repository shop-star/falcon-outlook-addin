/* HTTP API for the "Create Fergus job" task pane.
 *
 *   GET  /api/health    – liveness check, no auth
 *   POST /api/analyze   – { email } → extraction + match candidates, or the existing job
 *   POST /api/jobs      – { email, customer, site, jobType, group, title, description } → the job
 *
 * Plain node:http rather than a framework: three routes don't need one.
 */
import http from "node:http";
import { pathToFileURL } from "node:url";
import Anthropic from "@anthropic-ai/sdk";
import { createAuthenticator } from "./auth.js";
import { loadConfig } from "./config.js";
import { AppError } from "./errors.js";
import { FergusClient } from "./fergus.js";
import { JobService } from "./jobs.js";
import { JobStore } from "./store.js";

const MAX_REQUEST_BYTES = 2 * 1024 * 1024;

export function createServer({ config, service, authenticate }) {
  const routes = {
    "POST /api/analyze": async (body) => service.analyze(body.email),
    "POST /api/jobs": async (body, user) => service.create(body, user),
  };

  return http.createServer(async (req, res) => {
    const origin = req.headers.origin;
    if (origin && config.allowedOrigins.includes(origin)) {
      res.setHeader("Access-Control-Allow-Origin", origin);
      res.setHeader("Vary", "Origin");
      res.setHeader("Access-Control-Allow-Headers", "Authorization, Content-Type");
      res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
      res.setHeader("Access-Control-Max-Age", "600");
    }

    try {
      const url = new URL(req.url, "http://localhost");
      if (req.method === "OPTIONS") return send(res, 204);
      if (req.method === "GET" && url.pathname === "/api/health") return send(res, 200, { ok: true });

      const handler = routes[`${req.method} ${url.pathname}`];
      if (!handler) throw new AppError(404, "NOT_FOUND", "Not found.");

      const user = await authenticate(req);
      const body = await readJson(req);
      send(res, 200, await handler(body, user));
    } catch (err) {
      if (err instanceof AppError) {
        if (err.retryAfter) res.setHeader("Retry-After", String(err.retryAfter));
        return send(res, err.status, err.toJSON());
      }
      console.error(err);
      send(res, 500, { error: { code: "INTERNAL", message: "Something went wrong on the server. Try again, and tell IT if it keeps happening." } });
    }
  });
}

function send(res, status, body) {
  if (body === undefined) {
    res.writeHead(status);
    return res.end();
  }
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(body));
}

async function readJson(req) {
  let size = 0;
  const chunks = [];
  for await (const chunk of req) {
    size += chunk.length;
    if (size > MAX_REQUEST_BYTES) throw new AppError(413, "EMAIL_TOO_LARGE", "This email is too large to send to the server.");
    chunks.push(chunk);
  }
  try {
    return JSON.parse(Buffer.concat(chunks).toString("utf8") || "{}");
  } catch {
    throw new AppError(400, "BAD_REQUEST", "The request wasn't valid JSON.");
  }
}

// Run directly (npm start), as opposed to imported by the tests.
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const config = loadConfig();
  const service = new JobService({
    config,
    fergus: new FergusClient({ apiToken: config.fergus.apiToken, baseUrl: config.fergus.baseUrl }),
    anthropic: new Anthropic({ apiKey: config.anthropic.apiKey }),
    store: new JobStore(config.dataDir),
  });
  const server = createServer({ config, service, authenticate: createAuthenticator(config.auth) });
  server.listen(config.port, () => {
    console.log(`Fergus job backend listening on :${config.port}`);
    if (config.auth.disabled) console.warn("AUTH_DISABLED=true — anyone who can reach this port can create Fergus jobs. Local testing only.");
    if (!config.fergus.groups.length) console.warn("FERGUS_GROUPS is empty — no group will be suggested.");
  });
}
