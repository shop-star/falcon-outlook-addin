import assert from "node:assert/strict";
import { test } from "node:test";
import { loadConfig } from "../src/config.js";
import { extractionSchema } from "../src/extract.js";
import { FergusClient, fergusError, parseRetryAfter, toFergusContact } from "../src/fergus.js";
import { JobService } from "../src/jobs.js";
import { JobStore } from "../src/store.js";
import { createHandler } from "../src/worker.js";
import { testD1 } from "./d1.js";
import { fakeAnthropic, fakeFergus, sampleEmail, sampleExtraction } from "./fakes.js";

const env = {
  FERGUS_API_TOKEN: "test-token",
  ANTHROPIC_API_KEY: "test-key",
  AUTH_DISABLED: "true",
  FERGUS_GROUPS: "Decking,Framing,Joinery",
};

async function setup({ fergusData, extraction = sampleExtraction } = {}) {
  const config = loadConfig(env);
  const fergusFake = fakeFergus(fergusData);
  const anthropic = fakeAnthropic(extraction);
  const fergus = new FergusClient({ apiToken: "test-token", baseUrl: "https://fergus.test", fetchImpl: fergusFake.fetchImpl, sleep: async () => {} });
  const store = new JobStore(testD1());
  const service = new JobService({ config, fergus, anthropic, store });
  return { config, service, store, fergusFake, anthropic };
}

function newJobRequest(overrides = {}) {
  return {
    email: sampleEmail,
    jobType: "Quote",
    group: "Decking",
    title: "Decking timber – 14 Kowhai Rd",
    description: "Supply 40 m² of 140x32 kwila decking with fixings.",
    customer: { new: { name: "Aroha Ngata", contactName: "Aroha Ngata", email: "aroha.ngata@example.com", phone: "021 555 0142", address: {} } },
    site: { new: { address: { address1: "14 Kowhai Road", suburb: "Mt Eden", city: "Auckland", postcode: "1024" }, contactName: "Aroha Ngata", phone: "021 555 0142" } },
    ...overrides,
  };
}

test("analyze returns Claude's extraction and ranks Fergus matches", async () => {
  const { service, anthropic, fergusFake } = await setup({
    fergusData: {
      customers: [
        { id: 1, customerFullName: "Aroha Ngata", mainContact: { firstName: "Aroha", lastName: "Ngata", contactItems: [{ contactType: "email", contactValue: "aroha.ngata@example.com" }, { contactType: "mobile", contactValue: "021 555 0142" }] } },
        { id: 2, customerFullName: "Aroha Ngata Trust", mainContact: { firstName: "Bob" } },
      ],
      sites: [{ id: 7, name: "", siteAddress: { address1: "14 Kowhai Road", addressCity: "Auckland", addressPostcode: "1024" } }],
    },
  });

  const result = await service.analyze(sampleEmail);

  assert.equal(result.status, "new");
  assert.equal(result.extraction.customerName.value, "Aroha Ngata");
  assert.deepEqual(result.candidates.customers.map((c) => c.id), [1, 2]);
  assert.deepEqual(result.candidates.customers[0].matchedOn.sort(), ["email", "name", "phone"]);
  assert.equal(result.candidates.sites[0].id, 7);
  assert.equal(result.candidates.sites[0].address, "14 Kowhai Road, Auckland, 1024");
  assert.deepEqual(result.options.groups, ["Decking", "Framing", "Joinery"]);
  assert.ok(fergusFake.calls.every((c) => c.auth === "Bearer test-token"));

  const req = anthropic.requests[0];
  assert.equal(req.model, "claude-opus-5-5");
  assert.equal(req.output_config.format.type, "json_schema");
  assert.match(req.messages[0].content, /<email>[\s\S]*Aroha Ngata[\s\S]*<\/email>/);
});

test("create makes customer, site, job and note, then is idempotent", async () => {
  const { service, fergusFake } = await setup();

  const first = await service.create(newJobRequest(), { name: "Pat Office", email: "pat@example.com" });
  assert.equal(first.status, "created");
  assert.ok(first.job.jobNo.startsWith("J"));
  assert.match(first.job.url, /^https:\/\/app\.fergus\.com\/jobs\/view\/\d+$/);

  const posts = fergusFake.calls.filter((c) => c.method === "POST");
  assert.deepEqual(posts.map((c) => c.path), ["/customers", "/sites", "/jobs", "/notes"]);

  const [customer, site, job, note] = posts.map((c) => c.body);
  assert.equal(customer.customerFullName, "Aroha Ngata");
  assert.deepEqual(customer.mainContact.contactItems, [
    { contactType: "email", contactValue: "aroha.ngata@example.com" },
    { contactType: "mobile", contactValue: "021 555 0142" },
  ]);
  assert.equal(site.siteAddress.addressPostcode, "1024");
  assert.equal(job.jobType, "Quote");
  assert.equal(job.isDraft, false);
  assert.equal(note.entityName, "job");
  assert.equal(note.entityId, first.job.id);
  assert.match(note.text, /Group: Decking/);
  assert.match(note.text, /From: Sam Builder <sam@builders\.example\.com>/);
  assert.match(note.text, /Attachments \(in Outlook\): deck-plan\.pdf/);

  const again = await service.create(newJobRequest(), {});
  assert.equal(again.status, "exists");
  assert.equal(again.job.id, first.job.id);
  assert.equal(fergusFake.calls.filter((c) => c.method === "POST").length, 4);

  const analyzed = await service.analyze(sampleEmail);
  assert.equal(analyzed.status, "exists");
});

test("existing customer and site are used as-is", async () => {
  const { service, fergusFake } = await setup();
  await service.create(newJobRequest({ customer: { id: "41" }, site: { id: 52 } }), {});
  const posts = fergusFake.calls.filter((c) => c.method === "POST");
  assert.deepEqual(posts.map((c) => c.path), ["/jobs", "/notes"]);
  assert.equal(posts[0].body.customerId, 41);
  assert.equal(posts[0].body.siteId, 52);
});

test("a retry after a failure reuses what was already created", async () => {
  const { service, fergusFake } = await setup({
    fergusData: { failures: { "POST /notes": [{ status: 429, headers: { "retry-after": "30" } }] } },
  });

  await assert.rejects(service.create(newJobRequest(), {}), (err) => {
    assert.equal(err.code, "FERGUS_RATE_LIMITED");
    assert.equal(err.status, 429);
    assert.equal(err.retryAfter, 30);
    assert.match(err.message, /Try again in 30 seconds/);
    return true;
  });

  const retry = await service.create(newJobRequest(), {});
  assert.equal(retry.status, "created");
  const posts = fergusFake.calls.filter((c) => c.method === "POST").map((c) => c.path);
  assert.deepEqual(posts, ["/customers", "/sites", "/jobs", "/notes", "/notes"]);
});

test("a short 429 is retried automatically", async () => {
  const { service, fergusFake } = await setup({
    fergusData: { failures: { "POST /jobs": [{ status: 429, headers: { "retry-after": "1" } }] } },
  });
  const result = await service.create(newJobRequest({ customer: { id: 1 }, site: { id: 2 } }), {});
  assert.equal(result.status, "created");
  assert.equal(fergusFake.calls.filter((c) => c.path === "/jobs").length, 2);
});

test("default fetch is called unbound, as Cloudflare Workers require", async () => {
  const realFetch = globalThis.fetch;
  globalThis.fetch = function () {
    // Workers throw "Illegal invocation" when fetch is called as a method.
    if (this !== undefined && this !== globalThis) throw new TypeError("Illegal invocation");
    return Promise.resolve(new Response(JSON.stringify({ data: [] }), { status: 200 }));
  };
  try {
    const fergus = new FergusClient({ apiToken: "t", baseUrl: "https://fergus.test" });
    assert.deepEqual(await fergus.searchSites("x"), []);
  } finally {
    globalThis.fetch = realFetch;
  }
});

test("Fergus errors become clear messages", () => {
  assert.equal(fergusError(401, {}, "creating the job").code, "FERGUS_AUTH");
  const v = fergusError(422, { errors: { siteId: ["is required"] } }, "creating the job");
  assert.equal(v.code, "FERGUS_VALIDATION");
  assert.match(v.message, /siteId: is required/);
  assert.equal(fergusError(503, {}, "searching sites").code, "FERGUS_UNAVAILABLE");
  assert.equal(parseRetryAfter("12"), 12);
  assert.equal(parseRetryAfter(undefined), undefined);
});

test("contacts: name split and NZ mobile detection", () => {
  assert.deepEqual(toFergusContact({ name: "Mary Jane Smith", phone: "09 555 1234" }), {
    firstName: "Mary Jane",
    lastName: "Smith",
    contactItems: [{ contactType: "phone", contactValue: "09 555 1234" }],
  });
  assert.equal(toFergusContact({ company: "Acme Ltd" }).firstName, "Acme");
  assert.equal(toFergusContact({ phone: "+64 27 123 4567" }).contactItems[0].contactType, "mobile");
});

test("create rejects bad input before touching Fergus", async () => {
  const { service, fergusFake } = await setup();
  await assert.rejects(service.create(newJobRequest({ jobType: "Banana" }), {}), /Job type must be one of/);
  await assert.rejects(service.create(newJobRequest({ group: "Plumbing" }), {}), /Pick a group/);
  await assert.rejects(service.create(newJobRequest({ customer: { new: {} } }), {}), /new customer's name/);
  await assert.rejects(service.create(newJobRequest({ site: { id: "1; DROP" } }), {}), /record ID/);
  await assert.rejects(service.analyze({ ...sampleEmail, internetMessageId: "" }), /Internet message ID/);
  assert.equal(fergusFake.calls.length, 0);
});

test("extraction schema follows structured-output rules", () => {
  const check = (schema, where) => {
    if (schema.type !== "object") return;
    assert.equal(schema.additionalProperties, false, `${where} must set additionalProperties: false`);
    assert.deepEqual([...schema.required].sort(), Object.keys(schema.properties).sort(), `${where} must require every property`);
    for (const [k, v] of Object.entries(schema.properties)) check(v, `${where}.${k}`);
  };
  check(extractionSchema(["Decking"]), "root");
  check(extractionSchema([]), "root");
  assert.deepEqual(extractionSchema(["Decking"]).properties.group.properties.value.enum, ["Decking", ""]);
});

test("two creates for the same email can't run at once", async () => {
  const { service, store, fergusFake } = await setup();
  // Another Worker instance is mid-way through creating this job.
  assert.equal(await store.acquire(sampleEmail.internetMessageId), true);

  await assert.rejects(service.create(newJobRequest(), {}), (err) => err.code === "IN_PROGRESS" && err.status === 409);
  assert.equal(fergusFake.calls.length, 0);

  await store.release(sampleEmail.internetMessageId);
  assert.equal((await service.create(newJobRequest(), {})).status, "created");
});

test("a lock left by a crashed instance expires", async () => {
  let now = 1_000_000;
  const store = new JobStore(testD1(), () => now);
  assert.equal(await store.acquire("<m@x>"), true);
  assert.equal(await store.acquire("<m@x>"), false);
  now += 3 * 60 * 1000;
  assert.equal(await store.acquire("<m@x>"), true);
  // The lock's placeholder row doesn't read as a record.
  assert.equal(await store.get("<m@x>"), null);
});

test("HTTP: CORS, auth and error bodies", async () => {
  const { config, service } = await setup();
  const { AppError } = await import("../src/errors.js");
  const authenticate = async (request) => {
    if (request.headers.get("Authorization") !== "Bearer good") throw new AppError(401, "UNAUTHENTICATED", "Sign in.");
    return { name: "Pat" };
  };
  const handle = createHandler(() => ({ config, service, authenticate }));
  const call = (path, init = {}) => handle(new Request(`https://worker.test${path}`, init), {});

  const pre = await call("/api/analyze", { method: "OPTIONS", headers: { Origin: "https://shop-star.github.io" } });
  assert.equal(pre.status, 204);
  assert.equal(pre.headers.get("access-control-allow-origin"), "https://shop-star.github.io");

  const evil = await call("/api/analyze", { method: "OPTIONS", headers: { Origin: "https://evil.example" } });
  assert.equal(evil.headers.get("access-control-allow-origin"), null);

  assert.equal((await call("/api/health")).status, 200);

  const noAuth = await call("/api/analyze", { method: "POST", body: "{}" });
  assert.equal(noAuth.status, 401);
  assert.equal((await noAuth.json()).error.code, "UNAUTHENTICATED");

  const ok = await call("/api/analyze", {
    method: "POST",
    headers: { Authorization: "Bearer good", "Content-Type": "application/json" },
    body: JSON.stringify({ email: sampleEmail }),
  });
  assert.equal(ok.status, 200);
  assert.equal((await ok.json()).status, "new");

  const bad = await call("/api/jobs", {
    method: "POST",
    headers: { Authorization: "Bearer good" },
    body: JSON.stringify({ email: sampleEmail, jobType: "Nope" }),
  });
  assert.equal(bad.status, 400);
  assert.equal((await bad.json()).error.code, "BAD_REQUEST");
});

test("HTTP: missing settings give a clear error", async () => {
  const handle = createHandler((env) => loadConfig(env));
  const res = await handle(new Request("https://worker.test/api/health"), {});
  assert.equal(res.status, 500);
  assert.equal((await res.json()).error.code, "MISCONFIGURED");
});
