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
  assert.match(req.messages[0].content.at(-1).text, /<email>[\s\S]*Aroha Ngata[\s\S]*<\/email>/);
});

test("create makes customer, site, job and note, then is idempotent", async () => {
  const { service, fergusFake } = await setup();

  const first = await service.create(newJobRequest(), { name: "Pat Office", email: "pat@example.com" });
  assert.equal(first.status, "created");
  assert.notEqual(first.job.jobNo, first.job.id);
  assert.equal(first.job.url, `https://app.fergus.com/jobs/view/${first.job.jobNo}/summary`);

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
  assert.match(note.text, /Attachments: deck-plan\.pdf/);

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

// ---------------------------------------------------------------------------
// Attachments

function attachmentRequest(overrides = {}) {
  return {
    internetMessageId: sampleEmail.internetMessageId,
    key: "deck-plan.pdf|120000",
    name: "deck-plan.pdf",
    contentType: "application/pdf",
    contentBase64: Buffer.from("%PDF-1.7 pretend plan").toString("base64"),
    ...overrides,
  };
}

test("an attachment is uploaded to the job once, as multipart", async () => {
  const { service, fergusFake } = await setup();
  const created = await service.create(newJobRequest(), {});
  assert.deepEqual(created.job.attachments, []);

  const first = await service.addAttachment(attachmentRequest());
  assert.equal(first.status, "created");
  assert.deepEqual(first.job.attachments, ["deck-plan.pdf|120000"]);

  const uploads = fergusFake.calls.filter((c) => c.path === "/attachments");
  assert.equal(uploads.length, 1);
  assert.equal(uploads[0].body.entityType, "job");
  assert.equal(uploads[0].body.entityId, String(created.job.id));
  assert.deepEqual(uploads[0].body.file, { name: "deck-plan.pdf", type: "application/pdf", text: "%PDF-1.7 pretend plan" });

  // Trying again, or reopening the pane, doesn't copy it a second time.
  const again = await service.addAttachment(attachmentRequest());
  assert.equal(again.status, "exists");
  assert.equal(fergusFake.calls.filter((c) => c.path === "/attachments").length, 1);

  // The job as the pane sees it on reopening lists what's been copied.
  const reopened = await service.analyze(sampleEmail);
  assert.deepEqual(reopened.job.attachments, ["deck-plan.pdf|120000"]);
});

test("an attachment needs the job to exist first", async () => {
  const { service, fergusFake } = await setup();
  await assert.rejects(service.addAttachment(attachmentRequest()), (err) => err.code === "NO_JOB" && err.status === 409);
  assert.equal(fergusFake.calls.length, 0);
});

test("a failed upload isn't recorded, so trying again copies it", async () => {
  const { service, fergusFake } = await setup({
    fergusData: { failures: { "POST /attachments": [{ status: 422, body: { message: "File type not allowed" } }] } },
  });
  await service.create(newJobRequest(), {});

  await assert.rejects(service.addAttachment(attachmentRequest()), (err) => err.code === "FERGUS_VALIDATION" && /File type not allowed/.test(err.message));
  const retry = await service.addAttachment(attachmentRequest());
  assert.equal(retry.status, "created");
  assert.equal(fergusFake.calls.filter((c) => c.path === "/attachments").length, 2);
});

test("oversized or malformed attachments are refused before Fergus", async () => {
  const { service, fergusFake } = await setup();
  await service.create(newJobRequest(), {});
  const before = fergusFake.calls.length;

  await assert.rejects(
    service.addAttachment(attachmentRequest({ contentBase64: "A".repeat(36 * 1024 * 1024) })),
    (err) => err.code === "ATTACHMENT_TOO_LARGE" && err.status === 413
  );
  await assert.rejects(service.addAttachment(attachmentRequest({ contentBase64: "" })), (err) => err.code === "BAD_REQUEST");
  await assert.rejects(service.addAttachment(attachmentRequest({ contentBase64: "not base64!" })), (err) => err.code === "BAD_REQUEST");
  await assert.rejects(service.addAttachment(attachmentRequest({ key: "" })), (err) => err.code === "BAD_REQUEST");
  assert.equal(fergusFake.calls.length, before);
});

test("HTTP: attachments route takes a body bigger than an email's limit", async () => {
  const { config, service } = await setup();
  await service.create(newJobRequest(), {});
  const handle = createHandler(() => ({ config, service, authenticate: async () => ({ name: "Pat" }) }));

  const bytes = Buffer.alloc(3 * 1024 * 1024, 7);
  const res = await handle(
    new Request("https://worker.test/api/jobs/attachments", {
      method: "POST",
      body: JSON.stringify(attachmentRequest({ contentBase64: bytes.toString("base64") })),
    }),
    {}
  );
  assert.equal(res.status, 200);
  assert.equal((await res.json()).status, "created");
});

// ---------------------------------------------------------------------------
// Images read by Claude

const png = Buffer.from("fake png bytes").toString("base64");

test("images go to Claude numbered, ahead of the email; unusable ones are dropped", async () => {
  const { service, anthropic } = await setup();
  await service.analyze(sampleEmail, [
    { name: "Screenshot_Messages.jpg", contentType: "image/jpg", contentBase64: png },
    { name: "image001.png", contentType: "image/png", contentBase64: png, inline: true },
    { name: "photo.heic", contentType: "image/heic", contentBase64: png },
    { name: "bad.png", contentType: "image/png", contentBase64: "not base64!" },
    { name: "empty.png", contentType: "image/png", contentBase64: "" },
  ]);

  const content = anthropic.requests[0].messages[0].content;
  assert.deepEqual(content.map((b) => b.type), ["text", "image", "text", "image", "text"]);
  assert.equal(content[0].text, 'Image 1: "Screenshot_Messages.jpg", attached to the email.');
  assert.deepEqual(content[1].source, { type: "base64", media_type: "image/jpeg", data: png });
  assert.equal(content[2].text, 'Image 2: "image001.png", pasted into the email body.');
  assert.match(content[4].text, /^<email>/);
});

test("the pane is told which images Claude read and why others were left out", async () => {
  const { service, anthropic } = await setup();
  const result = await service.analyze(sampleEmail, [
    { name: "shot.jpg", contentType: "image/jpeg", contentBase64: `data:image/jpeg;base64,${png}` },
    { name: "urlsafe.png", contentType: "image/png", contentBase64: "-_8" },
    { name: "photo.heic", contentType: "image/heic", contentBase64: png },
    { name: "bad.png", contentType: "image/png", contentBase64: "not base64!" },
  ]);
  assert.deepEqual(result.images.read, ["shot.jpg", "urlsafe.png"]);
  assert.deepEqual(result.images.skipped.map((s) => s.name), ["photo.heic", "bad.png"]);
  assert.match(result.images.skipped[0].reason, /image\/heic/);
  const sources = anthropic.requests[0].messages[0].content.filter((b) => b.type === "image").map((b) => b.source.data);
  assert.deepEqual(sources, [png, "+/8="]);
});

test("no images means a single text block, as before", async () => {
  const { service, anthropic } = await setup();
  await service.analyze(sampleEmail);
  assert.deepEqual(anthropic.requests[0].messages[0].content.map((b) => b.type), ["text"]);
});

test("staff domains and text-message screenshots are explained to Claude", async () => {
  const { service, anthropic } = await setup();
  await service.analyze(sampleEmail);
  const config = loadConfig({ ...env, OWN_EMAIL_DOMAINS: "citytimber.co.nz,falconflooring.co.nz" });
  assert.deepEqual(config.ownEmailDomains, ["citytimber.co.nz", "falconflooring.co.nz"]);
  assert.match(anthropic.requests[0].system, /screenshot of a text-message conversation/);
});

test("HTTP: analyze accepts images bigger than an email-only request", async () => {
  const { config, service, anthropic } = await setup();
  const handle = createHandler(() => ({ config, service, authenticate: async () => ({ name: "Pat" }) }));
  const big = Buffer.alloc(3 * 1024 * 1024, 1).toString("base64");
  const res = await handle(
    new Request("https://worker.test/api/analyze", {
      method: "POST",
      body: JSON.stringify({ email: sampleEmail, images: [{ name: "shot.png", contentType: "image/png", contentBase64: big }] }),
    }),
    {}
  );
  assert.equal(res.status, 200);
  assert.equal(anthropic.requests[0].messages[0].content[1].type, "image");
});

// ---------------------------------------------------------------------------
// Another job from the same email

test("an email that already made a job can make another, alongside it", async () => {
  const { service, fergusFake } = await setup();
  const first = (await service.create(newJobRequest(), { name: "Pat" })).job;
  await service.addAttachment(attachmentRequest());

  // Reopened: the existing job, unless another is asked for.
  assert.equal((await service.analyze(sampleEmail)).status, "exists");
  const again = await service.analyze(sampleEmail, [], { again: true });
  assert.equal(again.status, "new");
  assert.equal(again.currentJob.id, first.id);
  assert.equal(again.partial, null);

  const second = await service.create(newJobRequest({ alongsideJobId: first.id }), { name: "Pat" });
  assert.equal(second.status, "created");
  assert.notEqual(second.job.id, first.id);
  assert.deepEqual(second.job.attachments, []);
  assert.deepEqual(second.job.previousJobs.map((p) => [p.jobId, p.jobNo, p.url]), [[first.id, first.jobNo, first.url]]);

  const posts = fergusFake.calls.filter((c) => c.method === "POST").map((c) => c.path);
  assert.deepEqual(posts, ["/customers", "/sites", "/jobs", "/notes", "/attachments", "/customers", "/sites", "/jobs", "/notes"]);

  // The same attachment can now be copied to the new job too.
  assert.equal((await service.addAttachment(attachmentRequest())).status, "created");
  assert.equal((await service.analyze(sampleEmail)).job.id, second.job.id);
});

test("a repeated 'make another' click doesn't make a third job", async () => {
  const { service, fergusFake } = await setup();
  const first = (await service.create(newJobRequest(), {})).job;
  const second = await service.create(newJobRequest({ alongsideJobId: first.id }), {});
  const repeat = await service.create(newJobRequest({ alongsideJobId: first.id }), {});
  assert.equal(repeat.status, "exists");
  assert.equal(repeat.job.id, second.job.id);
  assert.equal(fergusFake.calls.filter((c) => c.path === "/jobs").length, 2);
});

test("making another job reuses nothing from the first, even existing records picked", async () => {
  const { service, fergusFake } = await setup();
  const first = (await service.create(newJobRequest({ customer: { id: "5" }, site: { id: "9" } }), {})).job;
  await service.create(newJobRequest({ alongsideJobId: first.id, customer: { id: "6" }, site: { id: "9" } }), {});
  const jobs = fergusFake.calls.filter((c) => c.path === "/jobs").map((c) => [c.body.customerId, c.body.siteId]);
  assert.deepEqual(jobs, [[5, 9], [6, 9]]);
});

// ---------------------------------------------------------------------------
// Placeholder site when the email gives no address

test("'site unknown' makes a placeholder site of its own, named after the job", async () => {
  const { service, fergusFake } = await setup();
  const result = await service.create(
    newJobRequest({ site: { unknown: true, contactName: "Mere Tane", phone: "027 555 1234", email: "" } }),
    { name: "Pat" }
  );
  assert.equal(result.status, "created");
  assert.equal(result.job.siteCreated, true);
  assert.equal(result.job.siteUnknown, true);
  assert.equal(result.job.siteAddress, `Unknown – Job ${result.job.jobNo}`);
  assert.equal(result.job.siteRenameError, null);

  // Created under a one-off reference, so Fergus can't match an existing site...
  const created = fergusFake.calls.find((c) => c.method === "POST" && c.path === "/sites");
  assert.match(created.body.siteAddress.address1, /^Unknown \(ref [0-9a-f]{8}\)$/);
  assert.equal(created.body.defaultContact.firstName, "Mere");
  assert.deepEqual(created.body.defaultContact.contactItems, [{ contactType: "mobile", contactValue: "027 555 1234" }]);

  // ...then renamed to the job number, contact kept, before the note.
  const order = fergusFake.calls.filter((c) => c.method !== "GET").map((c) => `${c.method} ${c.path.replace(/\d+$/, "{id}")}`);
  assert.deepEqual(order, ["POST /customers", "POST /sites", "POST /jobs", "PUT /sites/{id}", "POST /notes"]);
  const renamed = fergusFake.calls.find((c) => c.method === "PUT");
  assert.equal(renamed.path, "/sites/1001"); // the fake's ids: customer 1000, then this site
  assert.deepEqual(renamed.body.siteAddress, { address1: `Unknown – Job ${result.job.jobNo}` });
  assert.equal(renamed.body.defaultContact.firstName, "Mere");

  const note = fergusFake.calls.find((c) => c.path === "/notes").body;
  assert.match(note.text, new RegExp(`placeholder site "Unknown – Job ${result.job.jobNo}"`));
});

test("two jobs with an unknown site get two different placeholder sites", async () => {
  const { service, fergusFake } = await setup();
  await service.create(newJobRequest({ site: { unknown: true } }), {});
  await service.create(newJobRequest({ email: { ...sampleEmail, internetMessageId: "<other@x>" }, site: { unknown: true } }), {});
  const refs = fergusFake.calls.filter((c) => c.method === "POST" && c.path === "/sites").map((c) => c.body.siteAddress.address1);
  assert.equal(refs.length, 2);
  assert.notEqual(refs[0], refs[1]);
});

test("if Fergus won't rename the placeholder site, the job still goes ahead", async () => {
  const { service, fergusFake } = await setup({
    fergusData: { failures: { "PUT /sites/1001": [{ status: 404, body: { message: "Not found" } }] } },
  });
  const result = await service.create(newJobRequest({ site: { unknown: true } }), {});
  assert.equal(result.status, "created");
  assert.match(result.job.siteAddress, /^Unknown \(ref /);
  assert.match(result.job.siteRenameError, /renaming the placeholder site/);
  assert.equal(fergusFake.calls.filter((c) => c.path === "/notes").length, 1);
});

test("'site unknown' with no contact still gives Fergus a contact name", async () => {
  const { service, fergusFake } = await setup();
  await service.create(newJobRequest({ site: { unknown: true } }), {});
  const site = fergusFake.calls.find((c) => c.path === "/sites").body;
  assert.equal(site.defaultContact.firstName, "Unknown");
});

test("an ordinary new site still needs a street address", async () => {
  const { service } = await setup();
  await assert.rejects(service.create(newJobRequest({ site: { new: { address: {} } } }), {}), /street address/);
});
