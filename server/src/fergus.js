/* Fergus API client — only the handful of calls this add-in needs.
 *
 * Endpoints and field names (https://api.fergus.com, Bearer personal access
 * token):
 *   GET  /customers?filterSearchText=…     POST /customers
 *   GET  /sites?filterSearchText=…         POST /sites
 *   POST /jobs                             POST /notes  (entityName "job")
 * Lists come back as { data: [...], paging: {...} } and single records as
 * { data: {...} }; unwrap() tolerates either that or a bare object, since
 * response envelopes are the least documented part of the API.
 */
import { AppError } from "./errors.js";

const MAX_AUTO_RETRY_WAIT_S = 5;

export class FergusClient {
  // The default wraps fetch rather than storing it: Workers reject a fetch
  // called as a method of another object ("Illegal invocation").
  constructor({ apiToken, baseUrl, fetchImpl = (...args) => fetch(...args), sleep = defaultSleep }) {
    this.apiToken = apiToken;
    this.baseUrl = baseUrl;
    this.fetch = fetchImpl;
    this.sleep = sleep;
  }

  // `action` reads as the end of "Fergus … while <action>", for messages
  // the task pane shows verbatim.
  async request(method, path, { body, action }) {
    for (let attempt = 0; ; attempt++) {
      let response;
      try {
        response = await this.fetch(`${this.baseUrl}${path}`, {
          method,
          headers: {
            Authorization: `Bearer ${this.apiToken}`,
            Accept: "application/json",
            ...(body ? { "Content-Type": "application/json" } : {}),
          },
          body: body ? JSON.stringify(body) : undefined,
        });
      } catch (err) {
        throw new AppError(502, "FERGUS_UNREACHABLE", `Couldn't reach Fergus while ${action}. Check your connection and try again.`, {
          details: String(err?.message || err),
        });
      }

      const parsed = await readBody(response);
      if (response.ok) return parsed;

      // A 429 is refused before Fergus does anything, so one short automatic
      // retry is safe even for a POST. Longer waits go back to the person.
      const retryAfter = parseRetryAfter(response.headers.get("retry-after"));
      if (response.status === 429 && attempt === 0 && (retryAfter ?? 1) <= MAX_AUTO_RETRY_WAIT_S) {
        await this.sleep((retryAfter ?? 1) * 1000);
        continue;
      }
      throw fergusError(response.status, parsed, action, retryAfter);
    }
  }

  // ----- Customers -----

  async searchCustomers(text) {
    const params = new URLSearchParams({ pageSize: "10", filterSearchText: text });
    const res = await this.request("GET", `/customers?${params}`, { action: "searching customers" });
    return unwrapList(res).map(normalizeCustomer);
  }

  async createCustomer({ name, contact, address }) {
    const body = { customerFullName: name, mainContact: toFergusContact(contact) };
    if (hasAddress(address)) body.physicalAddress = toFergusAddress(address);
    const res = await this.request("POST", "/customers", { body, action: "creating the customer" });
    return normalizeCustomer(unwrap(res));
  }

  // ----- Sites -----

  async searchSites(text, postcode) {
    const params = new URLSearchParams({ pageSize: "10", filterSearchText: text });
    if (postcode) params.set("filterAddressPostalCode", postcode);
    const res = await this.request("GET", `/sites?${params}`, { action: "searching sites" });
    return unwrapList(res).map(normalizeSite);
  }

  async createSite({ name, address, contact }) {
    const body = {
      siteAddress: toFergusAddress(address),
      defaultContact: toFergusContact(contact),
    };
    if (name) body.name = name;
    const res = await this.request("POST", "/sites", { body, action: "creating the site" });
    return normalizeSite(unwrap(res));
  }

  // ----- Jobs and notes -----

  async createJob({ jobType, title, description, customerId, siteId, customerReference }) {
    const body = { jobType, title, description, customerId, siteId, isDraft: false };
    if (customerReference) body.customerReference = customerReference;
    const res = await this.request("POST", "/jobs", { body, action: "creating the job" });
    const job = unwrap(res);
    return {
      id: job.id ?? job.jobId,
      jobNo: job.jobNo ?? job.jobNumber ?? job.internalJobId ?? null,
      raw: job,
    };
  }

  async addJobNote(jobId, text) {
    // GET /notes filters on "JOB", but POST /notes validates the lowercase "job".
    const body = { text, entityName: "job", entityId: Number(jobId), isPinned: false, parentId: null };
    const res = await this.request("POST", "/notes", { body, action: "adding the email as a job note" });
    const note = unwrap(res);
    return { id: note?.id ?? null };
  }
}

// ---------------------------------------------------------------------------
// Request/response plumbing

function defaultSleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function readBody(response) {
  const text = await response.text();
  if (!text) return {};
  try {
    return JSON.parse(text);
  } catch {
    return { message: text.slice(0, 500) };
  }
}

export function parseRetryAfter(header) {
  if (!header) return undefined;
  const seconds = Number(header);
  if (Number.isFinite(seconds)) return Math.max(0, Math.ceil(seconds));
  const date = Date.parse(header);
  if (Number.isFinite(date)) return Math.max(0, Math.ceil((date - Date.now()) / 1000));
  return undefined;
}

function fergusMessage(body) {
  if (!body || typeof body !== "object") return "";
  const parts = [];
  if (typeof body.message === "string") parts.push(body.message);
  if (typeof body.error === "string") parts.push(body.error);
  if (typeof body.error?.message === "string") parts.push(body.error.message);
  const errors = body.errors ?? body.error?.errors;
  if (Array.isArray(errors)) {
    for (const e of errors) parts.push(typeof e === "string" ? e : e?.message || JSON.stringify(e));
  } else if (errors && typeof errors === "object") {
    for (const [field, msg] of Object.entries(errors)) parts.push(`${field}: ${[].concat(msg).join(", ")}`);
  }
  return [...new Set(parts)].join(" — ");
}

export function fergusError(status, body, action, retryAfter) {
  const detail = fergusMessage(body);
  const suffix = detail ? ` Fergus said: ${detail}` : "";
  if (status === 401) {
    return new AppError(502, "FERGUS_AUTH", `Fergus rejected the API token while ${action}. An admin needs to check FERGUS_API_TOKEN on the server.`);
  }
  if (status === 403) {
    return new AppError(502, "FERGUS_FORBIDDEN", `The Fergus API token isn't allowed to do this (while ${action}).${suffix}`);
  }
  if (status === 404) {
    return new AppError(502, "FERGUS_NOT_FOUND", `Fergus couldn't find a record it needed while ${action}.${suffix}`);
  }
  if (status === 429) {
    const wait = retryAfter ? ` Try again in ${retryAfter} second${retryAfter === 1 ? "" : "s"}.` : " Wait a minute and try again.";
    return new AppError(429, "FERGUS_RATE_LIMITED", `Fergus is limiting how fast we can send requests (while ${action}).${wait}`, { retryAfter });
  }
  if (status >= 400 && status < 500) {
    return new AppError(422, "FERGUS_VALIDATION", `Fergus didn't accept the details while ${action}.${suffix || " Check the form and try again."}`, { details: body });
  }
  return new AppError(502, "FERGUS_UNAVAILABLE", `Fergus had a problem (HTTP ${status}) while ${action}. Try again shortly.${suffix}`);
}

function unwrap(res) {
  return res && typeof res === "object" && res.data && !Array.isArray(res.data) ? res.data : res;
}

function unwrapList(res) {
  if (Array.isArray(res)) return res;
  if (Array.isArray(res?.data)) return res.data;
  return [];
}

// ---------------------------------------------------------------------------
// Shape conversion: our flat form fields <-> Fergus records

export function splitName(fullName) {
  const parts = (fullName || "").trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return { firstName: "", lastName: "" };
  if (parts.length === 1) return { firstName: parts[0], lastName: "" };
  return { firstName: parts.slice(0, -1).join(" "), lastName: parts[parts.length - 1] };
}

// NZ mobile numbers start 02 (or +64 2); everything else is filed as a phone.
export function phoneType(phone) {
  const digits = (phone || "").replace(/[^\d+]/g, "");
  return /^(\+?642|02)/.test(digits) ? "mobile" : "phone";
}

export function toFergusContact({ name, company, email, phone } = {}) {
  const { firstName, lastName } = splitName(name || company);
  const contact = { firstName: firstName || "Unknown" };
  if (lastName) contact.lastName = lastName;
  if (company) contact.company = company;
  const items = [];
  if (email) items.push({ contactType: "email", contactValue: email });
  if (phone) items.push({ contactType: phoneType(phone), contactValue: phone });
  if (items.length) contact.contactItems = items;
  return contact;
}

function hasAddress(a) {
  return Boolean(a && (a.address1 || a.city || a.postcode));
}

export function toFergusAddress(a = {}) {
  const out = {};
  if (a.address1) out.address1 = a.address1;
  if (a.address2) out.address2 = a.address2;
  if (a.suburb) out.addressSuburb = a.suburb;
  if (a.city) out.addressCity = a.city;
  if (a.region) out.addressRegion = a.region;
  if (a.postcode) out.addressPostcode = a.postcode;
  if (a.country) out.addressCountry = a.country;
  return out;
}

function fromFergusAddress(a) {
  if (!a || typeof a !== "object") return null;
  const address = {
    address1: a.address1 ?? a.line1 ?? "",
    address2: a.address2 ?? a.line2 ?? "",
    suburb: a.addressSuburb ?? a.suburb ?? "",
    city: a.addressCity ?? a.city ?? "",
    region: a.addressRegion ?? a.region ?? "",
    postcode: a.addressPostcode ?? a.addressPostalCode ?? a.postalCode ?? a.postcode ?? "",
  };
  return Object.values(address).some(Boolean) ? address : null;
}

export function formatAddress(a) {
  if (!a) return "";
  return [a.address1, a.address2, a.suburb, a.city, a.postcode].filter(Boolean).join(", ");
}

function contactSummary(c) {
  if (!c) return { name: "", email: "", phone: "" };
  const items = c.contactItems || [];
  const find = (...types) => items.find((i) => types.includes(i.contactType))?.contactValue || "";
  const name = [c.firstName, c.lastName].filter(Boolean).join(" ") || c.name || "";
  return {
    name,
    email: find("email") || c.email || "",
    phone: find("mobile", "phone") || c.phone || c.mobile || "",
  };
}

function normalizeCustomer(c) {
  const contact = contactSummary(c.mainContact);
  const address = fromFergusAddress(c.physicalAddress);
  return {
    id: c.id ?? c.customerId,
    name: c.customerFullName ?? c.name ?? "",
    contactName: contact.name,
    email: contact.email,
    phone: contact.phone,
    address: formatAddress(address),
  };
}

function normalizeSite(s) {
  const address = fromFergusAddress(s.siteAddress ?? s.address);
  const contact = contactSummary(s.defaultContact);
  return {
    id: s.id ?? s.siteId,
    name: s.name ?? "",
    address: formatAddress(address),
    postcode: address?.postcode || "",
    contactName: contact.name,
  };
}
