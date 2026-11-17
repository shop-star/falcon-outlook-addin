/* The things the task pane asks for:
 *   analyze(email)          → prefilled form: Claude's extraction + Fergus match candidates
 *   create(request)         → customer/site (if new) → job → email as a job note
 *   addAttachment(request)  → one of the email's attachments copied onto that job
 */
import { JOB_TYPES } from "./config.js";
import { AppError, badRequest } from "./errors.js";
import { extractJobDetails } from "./extract.js";
import { findCandidates } from "./match.js";

const MAX_BODY_CHARS = 200_000;
// Outlook itself won't attach much more than this, and it keeps a request
// (base64 adds a third) well inside what a Worker accepts.
export const MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024;
// Images sent to Claude with the email. Claude takes up to 10 MB of base64
// per image; the total keeps the request a size a Worker handles easily.
const CLAUDE_IMAGE_TYPES = ["image/jpeg", "image/png", "image/gif", "image/webp"];
export const MAX_IMAGES = 8;
const MAX_IMAGE_BASE64 = 10 * 1024 * 1024;
export const MAX_IMAGES_BASE64_TOTAL = 20 * 1024 * 1024;

export class JobService {
  constructor({ config, fergus, anthropic, store }) {
    this.config = config;
    this.fergus = fergus;
    this.anthropic = anthropic;
    this.store = store;
  }

  options() {
    const { groups, defaultJobType } = this.config.fergus;
    return { jobTypes: JOB_TYPES, defaultJobType, groups };
  }

  async analyze(rawEmail, rawImages) {
    const email = validateEmail(rawEmail);
    const { images, skipped } = validateImages(rawImages);

    const existing = await this.store.get(email.internetMessageId);
    if (existing?.status === "done") {
      return { status: "exists", job: this.publicJob(existing) };
    }

    let extraction = null;
    let extractionError = null;
    try {
      extraction = await extractJobDetails(
        {
          client: this.anthropic,
          model: this.config.anthropic.model,
          groups: this.config.fergus.groups,
          ownDomains: this.config.ownEmailDomains,
        },
        email,
        images
      );
    } catch (err) {
      if (!(err instanceof AppError)) throw err;
      extractionError = err.toJSON().error;
    }

    let candidates = { customers: [], sites: [] };
    let searchError = null;
    if (extraction) {
      const found = await findCandidates(this.fergus, extraction);
      candidates = { customers: found.customers, sites: found.sites };
      const first = found.errors[0];
      if (first) searchError = first instanceof AppError ? first.toJSON().error : { code: "SEARCH_FAILED", message: "Searching Fergus failed." };
    }

    return {
      status: "new",
      extraction,
      extractionError,
      candidates,
      searchError,
      options: this.options(),
      // Which images Claude was given, and which were left out and why.
      images: { read: images.map((img) => img.name), skipped },
      // A previous attempt that stopped part-way: the pane says what will be reused.
      partial: existing ? { customerId: existing.customerId ?? null, siteId: existing.siteId ?? null, jobId: existing.jobId ?? null } : null,
    };
  }

  async create(raw, user) {
    const req = validateCreate(raw, this.options());
    const id = req.email.internetMessageId;

    if (!(await this.store.acquire(id))) {
      throw new AppError(409, "IN_PROGRESS", "A Fergus job is already being created for this email. Wait a moment, then reopen the pane.");
    }
    try {
      return await this.createSteps(req, user);
    } finally {
      await this.store.release(id);
    }
  }

  async createSteps(req, user) {
    const id = req.email.internetMessageId;
    let record = await this.store.get(id);
    if (record?.status === "done") {
      return { status: "exists", job: this.publicJob(record) };
    }
    const save = async (changes) => {
      record = await this.store.update(id, { status: "in_progress", subject: req.email.subject, ...changes });
    };

    // Reuse anything an earlier, interrupted attempt already created.
    let customerId = record?.customerId ?? req.customer.id;
    if (customerId == null) {
      const c = req.customer.new;
      const customer = await this.fergus.createCustomer({
        name: c.name,
        contact: { name: c.contactName, company: c.company, email: c.email, phone: c.phone },
        address: c.address,
      });
      customerId = customer.id;
      await save({ customerId, customerCreated: true });
    }

    let siteId = record?.siteId ?? req.site.id;
    if (siteId == null) {
      const s = req.site.new;
      const site = await this.fergus.createSite({
        name: s.name,
        address: s.address,
        contact: { name: s.contactName, email: s.email, phone: s.phone },
      });
      siteId = site.id;
      await save({ siteId, siteCreated: true });
    }

    if (record?.jobId == null) {
      const job = await this.fergus.createJob({
        jobType: req.jobType,
        title: req.title,
        description: req.description,
        customerId: Number(customerId),
        siteId: Number(siteId),
      });
      await save({ customerId, siteId, jobId: job.id, jobNo: job.jobNo });
    }

    if (!record.noteAdded) {
      await this.fergus.addJobNote(record.jobId, noteText(req, user));
      await save({ noteAdded: true });
    }

    record = await this.store.update(id, { status: "done", group: req.group, createdBy: user?.email || user?.name || "" });
    return { status: "created", job: this.publicJob(record) };
  }

  // Attachments go one per request, after the job exists, so a big email
  // never has to fit in one request and one bad file doesn't stop the rest.
  // Each copied attachment is recorded, so trying again (or a second person
  // opening the email) never copies the same file twice.
  async addAttachment(raw) {
    const req = validateAttachment(raw);
    const id = req.internetMessageId;

    let record = await this.store.get(id);
    if (record?.status !== "done" || record.jobId == null) {
      throw new AppError(409, "NO_JOB", "This email hasn't been turned into a Fergus job yet, so there's nothing to copy the attachment to.");
    }
    if (record.attachments?.[req.key]) return { status: "exists", job: this.publicJob(record) };

    if (!(await this.store.acquire(id))) {
      throw new AppError(409, "IN_PROGRESS", "Something else is being added to this job right now. Wait a moment and try again.");
    }
    try {
      record = await this.store.get(id);
      if (record.attachments?.[req.key]) return { status: "exists", job: this.publicJob(record) };
      const uploaded = await this.fergus.uploadJobAttachment(record.jobId, req);
      record = await this.store.update(id, {
        attachments: { ...record.attachments, [req.key]: { name: req.name, fergusId: uploaded.id } },
      });
      return { status: "created", job: this.publicJob(record) };
    } finally {
      await this.store.release(id);
    }
  }

  // The link is built each time rather than stored, so a change to the
  // template also fixes jobs created before it.
  publicJob(record) {
    const url = this.config.fergus.jobUrlTemplate
      .replace("{jobId}", encodeURIComponent(record.jobId ?? ""))
      // Fergus's web app addresses jobs by number. Should a job have come
      // back without one, its internal ID is the best guess left.
      .replace("{jobNo}", encodeURIComponent(record.jobNo ?? record.jobId ?? ""));
    return publicJobFields(record, url);
  }
}

function publicJobFields(record, url) {
  return {
    id: record.jobId,
    jobNo: record.jobNo ?? null,
    url,
    createdAt: record.updatedAt,
    createdBy: record.createdBy || "",
    customerCreated: Boolean(record.customerCreated),
    siteCreated: Boolean(record.siteCreated),
    // Keys (see attachmentKey in the task pane) of attachments already copied.
    attachments: Object.keys(record.attachments || {}),
  };
}

// Fergus's API has no way to set a job's group, so the note carries it
// where whoever picks the job up will see it.
export function noteText(req, user) {
  const e = req.email;
  const from = e.from.name ? `${e.from.name} <${e.from.address}>` : e.from.address;
  const attachments = e.attachments.map((a) => a.name).filter(Boolean);
  const lines = [
    `Created from email${user?.name ? ` by ${user.name}` : ""} (Outlook add-in)`,
  ];
  if (req.group) lines.push(`Group: ${req.group}`);
  lines.push(
    "",
    `From: ${from}`,
    `Date: ${e.receivedAt}`,
    `Subject: ${e.subject}`
  );
  if (attachments.length) lines.push(`Attachments: ${attachments.join(", ")}`);
  lines.push("", e.bodyText.trim());
  return lines.join("\n");
}

// ---------------------------------------------------------------------------
// Input validation. Everything from the task pane is untrusted.

function str(value, max = 1000) {
  return typeof value === "string" ? value.trim().slice(0, max) : "";
}

function validateEmail(e) {
  if (!e || typeof e !== "object") throw badRequest("Missing email details.");
  const internetMessageId = str(e.internetMessageId, 1000);
  if (!internetMessageId) throw badRequest("This email has no Internet message ID, so it can't be tracked. Try again once it has finished syncing.");
  if (typeof e.bodyText === "string" && e.bodyText.length > MAX_BODY_CHARS) {
    throw new AppError(413, "EMAIL_TOO_LARGE", `This email is too long to process (over ${MAX_BODY_CHARS.toLocaleString()} characters).`);
  }
  return {
    internetMessageId,
    subject: str(e.subject, 1000),
    from: { name: str(e.from?.name, 300), address: str(e.from?.address, 300) },
    receivedAt: str(e.receivedAt, 100),
    bodyText: typeof e.bodyText === "string" ? e.bodyText : "",
    attachments: Array.isArray(e.attachments)
      ? e.attachments.slice(0, 50).map((a) => ({ name: str(a?.name, 300), size: Number(a?.size) || 0, contentType: str(a?.contentType, 200) }))
      : [],
  };
}

// Images that can't be used are left out rather than refused, since the
// email can still be read without them; `skipped` says which and why, so
// the pane can tell the person.
export function validateImages(raw) {
  const images = [];
  const skipped = [];
  if (!Array.isArray(raw)) return { images, skipped };
  let total = 0;
  raw.forEach((img, i) => {
    const name = str(img?.name, 300) || "image";
    const skip = (reason) => skipped.push({ name, reason });
    if (i >= MAX_IMAGES) return skip(`only the first ${MAX_IMAGES} images are read`);
    const mediaType = str(img?.contentType, 100).toLowerCase().replace("image/jpg", "image/jpeg");
    if (!CLAUDE_IMAGE_TYPES.includes(mediaType)) return skip(`Claude can't read ${mediaType || "this type of"} images`);
    const data = cleanBase64(img?.contentBase64);
    if (!data) return skip("it came through empty or garbled");
    if (data.length > MAX_IMAGE_BASE64) return skip("it's too large");
    if (total + data.length > MAX_IMAGES_BASE64_TOTAL) return skip("the email's images are too large in total");
    total += data.length;
    images.push({ name, inline: Boolean(img.inline), mediaType, data });
  });
  return { images, skipped };
}

// Outlook hands back plain base64, but some versions wrap it in a data: URL
// or use the URL-safe alphabet; accept those too.
function cleanBase64(value) {
  if (typeof value !== "string") return "";
  const data = value
    .replace(/^data:[^,]*,/, "")
    .replace(/\s/g, "")
    .replace(/-/g, "+")
    .replace(/_/g, "/")
    .replace(/=+$/, "");
  if (!data || !/^[A-Za-z0-9+/]+$/.test(data) || data.length % 4 === 1) return "";
  return data + "=".repeat((4 - (data.length % 4)) % 4);
}

function validateAttachment(raw) {
  if (!raw || typeof raw !== "object") throw badRequest("Missing attachment details.");
  const internetMessageId = str(raw.internetMessageId, 1000);
  const key = str(raw.key, 400);
  const name = str(raw.name, 300).replace(/[\\/\u0000-\u001f]/g, "_");
  if (!internetMessageId || !key || !name) throw badRequest("Missing attachment details.");
  if (typeof raw.contentBase64 !== "string" || !raw.contentBase64) throw badRequest(`"${name}" came through empty.`);
  // Checked before decoding: 4 base64 characters are 3 bytes.
  if (raw.contentBase64.length > Math.ceil(MAX_ATTACHMENT_BYTES / 3) * 4 + 4) throw tooLarge(name);
  let bytes;
  try {
    bytes = base64ToBytes(raw.contentBase64);
  } catch {
    throw badRequest(`"${name}" couldn't be read.`);
  }
  if (bytes.length > MAX_ATTACHMENT_BYTES) throw tooLarge(name);
  return { internetMessageId, key, name, contentType: str(raw.contentType, 200), bytes };
}

function tooLarge(name) {
  const mb = Math.round(MAX_ATTACHMENT_BYTES / 1024 / 1024);
  return new AppError(413, "ATTACHMENT_TOO_LARGE", `"${name}" is over ${mb} MB, too big to copy. Save it from Outlook and add it in Fergus instead.`);
}

function base64ToBytes(b64) {
  const binary = atob(b64.replace(/\s/g, ""));
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function validAddress(a) {
  return {
    address1: str(a?.address1, 300),
    address2: str(a?.address2, 300),
    suburb: str(a?.suburb, 200),
    city: str(a?.city, 200),
    region: str(a?.region, 200),
    postcode: str(a?.postcode, 20),
  };
}

function existingId(value) {
  if (value == null || value === "") return null;
  if (!/^\d+$/.test(String(value))) throw badRequest("That Fergus record ID doesn't look right.");
  return Number(value);
}

export function validateCreate(raw, options) {
  if (!raw || typeof raw !== "object") throw badRequest("Missing job details.");
  const email = validateEmail(raw.email);

  const jobType = str(raw.jobType, 50);
  if (!options.jobTypes.includes(jobType)) throw badRequest(`Job type must be one of: ${options.jobTypes.join(", ")}.`);

  const group = str(raw.group, 200);
  if (group && options.groups.length && !options.groups.includes(group)) throw badRequest("Pick a group from the list.");

  const title = str(raw.title, 255);
  const description = str(raw.description, 10_000);
  if (!title) throw badRequest("Give the job a title.");
  if (!description) throw badRequest("Give the job a description.");

  const customer = { id: existingId(raw.customer?.id), new: null };
  if (customer.id == null) {
    const n = raw.customer?.new || {};
    customer.new = {
      name: str(n.name, 300),
      company: str(n.company, 300),
      contactName: str(n.contactName, 300),
      email: str(n.email, 300),
      phone: str(n.phone, 100),
      address: validAddress(n.address),
    };
    if (!customer.new.name) throw badRequest("Enter the new customer's name, or pick an existing customer.");
  }

  const site = { id: existingId(raw.site?.id), new: null };
  if (site.id == null) {
    const n = raw.site?.new || {};
    site.new = {
      name: str(n.name, 300),
      address: validAddress(n.address),
      contactName: str(n.contactName, 300),
      email: str(n.email, 300),
      phone: str(n.phone, 100),
    };
    if (!site.new.address.address1) throw badRequest("Enter the new site's street address, or pick an existing site.");
  }

  return { email, jobType, group, title, description, customer, site };
}
