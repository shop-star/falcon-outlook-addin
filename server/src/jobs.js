/* The two things the task pane asks for:
 *   analyze(email)  → prefilled form: Claude's extraction + Fergus match candidates
 *   create(request) → customer/site (if new) → job → email as a job note
 */
import { JOB_TYPES } from "./config.js";
import { AppError, badRequest } from "./errors.js";
import { extractJobDetails } from "./extract.js";
import { findCandidates } from "./match.js";

const MAX_BODY_CHARS = 200_000;

export class JobService {
  constructor({ config, fergus, anthropic, store }) {
    this.config = config;
    this.fergus = fergus;
    this.anthropic = anthropic;
    this.store = store;
    this.inFlight = new Set();
  }

  options() {
    const { groups, defaultJobType } = this.config.fergus;
    return { jobTypes: JOB_TYPES, defaultJobType, groups };
  }

  async analyze(rawEmail) {
    const email = validateEmail(rawEmail);

    const existing = await this.store.get(email.internetMessageId);
    if (existing?.status === "done") {
      return { status: "exists", job: publicJob(existing) };
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
        email
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
      // A previous attempt that stopped part-way: the pane says what will be reused.
      partial: existing ? { customerId: existing.customerId ?? null, siteId: existing.siteId ?? null, jobId: existing.jobId ?? null } : null,
    };
  }

  async create(raw, user) {
    const req = validateCreate(raw, this.options());
    const id = req.email.internetMessageId;

    if (this.inFlight.has(id)) {
      throw new AppError(409, "IN_PROGRESS", "A Fergus job is already being created for this email. Wait a moment, then reopen the pane.");
    }
    this.inFlight.add(id);
    try {
      return await this.createSteps(req, user);
    } finally {
      this.inFlight.delete(id);
    }
  }

  async createSteps(req, user) {
    const id = req.email.internetMessageId;
    let record = await this.store.get(id);
    if (record?.status === "done") {
      return { status: "exists", job: publicJob(record) };
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
      await save({ customerId, siteId, jobId: job.id, jobNo: job.jobNo, jobUrl: this.jobUrl(job) });
    }

    if (!record.noteAdded) {
      await this.fergus.addJobNote(record.jobId, noteText(req, user));
      await save({ noteAdded: true });
    }

    record = await this.store.update(id, { status: "done", group: req.group, createdBy: user?.email || user?.name || "" });
    return { status: "created", job: publicJob(record) };
  }

  jobUrl(job) {
    return this.config.fergus.jobUrlTemplate
      .replace("{jobId}", encodeURIComponent(job.id ?? ""))
      .replace("{jobNo}", encodeURIComponent(job.jobNo ?? ""));
  }
}

function publicJob(record) {
  return {
    id: record.jobId,
    jobNo: record.jobNo ?? null,
    url: record.jobUrl,
    createdAt: record.updatedAt,
    createdBy: record.createdBy || "",
    customerCreated: Boolean(record.customerCreated),
    siteCreated: Boolean(record.siteCreated),
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
  if (attachments.length) lines.push(`Attachments (in Outlook): ${attachments.join(", ")}`);
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
