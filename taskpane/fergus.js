/* "Create Fergus job" task pane
 *
 * Reads the open email through Office.js, sends it to our backend (which
 * holds the Fergus and Claude credentials — nothing secret lives here), and
 * shows an editable form prefilled with what Claude pulled out of the email
 * plus any matching Fergus customers and sites. On confirm, the backend
 * creates the job and this pane tags the email with an Outlook category,
 * then copies the ticked attachments onto the job, one request each.
 *
 * The email's images (attached, or pasted into the body) go to Claude with
 * it, so details in a screenshot of a customer's text message are read too.
 * Pasted-in images can be copied to the job like attachments; any Claude
 * says are only logos or signature graphics start unticked.
 *
 * Requests to the backend carry a Microsoft 365 access token obtained via
 * nested app authentication (MSAL's createNestablePublicClientApplication),
 * so only signed-in staff in our tenant can create jobs.
 */

const CATEGORY = "Fergus job created";
const MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024; // the backend's limit too
// Images sent to Claude: types it reads, and limits matching the backend's.
const CLAUDE_IMAGE_TYPES = { jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png", gif: "image/gif", webp: "image/webp" };
const MAX_CLAUDE_IMAGES = 8;
const MAX_CLAUDE_IMAGE_BYTES = 7 * 1024 * 1024; // 10 MB once base64
const MAX_CLAUDE_IMAGES_TOTAL = 15 * 1024 * 1024;
// Pasted-in images smaller than this are spacers and tiny icons.
const MIN_INLINE_IMAGE_BYTES = 3 * 1024;
const cfg = window.FERGUS_ADDIN_CONFIG || {};

let email = null;
let analysis = null;
// Set when making another job from an email that already has one: that job,
// which the new one sits alongside.
let wantAnother = false;
let alongsideJob = null;
let decorativeKeys = new Set(); // images Claude called logos or signature graphics
const contentCache = new Map(); // attachment key → content already read from Outlook
let retryTimer = null;

Office.onReady(() => {
  document.getElementById("retryBtn").addEventListener("click", () => start({ again: wantAnother }));
  document.getElementById("signInBtn").addEventListener("click", () => start({ interactive: true, again: wantAnother }));
  document.getElementById("anotherBtn").addEventListener("click", () => start({ again: true }));
  document.getElementById("jobForm").addEventListener("submit", onSubmit);
  document.getElementById("copyAttachBtn").addEventListener("click", () => copyAttachments());
  start();
});

// ---------------------------------------------------------------------------
// Views

function show(view) {
  for (const id of ["statusView", "errorView", "doneView", "jobForm"]) {
    document.getElementById(id).hidden = id !== view;
  }
}

function showStatus(text) {
  document.getElementById("statusText").textContent = text;
  show("statusView");
}

function showError(err) {
  document.getElementById("errorText").textContent = err.message || String(err);
  document.getElementById("signInBtn").hidden = err.code !== "SIGN_IN_NEEDED" && err.code !== "UNAUTHENTICATED";
  countdown(document.getElementById("retryBtn"), "Try again", err.retryAfter);
  show("errorView");
}

// Rate-limited? Keep the retry button disabled until the wait is over, so
// it's obvious when trying again will actually work.
function countdown(button, label, seconds) {
  clearInterval(retryTimer);
  if (!seconds) {
    button.disabled = false;
    button.textContent = label;
    return;
  }
  let left = seconds;
  const tick = () => {
    if (left <= 0) {
      clearInterval(retryTimer);
      button.disabled = false;
      button.textContent = label;
      return;
    }
    button.disabled = true;
    button.textContent = `${label} (${left}s)`;
    left--;
  };
  tick();
  retryTimer = setInterval(tick, 1000);
}

// ---------------------------------------------------------------------------
// Flow

async function start({ interactive = false, again = false } = {}) {
  wantAnother = again;
  try {
    showStatus("Reading this email…");
    if (interactive) await getToken({ interactive: true });
    email = await readEmail();
    const sent = claudeImages();
    if (sent.length) showStatus(`Reading this email's ${sent.length === 1 ? "image" : `${sent.length} images`}…`);
    const images = await readClaudeImages(sent);
    showStatus("Reading the email with Claude and searching Fergus…");
    const result = await api("/api/analyze", { email, images: images.map((i) => i.content), again });
    if (result.status === "exists") {
      showDone(result.job, false);
      return;
    }
    analysis = result;
    alongsideJob = result.currentJob || null;
    // Claude numbers the images it was sent from 1.
    decorativeKeys = new Set((result.extraction?.decorativeImages || []).map((n) => images[n - 1]?.att.key).filter(Boolean));
    renderForm(result);
  } catch (err) {
    showError(err);
  }
}

async function onSubmit(event) {
  event.preventDefault();
  const formError = document.getElementById("formError");
  const button = document.getElementById("createBtn");
  formError.hidden = true;

  let payload;
  try {
    payload = buildPayload();
  } catch (err) {
    formError.textContent = err.message;
    formError.hidden = false;
    return;
  }

  const selected = selectedAttachmentKeys();
  button.disabled = true;
  button.textContent = "Creating job in Fergus…";
  try {
    const result = await api("/api/jobs", payload);
    showDone(result.job, result.status === "created", selected);
  } catch (err) {
    // Stay on the form so nothing typed is lost; the backend remembers any
    // customer/site it already created, so trying again won't duplicate them.
    formError.textContent = err.message;
    formError.hidden = false;
    countdown(button, alongsideJob ? "Create another job in Fergus" : "Create job in Fergus", err.retryAfter);
  }
}

// `toCopy`: keys of attachments ticked on the form, copied straight away.
// Without it (reopening an email already actioned) any not yet copied are
// listed with a button, rather than copied without anyone asking.
async function showDone(job, created, toCopy = null) {
  wantAnother = false;
  alongsideJob = null;
  const label = job.jobNo ? `job ${job.jobNo}` : "a job";
  document.getElementById("doneHeading").textContent = created ? `Created ${label} in Fergus` : `Already in Fergus as ${label}`;

  let detail = "";
  if (created) {
    const made = [job.customerCreated && "customer", job.siteCreated && (job.siteUnknown ? "placeholder site" : "site")].filter(Boolean);
    detail = made.length ? `Also created a new ${made.join(" and ")}. The email was added as a job note.` : "The email was added as a job note.";
    if (job.siteUnknown) detail += ' The site\'s address is "Unknown": update it in Fergus once you know it.';
  } else {
    const when = job.createdAt ? new Date(job.createdAt).toLocaleString() : "";
    detail = ["This email was already turned into a job", when && `on ${when}`, job.createdBy && `by ${job.createdBy}`].filter(Boolean).join(" ") + ".";
  }
  document.getElementById("doneDetail").textContent = detail;

  const link = document.getElementById("jobLink");
  link.href = job.url;
  link.onclick = (e) => openJob(e, job.url);
  showPreviousJobs(job.previousJobs || []);
  // Offered only when reopening, not straight after making a job.
  document.getElementById("anotherBtn").hidden = created;
  show("doneView");
  showAttachments(job, toCopy);

  const note = document.getElementById("categoryNote");
  const tagged = await addCategory();
  if (!tagged.ok) {
    note.textContent =
      `Couldn't tag this email "${CATEGORY}". If that category doesn't exist in your Outlook yet, ` +
      "create it once (Categorize → Manage categories), then reopen this pane to tag it.";
    note.hidden = false;
  }
}

function openJob(e, url) {
  if (Office.context.requirements.isSetSupported("OpenBrowserWindowApi", "1.1")) {
    e.preventDefault();
    Office.context.ui.openBrowserWindow(url);
  }
}

function showPreviousJobs(jobs) {
  const p = document.getElementById("previousJobs");
  p.hidden = jobs.length === 0;
  if (!jobs.length) return;
  const links = jobs.map((j) => {
    const a = document.createElement("a");
    a.href = j.url;
    a.target = "_blank";
    a.rel = "noopener";
    a.textContent = j.jobNo ? `job ${j.jobNo}` : "a job";
    a.onclick = (e) => openJob(e, j.url);
    return a;
  });
  p.replaceChildren(`Earlier ${jobs.length === 1 ? "job" : "jobs"} from this email: `);
  links.forEach((a, i) => p.append(...(i ? [", ", a] : [a])));
}

// ---------------------------------------------------------------------------
// Office.js

function readEmail() {
  const item = Office.context.mailbox.item;
  if (!item) return Promise.reject(uiError("Open an email first."));
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(uiError("Couldn't read this email's body. Try reopening it."));
        return;
      }
      resolve({
        internetMessageId: item.internetMessageId || "",
        subject: item.subject || "",
        from: { name: item.from?.displayName || "", address: item.from?.emailAddress || "" },
        // For a received message, the item is created in our mailbox when it arrives.
        receivedAt: item.dateTimeCreated instanceof Date ? item.dateTimeCreated.toISOString() : "",
        bodyText: result.value || "",
        // Inline ones are mostly signature logos; leave them out.
        attachments: (item.attachments || [])
          .filter((a) => !a.isInline)
          .map((a) => ({ name: a.name, size: a.size, contentType: a.contentType || "" })),
      });
    });
  });
}

// The email's attachments as this pane works with them. `key` names one
// the same way for everyone (Outlook's attachment IDs differ per mailbox),
// so the backend can tell which are already on the job.
function emailAttachments() {
  const item = Office.context.mailbox.item;
  return (item?.attachments || [])
    // Of what's pasted into the body, only images are worth having, and not
    // the tiniest (spacers, icons).
    .filter((a) => !a.isInline || (imageType(a) && a.size >= MIN_INLINE_IMAGE_BYTES))
    .map((a) => {
      let problem = "";
      if (a.attachmentType === Office.MailboxEnums.AttachmentType.Cloud) problem = "A link to a cloud file — add it in Fergus by hand.";
      else if (a.size > MAX_ATTACHMENT_BYTES) problem = `Over ${MAX_ATTACHMENT_BYTES / 1024 / 1024} MB — save it from Outlook and add it in Fergus.`;
      return { id: a.id, name: a.name || "attachment", size: a.size || 0, contentType: a.contentType || imageType(a) || "", inline: Boolean(a.isInline), key: `${a.name}|${a.size}`, problem };
    });
}

// The image type Claude would read it as, or "" if it can't. Pasted-in
// images often have no content type, so the extension decides.
function imageType(a) {
  const ct = (a.contentType || "").toLowerCase();
  if (Object.values(CLAUDE_IMAGE_TYPES).includes(ct)) return ct;
  if (ct === "image/jpg") return "image/jpeg";
  return CLAUDE_IMAGE_TYPES[(a.name || "").split(".").pop().toLowerCase()] || "";
}

// Images left out before reaching the server, and why, for the form's
// warnings: { name, reason }.
let imagesSkippedHere = [];

// The images to show Claude, within its limits, in the email's order.
// Only image-looking files count: a PDF isn't "skipped".
function claudeImages() {
  const picked = [];
  imagesSkippedHere = [];
  let total = 0;
  for (const att of emailAttachments()) {
    const looksLikeImage = /^image\//i.test(att.contentType) || imageType(att);
    if (!looksLikeImage) continue;
    const skip = (reason) => imagesSkippedHere.push({ name: att.name, reason });
    if (!imageType(att)) skip("Claude can't read this type of image");
    else if (picked.length >= MAX_CLAUDE_IMAGES) skip(`only the first ${MAX_CLAUDE_IMAGES} images are read`);
    else if (att.size > MAX_CLAUDE_IMAGE_BYTES) skip("it's too large");
    else if (total + att.size > MAX_CLAUDE_IMAGES_TOTAL) skip("the email's images are too large in total");
    else {
      total += att.size;
      picked.push(att);
    }
  }
  return picked;
}

// An image Outlook can't read is left out; the email is read without it.
async function readClaudeImages(atts) {
  const read = [];
  for (const att of atts) {
    try {
      const c = await readAttachment(att);
      read.push({ att, content: { name: c.name, contentType: imageType(att), contentBase64: c.contentBase64, inline: att.inline } });
    } catch (err) {
      imagesSkippedHere.push({ name: att.name, reason: err.message || "Outlook couldn't read it" });
    }
  }
  return read;
}

// Base64 for a file; an attached email comes back as .eml text. Each is
// read from Outlook once: images sent to Claude are reused when copied.
async function readAttachment(att) {
  if (!contentCache.has(att.key)) contentCache.set(att.key, readAttachmentFromOutlook(att));
  try {
    return await contentCache.get(att.key);
  } catch (err) {
    contentCache.delete(att.key);
    throw err;
  }
}

function readAttachmentFromOutlook(att) {
  const item = Office.context.mailbox.item;
  return new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(att.id, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(uiError(`Outlook couldn't read it${result.error?.message ? ` (${result.error.message})` : ""}.`));
        return;
      }
      const { format, content } = result.value;
      const F = Office.MailboxEnums.AttachmentContentFormat;
      if (format === F.Base64) return resolve({ name: att.name, contentType: att.contentType, contentBase64: content });
      if (format === F.Eml || format === F.ICalendar) {
        const ext = format === F.Eml ? ".eml" : ".ics";
        const name = att.name.toLowerCase().endsWith(ext) ? att.name : `${att.name}${ext}`;
        return resolve({ name, contentType: format === F.Eml ? "message/rfc822" : "text/calendar", contentBase64: textToBase64(content) });
      }
      reject(uiError("It's a link to a cloud file — add it in Fergus by hand."));
    });
  });
}

function textToBase64(text) {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}

// Categories must already exist in the mailbox's master list to be applied;
// creating one there needs the broader ReadWriteMailbox permission, which
// this add-in deliberately doesn't ask for.
function addCategory() {
  const item = Office.context.mailbox.item;
  if (!item?.categories) return Promise.resolve({ ok: false });
  return new Promise((resolve) => {
    item.categories.getAsync((current) => {
      const has = current.status === Office.AsyncResultStatus.Succeeded && (current.value || []).some((c) => c.displayName === CATEGORY);
      if (has) return resolve({ ok: true });
      item.categories.addAsync([CATEGORY], (added) => {
        resolve({ ok: added.status === Office.AsyncResultStatus.Succeeded });
      });
    });
  });
}

// ---------------------------------------------------------------------------
// Backend calls

function uiError(message, extra = {}) {
  return Object.assign(new Error(message), extra);
}

let msalApp = null;

async function getToken({ interactive = false } = {}) {
  if (!cfg.entraClientId) return null; // local testing against AUTH_DISABLED backend

  if (!Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")) {
    throw uiError("This version of Outlook can't sign you in to add-ins. Update Outlook, or use Outlook on the web.");
  }
  if (!msalApp) {
    msalApp = await msal.createNestablePublicClientApplication({
      auth: { clientId: cfg.entraClientId, authority: `https://login.microsoftonline.com/${cfg.entraTenantId}` },
    });
  }
  const request = { scopes: [`api://${cfg.entraClientId}/access_as_user`] };
  try {
    return (await msalApp.acquireTokenSilent(request)).accessToken;
  } catch (silentErr) {
    try {
      return (await msalApp.acquireTokenPopup(request)).accessToken;
    } catch (popupErr) {
      if (interactive) throw uiError(`Sign-in didn't complete: ${popupErr.message || popupErr}`);
      throw uiError("Sign in to Microsoft 365 to create Fergus jobs.", { code: "SIGN_IN_NEEDED" });
    }
  }
}

async function api(path, body) {
  if (!cfg.backendUrl || cfg.backendUrl.includes("YOUR-BACKEND-HOST")) {
    throw uiError("The add-in isn't set up yet: backendUrl is missing from taskpane/config.js.");
  }
  const token = await getToken();
  let response;
  try {
    response = await fetch(`${cfg.backendUrl}${path}`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...(token ? { Authorization: `Bearer ${token}` } : {}) },
      body: JSON.stringify(body),
    });
  } catch {
    throw uiError("Couldn't reach the Fergus job server. Check your connection; if it keeps happening, the server may be down.");
  }
  let data = {};
  try {
    data = await response.json();
  } catch {
    // Non-JSON error page from a proxy; fall through to the generic message.
  }
  if (!response.ok) {
    const e = data.error || {};
    const retryAfter = e.retryAfter ?? (Number(response.headers.get("Retry-After")) || undefined);
    throw uiError(e.message || `The Fergus job server returned an error (HTTP ${response.status}).`, { code: e.code, retryAfter });
  }
  return data;
}

// ---------------------------------------------------------------------------
// Copying attachments onto the job

let attachState = null; // { job, rows: [{ att, status: "copied"|"pending"|"copying"|"failed"|"skipped", message }] }

function showAttachments(job, toCopy) {
  const atts = emailAttachments();
  const copied = new Set(job.attachments || []);
  attachState = {
    job,
    rows: atts.map((att) => {
      if (copied.has(att.key)) return { att, status: "copied", checked: false };
      if (att.problem) return { att, status: "skipped", message: att.problem, checked: false };
      // Reopened later, pasted-in images (often logos) wait to be ticked.
      return { att, status: "pending", checked: toCopy ? toCopy.includes(att.key) : !att.inline };
    }),
  };
  document.getElementById("attachView").hidden = atts.length === 0;
  renderAttachments();
  if (toCopy?.length) copyAttachments();
}

function renderAttachments() {
  const list = document.getElementById("attachList");
  const rows = attachState.rows;
  const busy = rows.some((r) => r.status === "copying");
  list.replaceChildren(
    ...rows.map((row) => {
      const li = document.createElement("li");
      const label = document.createElement("label");
      if (row.status === "pending" || row.status === "failed") {
        const box = document.createElement("input");
        box.type = "checkbox";
        box.checked = row.checked;
        box.disabled = busy;
        box.addEventListener("change", () => {
          row.checked = box.checked;
          renderAttachments();
        });
        label.append(box);
      }
      const name = document.createElement("span");
      name.className = "attach-name";
      name.textContent = `${row.att.name} (${row.att.inline ? "in the email body, " : ""}${formatSize(row.att.size)})`;
      label.append(name);
      const status = document.createElement("span");
      status.className = `attach-status ${row.status === "copied" ? "ok" : row.status === "failed" ? "failed" : ""}`;
      status.textContent = { copied: "On the job", copying: "Copying…", failed: "Failed", skipped: "Not copied", pending: "" }[row.status];
      li.append(label, status);
      if (row.message) {
        const msg = document.createElement("div");
        msg.className = `small ${row.status === "failed" ? "attach-status failed" : "muted"}`;
        msg.textContent = row.message;
        li.style.flexWrap = "wrap";
        msg.style.flexBasis = "100%";
        msg.style.textAlign = "left";
        li.append(msg);
      }
      return li;
    })
  );

  const ticked = rows.filter((r) => r.checked && (r.status === "pending" || r.status === "failed")).length;
  const button = document.getElementById("copyAttachBtn");
  button.hidden = busy || !rows.some((r) => r.status === "pending" || r.status === "failed");
  button.disabled = ticked === 0;
  button.textContent = rows.some((r) => r.status === "failed") ? `Try again (${ticked})` : `Copy ${ticked} to Fergus`;

  const done = rows.filter((r) => r.status === "copied").length;
  const failed = rows.filter((r) => r.status === "failed").length;
  const copying = rows.findIndex((r) => r.status === "copying");
  const summary = document.getElementById("attachSummary");
  if (copying >= 0) summary.textContent = `Copying attachments to Fergus… ${attachState.progress.done + 1} of ${attachState.progress.total}`;
  else if (failed) summary.textContent = `${failed} couldn't be copied. Tick them and try again, or add them in Fergus by hand.`;
  else summary.textContent = done ? `${done} on the job in Fergus.` : "";
}

// One at a time: keeps each request small and in step with Fergus's rate limit.
async function copyAttachments() {
  const queue = attachState.rows.filter((r) => r.checked && (r.status === "pending" || r.status === "failed"));
  attachState.progress = { done: 0, total: queue.length };
  for (const row of queue) {
    row.status = "copying";
    row.checked = false;
    row.message = "";
    renderAttachments();
    try {
      const content = await readAttachment(row.att);
      const result = await api("/api/jobs/attachments", {
        internetMessageId: email.internetMessageId,
        key: row.att.key,
        ...content,
      });
      row.status = "copied";
      attachState.job = result.job;
    } catch (err) {
      row.status = "failed";
      row.checked = true;
      row.message = err.message;
    }
    attachState.progress.done++;
    renderAttachments();
  }
}

function selectedAttachmentKeys() {
  return [...document.querySelectorAll('#attachChoices input[type="checkbox"]:checked')].map((b) => b.value);
}

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

// ---------------------------------------------------------------------------
// Form

const CONFIDENCE_LABEL = { medium: "inferred — check", low: "guess — check" };

function value(field) {
  return analysis?.extraction?.[field]?.value || "";
}

function setField(name, v) {
  const el = document.querySelector(`[name="${name}"]`);
  if (el) el.value = v || "";
}

function getField(name) {
  return (document.querySelector(`[name="${name}"]`)?.value || "").trim();
}

function renderForm(result) {
  const x = result.extraction;

  const warnings = document.getElementById("warnings");
  warnings.replaceChildren();
  const warn = (text) => {
    const p = document.createElement("p");
    p.className = "warning";
    p.textContent = text;
    warnings.append(p);
  };
  if (result.extractionError) warn(`${result.extractionError.message} The form is blank — fill it in from the email.`);
  if (result.searchError) warn(`Couldn't search Fergus for existing customers and sites: ${result.searchError.message}`);
  // Say which images Claude saw, so a missed screenshot is obvious.
  const imagesRead = result.images?.read || [];
  for (const { name, reason } of [...imagesSkippedHere, ...(result.images?.skipped || [])]) {
    warn(`Claude didn't see the image "${name}": ${reason.replace(/\.$/, "")}. Check it yourself for customer details.`);
  }
  if (imagesRead.length) {
    const p = document.createElement("p");
    p.className = "muted small";
    p.textContent = `Claude read ${imagesRead.length === 1 ? "the image" : `${imagesRead.length} images`}: ${imagesRead.join(", ")}.`;
    warnings.append(p);
  }
  if (alongsideJob) {
    const label = alongsideJob.jobNo ? `job ${alongsideJob.jobNo}` : "a job";
    warn(`This email is already ${label} in Fergus. Creating will make a separate, new job; ${label} isn't changed.`);
  }
  if (result.partial?.customerId || result.partial?.siteId || result.partial?.jobId) {
    warn("An earlier attempt for this email stopped part-way. Whatever it already created in Fergus will be reused rather than duplicated.");
  }

  const notes = document.getElementById("claudeNotes");
  notes.textContent = x?.notes || "";
  notes.hidden = !x?.notes;

  setField("customerName", value("customerName"));
  setField("company", value("company"));
  setField("contactName", value("contactName"));
  setField("email", value("email"));
  setField("phone", value("phone"));
  const site = x?.siteAddress || {};
  for (const k of ["address1", "address2", "suburb", "city", "region", "postcode"]) setField(k, site[k]);
  setField("title", value("jobTitle") || email.subject);
  setField("description", value("jobDescription"));

  for (const span of document.querySelectorAll(".conf")) {
    const key = span.dataset.for;
    const conf = key === "siteAddress" ? site.confidence : x?.[key]?.confidence;
    span.textContent = CONFIDENCE_LABEL[conf] || "";
    span.className = `conf ${conf || ""}`;
  }

  const { jobTypes, defaultJobType, groups } = result.options;
  fillSelect("jobType", jobTypes.map((t) => [t, t]), defaultJobType);
  fillSelect("group", [["", groups.length ? "— None —" : "— No groups configured —"], ...groups.map((g) => [g, g])], value("group"));

  // Only preselect an existing record on a strong match — a wrong customer
  // is worse than one extra click.
  const customers = result.candidates.customers;
  const strongCustomer = customers.find((c) => c.matchedOn.includes("email") || c.matchedOn.includes("phone"));
  renderChoices("customer", customers, strongCustomer?.id, (c) => ({
    main: c.name || "(no name)",
    detail: [c.contactName !== c.name && c.contactName, c.email, c.phone, c.address, `matched on ${c.matchedOn.join(", ")}`].filter(Boolean).join(" · "),
  }), "New customer");

  const sites = result.candidates.sites;
  const strongSite = sites.find((s) => s.score >= 3);
  // With no address in the email, a placeholder site is the likely choice.
  const siteKnown = Boolean(site.address1) && site.confidence !== "none";
  renderChoices("site", sites, strongSite?.id ?? (siteKnown ? null : "unknown"), (s) => ({
    main: s.address || s.name || "(no address)",
    detail: [s.name && s.address && s.name, s.contactName].filter(Boolean).join(" · "),
  }), "New site", {
    value: "unknown",
    main: "Site unknown",
    detail: 'Creates a site with the address "Unknown", to update in Fergus later.',
  });

  renderAttachmentChoices();

  document.getElementById("formError").hidden = true;
  const button = document.getElementById("createBtn");
  button.disabled = false;
  button.textContent = alongsideJob ? "Create another job in Fergus" : "Create job in Fergus";
  show("jobForm");
}

function renderAttachmentChoices() {
  const atts = emailAttachments();
  document.getElementById("attachFieldset").hidden = atts.length === 0;
  document.getElementById("attachChoices").replaceChildren(
    ...atts.map((att) => {
      const label = document.createElement("label");
      label.className = "choice";
      const box = document.createElement("input");
      box.type = "checkbox";
      box.value = att.key;
      const decorative = decorativeKeys.has(att.key);
      box.checked = !att.problem && !decorative;
      box.disabled = Boolean(att.problem);
      const text = document.createElement("span");
      const main = document.createElement("span");
      main.className = "attach-name";
      main.textContent = att.name;
      const detail = document.createElement("div");
      detail.className = "choice-detail";
      detail.textContent =
        att.problem ||
        [att.inline && "In the email body", formatSize(att.size), decorative && "looks like a logo or signature image"].filter(Boolean).join(" · ");
      text.append(main, detail);
      label.append(box, text);
      return label;
    })
  );
}

function fillSelect(name, options, selected) {
  const select = document.querySelector(`[name="${name}"]`);
  select.replaceChildren(
    ...options.map(([v, label]) => {
      const o = document.createElement("option");
      o.value = v;
      o.textContent = label;
      o.selected = v === selected;
      return o;
    })
  );
}

// `extra`: one more option after "New …", e.g. a placeholder site.
function renderChoices(kind, records, selectedId, describe, newLabel, extra = null) {
  const container = document.getElementById(`${kind}Choices`);
  const fields = document.getElementById(kind === "customer" ? "newCustomerFields" : "newSiteFields");
  const name = `${kind}Choice`;

  const option = (val, main, detail, checked) => {
    const label = document.createElement("label");
    label.className = "choice";
    const input = document.createElement("input");
    input.type = "radio";
    input.name = name;
    input.value = val;
    input.checked = checked;
    input.addEventListener("change", () => (fields.hidden = val !== "new"));
    const text = document.createElement("span");
    const m = document.createElement("span");
    m.className = "choice-main";
    m.textContent = main;
    text.append(m);
    if (detail) {
      const d = document.createElement("div");
      d.className = "choice-detail";
      d.textContent = detail;
      text.append(d);
    }
    label.append(input, text);
    return label;
  };

  const items = records.map((r) => {
    const { main, detail } = describe(r);
    return option(String(r.id), main, detail, r.id === selectedId);
  });
  const isNew = selectedId == null;
  items.push(option("new", newLabel, records.length ? "" : "No matches found in Fergus.", isNew));
  if (extra) items.push(option(extra.value, extra.main, extra.detail, selectedId === extra.value));
  container.replaceChildren(...items);
  fields.hidden = !isNew;
}

function choice(kind) {
  return document.querySelector(`input[name="${kind}Choice"]:checked`)?.value || "new";
}

function buildPayload() {
  const title = getField("title");
  const description = getField("description");
  if (!title) throw uiError("Give the job a title.");
  if (!description) throw uiError("Give the job a description.");

  const contact = {
    contactName: getField("contactName"),
    email: getField("email"),
    phone: getField("phone"),
  };
  const address = {
    address1: getField("address1"),
    address2: getField("address2"),
    suburb: getField("suburb"),
    city: getField("city"),
    region: getField("region"),
    postcode: getField("postcode"),
  };

  const customerChoice = choice("customer");
  const customer =
    customerChoice === "new"
      ? // No customer address: the site address is often not the customer's
        // own (a builder's client's house, say), so it's left for Fergus.
        { new: { name: getField("customerName"), company: getField("company"), ...contact } }
      : { id: customerChoice };
  if (customer.new && !customer.new.name) throw uiError("Enter the customer's name, or pick an existing customer.");

  const siteChoice = choice("site");
  const site =
    siteChoice === "new" ? { new: { address, ...contact } } : siteChoice === "unknown" ? { unknown: true, ...contact } : { id: siteChoice };
  if (site.new && !address.address1) throw uiError("Enter the site's street address, or pick an existing site.");

  return {
    email,
    customer,
    site,
    jobType: getField("jobType"),
    group: getField("group"),
    title,
    description,
    alongsideJobId: alongsideJob?.id ?? null,
  };
}
