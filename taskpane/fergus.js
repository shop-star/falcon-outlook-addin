/* "Create Fergus job" task pane
 *
 * Reads the open email through Office.js, sends it to our backend (which
 * holds the Fergus and Claude credentials — nothing secret lives here), and
 * shows an editable form prefilled with what Claude pulled out of the email
 * plus any matching Fergus customers and sites. On confirm, the backend
 * creates the job and this pane tags the email with an Outlook category.
 *
 * Requests to the backend carry a Microsoft 365 access token obtained via
 * nested app authentication (MSAL's createNestablePublicClientApplication),
 * so only signed-in staff in our tenant can create jobs.
 */

const CATEGORY = "Fergus job created";
const cfg = window.FERGUS_ADDIN_CONFIG || {};

let email = null;
let analysis = null;
let retryTimer = null;

Office.onReady(() => {
  document.getElementById("retryBtn").addEventListener("click", () => start());
  document.getElementById("signInBtn").addEventListener("click", () => start({ interactive: true }));
  document.getElementById("jobForm").addEventListener("submit", onSubmit);
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

async function start({ interactive = false } = {}) {
  try {
    showStatus("Reading this email…");
    if (interactive) await getToken({ interactive: true });
    email = await readEmail();
    showStatus("Reading the email with Claude and searching Fergus…");
    const result = await api("/api/analyze", { email });
    if (result.status === "exists") {
      showDone(result.job, false);
      return;
    }
    analysis = result;
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

  button.disabled = true;
  button.textContent = "Creating job in Fergus…";
  try {
    const result = await api("/api/jobs", payload);
    showDone(result.job, result.status === "created");
  } catch (err) {
    // Stay on the form so nothing typed is lost; the backend remembers any
    // customer/site it already created, so trying again won't duplicate them.
    formError.textContent = err.message;
    formError.hidden = false;
    countdown(button, "Create job in Fergus", err.retryAfter);
  }
}

async function showDone(job, created) {
  const label = job.jobNo ? `job ${job.jobNo}` : "a job";
  document.getElementById("doneHeading").textContent = created ? `Created ${label} in Fergus` : `Already in Fergus as ${label}`;

  let detail = "";
  if (created) {
    const made = [job.customerCreated && "customer", job.siteCreated && "site"].filter(Boolean);
    detail = made.length ? `Also created a new ${made.join(" and ")}. The email was added as a job note.` : "The email was added as a job note.";
  } else {
    const when = job.createdAt ? new Date(job.createdAt).toLocaleString() : "";
    detail = ["This email was already turned into a job", when && `on ${when}`, job.createdBy && `by ${job.createdBy}`].filter(Boolean).join(" ") + ".";
  }
  document.getElementById("doneDetail").textContent = detail;

  const link = document.getElementById("jobLink");
  link.href = job.url;
  link.onclick = (e) => {
    if (Office.context.requirements.isSetSupported("OpenBrowserWindowApi", "1.1")) {
      e.preventDefault();
      Office.context.ui.openBrowserWindow(job.url);
    }
  };
  show("doneView");

  const note = document.getElementById("categoryNote");
  const tagged = await addCategory();
  if (!tagged.ok) {
    note.textContent =
      `Couldn't tag this email "${CATEGORY}". If that category doesn't exist in your Outlook yet, ` +
      "create it once (Categorize → Manage categories), then reopen this pane to tag it.";
    note.hidden = false;
  }
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
        attachments: (item.attachments || [])
          .filter((a) => !a.isInline)
          .map((a) => ({ name: a.name, size: a.size, contentType: a.contentType || "" })),
      });
    });
  });
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
  renderChoices("site", sites, strongSite?.id, (s) => ({
    main: s.address || s.name || "(no address)",
    detail: [s.name && s.address && s.name, s.contactName].filter(Boolean).join(" · "),
  }), "New site");

  document.getElementById("formError").hidden = true;
  const button = document.getElementById("createBtn");
  button.disabled = false;
  button.textContent = "Create job in Fergus";
  show("jobForm");
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

function renderChoices(kind, records, selectedId, describe, newLabel) {
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
  const site = siteChoice === "new" ? { new: { address, ...contact } } : { id: siteChoice };
  if (site.new && !address.address1) throw uiError("Enter the site's street address, or pick an existing site.");

  return {
    email,
    customer,
    site,
    jobType: getField("jobType"),
    group: getField("group"),
    title,
    description,
  };
}
