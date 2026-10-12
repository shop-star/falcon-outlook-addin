# City Timber Outlook add-in

An Outlook add-in that adds two buttons to the ribbon when you're reading
an email:

- **Create Fergus job** — opens a task pane that reads the email, has
  Claude pull out the customer, contact, site address and job details
  (including from forwarded emails and builders writing on a client's
  behalf), finds matching customers and sites already in Fergus, and lets
  you check and edit everything before creating the job. The original email
  is attached to the job as a note and the email is tagged with the Outlook
  category **Fergus job created**. Details below.
- **Surface Takeoff** — opens the Surface Takeoff plan-measuring app in a
  real browser window, with the email's subject, PDF/zip attachment names
  and links carried over. The app itself now lives in its own repo,
  [shop-star/surface-takeoffs](https://github.com/shop-star/surface-takeoffs);
  this repo only holds the button (`commands.js`).

What's where:

| Path | What it is |
| --- | --- |
| `manifest.xml` | The add-in manifest (both buttons) |
| `commands.html`, `commands.js` | Surface Takeoff button: opens the app in a browser window |
| `taskpane/` | Create Fergus job task pane (static, hosted on GitHub Pages) |
| `server/` | Backend for the task pane: a Cloudflare Worker doing the Claude extraction and Fergus API calls. Holds the secrets; deployed by `.github/workflows/deploy-worker.yml` |
| `vendor/msal-browser/` | Microsoft's sign-in library, used by the task pane |

## Hosting

`.github/workflows/deploy-pages.yml` deploys the static add-in files
(`manifest.xml`, `commands.*`, `assets/`, `taskpane/`, `vendor/`) to GitHub
Pages on every push to `main`. `server/` is **not** part of that: it holds
API keys, so it runs as a Cloudflare Worker (see below). If you fork
this, update every `https://shop-star.github.io/falcon-outlook-addin` URL
in `manifest.xml`.

## Sideload the add-in into Outlook (work/school account)

**Outlook on the web** (works.outlook.com / outlook.office.com):

1. Open any email, then in the ribbon go to **Get Add-ins** (or the "..."
   menu → **Get Add-ins**).
2. Choose **My add-ins** (left side), scroll to **Custom Addins**, and
   click **Add a custom add-in → Add from file...**
3. Select `manifest.xml`. Accept the warning about unverified/custom
   add-ins.

**New Outlook for Windows / Outlook on Mac:** same steps — **Get Add-ins**
→ **My add-ins** → **Add a custom add-in** → **Add from file**.

**Classic Outlook for Windows:** **File → Manage Add-ins** opens the same
web dialog above, or use **Home → Get Add-ins → My add-ins**.

Once added, open any email and you'll see **Surface Takeoff** and
**Create Fergus job** buttons in the ribbon.

> **Note:** some organizations restrict end-users from sideloading custom
> add-ins (a tenant admin setting). If "Add from file" is missing or
> blocked, your Microsoft 365 admin can either enable it for you or install
> the manifest for you centrally via the Microsoft 365 admin center
> (Settings → Integrated apps → Upload custom apps).
>
> Also note: opening a real browser window from the add-in
> (`OpenBrowserWindowApi`) needs a reasonably current Outlook build. If your
> client is too old, clicking the ribbon button shows an error banner
> instead — open the app's URL directly in your browser as a fallback.


## Create Fergus job

### How it works

1. The task pane reads the open email with Office.js — subject, body as
   plain text, sender, received time, attachment names and the email's
   `internetMessageId` — and POSTs it to the backend (`/api/analyze`).
   Attachment *contents* are never sent.
2. The backend asks Claude to extract customer name, company, contact name,
   phone, email, site address, a job title and description, and a
   suggested group, each with a confidence (high / medium / low / none).
   Claude reads the whole body, so a forwarded email or a builder writing
   for their client still yields the real customer and site. It then
   searches Fergus for existing customers (by email, phone, name and
   company) and sites (by street address and postcode).
3. The pane shows an editable form. An existing customer is preselected
   only when it matched on email or phone, and an existing site only when
   street and postcode both match; otherwise **New customer** / **New
   site** is selected with the extracted details filled in. Fields Claude
   inferred or guessed are flagged "check".
4. **Create job in Fergus** (`/api/jobs`) creates the customer and site if
   new, creates the job (`isDraft: false`), adds the email (from, date,
   subject, attachment names, body, and the chosen group) as a job note,
   and returns the job number and link. The pane then tags the email
   **Fergus job created**.
5. The backend records `internetMessageId → Fergus job` in
   a Cloudflare D1 database. Opening the pane on an email that's
   already been actioned — by anyone, on any device — shows that job
   instead of the form. If a create fails part-way (e.g. Fergus rate
   limits the note), trying again reuses the customer, site and job already
   created rather than duplicating them.

Fergus errors are shown in plain words: a rejected token, validation
messages from Fergus, and rate limits (a short `Retry-After` is retried
automatically once; a longer one shows a countdown on the button and keeps
the form as typed).

### Not supported by the Fergus API (as far as we could verify)

- **Business-unit groups:** we found no endpoint for listing groups or
  setting a job's group. The group list comes from `FERGUS_GROUPS`, Claude
  suggests one, and the chosen group is written at the top of the job note
  for whoever picks the job up to set in Fergus.
- **Separate contacts:** a new customer's contact is created as the
  customer's main contact, and a new site's contact is the same person as
  its default contact. Nothing creates a standalone contact record.
- **Job link:** the API returns the job's ID and number but not a web-app
  URL. The link is built from `FERGUS_JOB_URL_TEMPLATE` — open one real job
  in Fergus and check the template matches its URL.

### 1. Register the add-in in Microsoft Entra ID

The backend accepts requests only from people signed in to your Microsoft
365 tenant. The task pane gets their token from Outlook itself (nested app
authentication), so nobody types a password into the add-in.

1. **Entra admin center → App registrations → New registration.** Name it
   e.g. "Fergus job add-in", single tenant. Note the **Application (client)
   ID** and **Directory (tenant) ID**.
2. **Authentication → Add a platform → Single-page application**, redirect
   URI `brk-multihub://shop-star.github.io` (the add-in's host).
3. **Expose an API →** set the Application ID URI to `api://<client-id>`,
   then **Add a scope** named `access_as_user` (who can consent: admins and
   users).
4. **API permissions → Add a permission → My APIs →** this app →
   `access_as_user`, then **Grant admin consent**.
5. Put the client and tenant IDs in `taskpane/config.js` and in the
   backend's `ENTRA_CLIENT_ID` / `ENTRA_TENANT_ID`.

Sign-in needs an Outlook that supports nested app authentication (current
Outlook on the web, new Outlook for Windows, classic Outlook on a recent
Microsoft 365 build, Outlook for Mac). Older clients show a message saying so.

### 2. Run the backend on Cloudflare

The backend is a Cloudflare Worker named `fergus-job`. The free plan
covers this volume. It keeps the record of actioned emails in a D1 database
(`server/migrations/`). Cloudflare runs as many copies as it needs, and a
lock in the database stops two people creating a job from the same email
at once.

One-time setup, in the Cloudflare dashboard (https://dash.cloudflare.com):

1. **Storage & Databases → D1 → Create database**, named `fergus-job`. Pick
   **Oceania** as its location if offered. Put its **database ID** in
   `server/wrangler.jsonc` (`database_id`).
2. **Workers & Pages** → note your **workers.dev subdomain** (Overview,
   right-hand side) and your **Account ID**. Put the account ID in
   `server/wrangler.jsonc` (`account_id`). The backend's URL is
   `https://fergus-job.<subdomain>.workers.dev`; put that in
   `taskpane/config.js` (`backendUrl`). Ours is
   https://fergus-job.martin-3f7.workers.dev.
3. **My Profile → API Tokens → Create Token →** the **Edit Cloudflare
   Workers** template, then **+ Add more → Account · D1 · Edit**. Limit it
   to your account and create it.
4. In GitHub (repo **Settings → Secrets and variables → Actions**), add the
   secret `CLOUDFLARE_API_TOKEN` = that token.
5. Push to `main` (or run **Actions → Deploy backend to Cloudflare → Run
   workflow**). It tests, creates the table, and deploys the Worker.
6. **Workers & Pages → fergus-job → Settings → Variables and Secrets**, add:

| Name | Type | Value |
| --- | --- | --- |
| `FERGUS_API_TOKEN` | Secret | Fergus personal access token (Fergus → Settings → Personal Access Tokens) |
| `ANTHROPIC_API_KEY` | Secret | Claude API key |
| `FERGUS_GROUPS` | Text | Comma-separated business-unit groups, named exactly as in Fergus |

Check https://fergus-job.martin-3f7.workers.dev/api/health shows
`{"ok":true}`.

The non-secret settings live in `server/wrangler.jsonc` under `vars`:
`ENTRA_CLIENT_ID` / `ENTRA_TENANT_ID` (from step 1 above),
`OWN_EMAIL_DOMAINS` (our staff domains, so a forwarder is never taken for
the customer), `FERGUS_DEFAULT_JOB_TYPE` (`Quote`, `Estimate` or `Charge
Up`) and `ALLOWED_ORIGINS`. Optional extras, set in the dashboard:
`FERGUS_JOB_URL_TEMPLATE` (job link, `{jobId}` / `{jobNo}` substituted;
default `https://app.fergus.com/jobs/view/{jobId}`, so check it against a
real job) and `ANTHROPIC_MODEL` (default `claude-opus-5-5`). Settings
added in the dashboard survive deploys (`keep_vars`).

### 3. Test it

**Automated tests** (no network or keys needed — Fergus and Claude are
faked):

```sh
cd server && npm test
```

**Backend against real Fergus and Claude, without Outlook.** Run the
Worker locally with sign-in turned off (`npm run dev` runs Cloudflare's
real Workers runtime, with a local copy of the database), then send it the
sample email in `server/samples/builder-email.json`, a builder asking for a
decking quote for their client:

```sh
cd server
npm ci
cp .dev.vars.example .dev.vars        # fill in the Fergus token, Claude key and groups
npx wrangler d1 migrations apply fergus-job --local
npm run dev

# in another terminal:
curl -s localhost:8787/api/analyze -H 'Content-Type: application/json' \
  -d @samples/builder-email.json
```

You should get `"status": "new"`, an `extraction` with customer **Aroha
Ngata** (not the builder, Sam), phone `021 555 0142`, site **14 Kowhai
Road, Mt Eden, Auckland 1024**, a suggested group, and any matching
Fergus customers/sites under `candidates`. `/api/analyze` only reads from
Fergus. To create a real job (which then exists in Fergus), POST to
`/api/jobs` with the email plus the form fields, e.g.:

```sh
curl -s localhost:8787/api/jobs -H 'Content-Type: application/json' -d '{
  "email": '"$(node -p 'JSON.stringify(require("./samples/builder-email.json").email)')"',
  "jobType": "Quote", "group": "",
  "title": "TEST – decking 14 Kowhai Rd", "description": "Test job from the add-in README",
  "customer": { "new": { "name": "Aroha Ngata", "contactName": "Aroha Ngata", "email": "aroha.ngata@example.com", "phone": "021 555 0142" } },
  "site": { "new": { "address": { "address1": "14 Kowhai Road", "suburb": "Mt Eden", "city": "Auckland", "postcode": "1024" }, "contactName": "Aroha Ngata", "phone": "021 555 0142" } }
}'
```

Run the same command again and you get `"status": "exists"` with the same
job, with no duplicate. Delete the test job, customer and site in Fergus
afterwards, and delete `server/.wrangler/` to reset the local database.
`.dev.vars` is git-ignored; never commit it.

**In Outlook.** With the backend deployed and `taskpane/config.js` filled
in, send yourself the sample (subject and body from
`samples/builder-email.json`), open it, and click **Create Fergus job**.
Check the form, create the job, and confirm the email gets the **Fergus job
created** category. Close and reopen the pane: it should show the existing
job instead of the form.

> **Outlook category:** add-ins can only apply a category that already
> exists in the mailbox's category list (creating one needs a much broader
> permission than this add-in asks for). Each person creates **Fergus job
> created** once — **Categorize → Manage categories → New** — or the pane
> says it couldn't tag the email. The job is created either way.
