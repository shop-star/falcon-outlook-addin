# Falcon Flooring Outlook add-in

An Outlook add-in that adds two buttons, under **Falcon Flooring** on the
ribbon, when you're reading an email:

- **Create Fergus job** — opens a task pane that reads the email, has
  Claude pull out the customer, contact, site address and job details
  (including from forwarded emails and builders writing on a client's
  behalf), finds matching customers and sites already in Fergus, and lets
  you check and edit everything before creating the job. The original email
  is attached to the job as a note and the email is tagged with the Outlook
  category **Fergus job created**. Details below.
- **Surface Takeoff** — opens the Surface Takeoff plan-measuring app in
  your browser, with the email's subject, PDF/zip attachment names and
  links carried over. It uses Outlook's `openBrowserWindow` where that
  works, falls back to a browser pop-up window (Outlook on the web), and
  otherwise shows a banner with the app's address. The app itself now lives in its own repo,
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

Microsoft's **Get Add-ins** button now opens the Microsoft 365 app store,
which has no way to add your own manifest. Use the direct link instead:

1. Open **Outlook on the web** (https://outlook.office.com) and stay signed in.
2. In another tab of the same browser, open **https://aka.ms/olksideload**.
   It brings up the **Add-ins for Outlook** dialog in Outlook.
3. Choose **My add-ins**, scroll to **Custom Addins**, remove any older
   copy of this add-in, then **Add a custom add-in → Add from file…** and
   pick `manifest.xml` (download it from
   https://shop-star.github.io/falcon-outlook-addin/manifest.xml).

Add-ins added in Outlook on the web also reach the new Outlook for Windows
and Outlook for Mac on the same account, usually within minutes; the
classic desktop apps may need a restart (see the Mac note below).

Once added, open any email and you'll see **Falcon Flooring** on the ribbon
with **Surface Takeoff** and **Create Fergus job** under it (in Outlook on
the web this may sit under **Apps** or the **…** menu).

> **Note:** some organizations restrict end-users from sideloading custom
> add-ins (a tenant admin setting). If "Add from file" is missing or
> blocked, your Microsoft 365 admin can either enable it for you or install
> the manifest for you centrally via the Microsoft 365 admin center
> (Settings → Integrated apps → Upload custom apps).
>
> **Outlook for Mac still showing an old button** (e.g. "Floor Area
> Takeoff"): Mac Outlook caches add-ins. Quit Outlook, delete the folder
> `~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef` (Finder →
> Go → Go to Folder…), and reopen Outlook. If the old one is still there,
> remove it from **Get Add-ins → My add-ins** (or the ⋯ menu → Get
> Add-ins) on the Mac.


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

## Public holidays on the Fergus calendar

The same Worker also keeps New Zealand public holidays on the Fergus
calendar a year ahead. A Cloudflare cron trigger (`triggers` in
`server/wrangler.jsonc`) runs it at 19:00 UTC on the 1st of each month,
which is early morning on the 2nd in New Zealand.

- **Which holidays:** the national ones (New Year, Waitangi, Easter,
  ANZAC, King's Birthday, Matariki, Labour Day, Christmas, Boxing Day)
  **and every region's anniversary day** (Wellington, Auckland, Canterbury,
  Otago and the rest). A regional one's description says which region it
  applies to. Each is entered on the day it's *observed* (a weekend holiday
  moves to the Monday or Tuesday). They're worked out in
  `server/src/holidays.js` from the Holidays Act rules, so no outside
  service is involved. Matariki's dates are set in law up to 2052.
- **How they look in Fergus:** an unassigned event (type `OTHER`, not
  tied to a job or a person), so everyone sees it, from 7am to 5pm New
  Zealand time, titled with the holiday's name.
- **No duplicates:** each holiday added is recorded in the D1 table
  `holiday_events`, so the first run adds about two dozen events and later
  runs add only what's come into the twelve-month window.
- **If Fergus refuses an event,** the run stops and that holiday is tried
  again next month.
- **Each run's outcome** (what was added, or Fergus's error) is shown at
  https://fergus-job.martin-3f7.workers.dev/api/holidays, along with every
  holiday added so far.
- **To run it now,** open
  https://fergus-job.martin-3f7.workers.dev/api/holidays/run in a browser.
  It shows what was added, or Fergus's error. Running it again only adds
  holidays not already added.

There's nothing to set up beyond the Worker itself: merging to `main`
deploys the cron trigger and creates the `holiday_events` table.

### Test it now

To run the sync now rather than waiting for the 1st, run the Worker
locally against real Fergus and trigger the cron:

```sh
cd server
npx wrangler d1 migrations apply fergus-job --local
npx wrangler dev --test-scheduled
# in another terminal:
curl "localhost:8787/__scheduled?cron=0+19+1+*+*"
```

This creates real events in Fergus. If Fergus rejects one, its message
appears in the `wrangler dev` output. A local run records holidays in the
local database only, so delete the test events in Fergus before the
deployed Worker's first run adds them again.
