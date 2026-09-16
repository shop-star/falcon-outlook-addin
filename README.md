# Floor Area Takeoff — Outlook add-in

Adds a "Floor Areas" button to the ribbon when you're reading an email.
Click it, pick a PDF plan attached to that email, and it opens in a
full-sized pop-up window — the task pane itself is too narrow to usefully
show a floor plan. In that window, click two points on a known measurement
(a scale bar or a labelled dimension) to calibrate, then click around each
room's outline. Results (area in m², with a running total) show live back
in the task pane as you trace, and you can copy them straight into a
spreadsheet or quote.

Everything runs in your browser/Outlook client — the PDF is never uploaded
anywhere. There's no subscription, no per-use cost, and no account to sign
up for.

## What you need

- A Microsoft 365 work/school account (this is what you said you're using).
- Somewhere to host four static files over HTTPS. **GitHub Pages** is the
  easiest free option and is what these instructions use, but any static
  HTTPS host works (Cloudflare Pages, Netlify, Azure Static Web Apps, your
  own web server, etc.) — the manifest just needs a stable HTTPS base URL.

## 1. Host the files

Using GitHub Pages (free):

1. Create a new **public** GitHub repository, e.g. `floor-area-addin`.
2. Upload everything in this folder (`manifest.xml`, `taskpane.html`,
   `taskpane.css`, `taskpane.js`, `viewer.html`, `viewer.css`, `viewer.js`,
   and the `assets/` and `vendor/` folders) to the repo root.
3. In the repo, go to **Settings → Pages**, set **Source** to the `main`
   branch, root folder, and save.
4. Wait a minute, then note the URL GitHub shows you, e.g.:
   `https://yourusername.github.io/floor-area-addin`
   That's your base URL — it should have **no trailing slash**.

If you'd rather use a different host, just make sure the same files
end up reachable at `<base-url>/taskpane.html`,
`<base-url>/taskpane.css`, `<base-url>/taskpane.js`,
`<base-url>/viewer.html`, `<base-url>/viewer.css`, `<base-url>/viewer.js`,
`<base-url>/assets/icon-16.png` (etc), and
`<base-url>/vendor/pdfjs/pdf.min.js` (and `pdf.worker.min.js`) — the PDF
viewer library is hosted alongside the add-in itself rather than pulled
from a public CDN, since some corporate networks block CDN domains from
inside the Outlook add-in sandbox.

## 2. Point the manifest at your hosted files

Open `manifest.xml` and replace every occurrence of the placeholder text

```
REPLACE-WITH-YOUR-BASE-URL
```

with your actual base URL from step 1 (no trailing slash), for example:

```
https://yourusername.github.io/floor-area-addin
```

There are 7 occurrences (IconUrl, HighResIconUrl, SupportUrl, AppDomain,
SourceLocation, and the three Resources URLs/images). A quick way to do this
on your own machine:

- **Mac/Linux (Terminal):**
  `sed -i '' 's|REPLACE-WITH-YOUR-BASE-URL|https://yourusername.github.io/floor-area-addin|g' manifest.xml`
- **Windows (PowerShell):**
  `(Get-Content manifest.xml) -replace 'REPLACE-WITH-YOUR-BASE-URL','https://yourusername.github.io/floor-area-addin' | Set-Content manifest.xml`
- Or just open it in any text editor and use Find & Replace.

Re-upload the edited `manifest.xml` to the same hosting location if you
already uploaded it in step 1 (or just edit it before uploading).

## 3. Sideload it into Outlook (work/school account)

**Outlook on the web** (works.outlook.com / outlook.office.com):

1. Open any email, then in the ribbon go to **Get Add-ins** (or the "..."
   menu → **Get Add-ins**).
2. Choose **My add-ins** (left side), scroll to **Custom Addins**, and
   click **Add a custom add-in → Add from file...**
3. Select your edited `manifest.xml`. Accept the warning about
   unverified/custom add-ins.

**New Outlook for Windows / Outlook on Mac:** same steps — **Get Add-ins**
→ **My add-ins** → **Add a custom add-in** → **Add from file**.

**Classic Outlook for Windows:** **File → Manage Add-ins** opens the same
web dialog above, or use **Home → Get Add-ins → My add-ins**.

Once added, open any email with a PDF plan attached and you'll see a
**Floor Areas** button in the ribbon.

> **Note:** some organizations restrict end-users from sideloading custom
> add-ins (a tenant admin setting). If "Add from file" is missing or
> blocked, your Microsoft 365 admin can either enable it for you or install
> the manifest for you centrally via the Microsoft 365 admin center
> (Settings → Integrated apps → Upload custom apps).

## Using it

1. Open the email with the plan attached, click **Floor Areas** in the
   ribbon.
2. Pick the PDF from the list and click **Open** — it opens in a separate,
   full-sized pop-up window.
3. In that window, click **Set scale**, then click the two ends of a
   labelled dimension or the plan's scale bar, and enter its real length
   and unit when prompted.
4. Click **Trace room**, click around a room's outline corner by corner,
   then **Finish room** and give it a name. Repeat for each room (and each
   page, if the plan has multiple floors — you'll need to set the scale
   once per page).
5. Back in the task pane, read the running list and total in the
   **Floor areas** section (it updates live as you trace), and use
   **Copy results** to paste a tab-separated list into Excel or your quote.
   If you close the pop-up early, **Reopen window** brings it back with your
   scale and traced rooms intact.

Your scale and traced rooms are saved automatically as you work — both for
the rest of the session and, longer-term, on the email itself (via
Outlook's own item storage, so no separate account or server is needed).
Reopening the same email later, even on a different device, brings your
progress back exactly where you left it. **Clear all** wipes this saved
progress too, so use it when you actually want to start over.

## Limitations to know about

- This is a **manual trace** tool, not automatic room detection — it won't
  guess wall lines for you. That trade-off is deliberate: automatic
  detection on real construction drawings (furniture, dimension lines,
  hatching) is unreliable, and for quoting you want areas you can trust.
- If an attachment is a OneDrive/SharePoint link rather than an actual
  attached file, reading its content may fail — download and re-attach the
  real PDF first.
- Multi-page plans need the scale set separately on each page (a ground
  floor and first floor page are usually drawn at different points on the
  sheet, so we don't assume they share a scale).
- If your organization has disabled add-in sideloading, you'll need your
  Microsoft 365 admin to deploy the manifest centrally (see note above).
