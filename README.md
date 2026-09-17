# Floor Area Takeoff

Two things live in this repo:

1. **A standalone browser app** (`app/`) — open a PDF plan (or a zip
   containing one), click two points on a known measurement to calibrate
   the scale, then trace each room's outline. Results (area in m², with a
   running total) update live, you can copy them straight into a
   spreadsheet or quote, and download an annotated copy of the PDF with the
   outlines burned onto the pages and the room data embedded in the file
   itself (so reopening that downloaded copy restores your exact progress —
   no account or server involved).
2. **A minimal Outlook add-in** (`manifest.xml`, `commands.html`,
   `commands.js`) — adds a "Floor Areas" button to the ribbon when you're
   reading an email. Clicking it opens the app above in a real browser
   window, with the open email's subject, PDF/zip attachment names, and any
   links found in its body carried over so you have a head start. If
   there's exactly one plan-like attachment, its bytes go on your clipboard
   too — the app picks that up from a paste, no manual save-then-drag
   needed for the common case of one plan per email.

Everything runs in your browser — a PDF is never uploaded anywhere. There's
no subscription, no per-use cost, and no account to sign up for.

## Why it's split this way

Outlook add-ins run inside an embedded browser control, and that embedded
control turned out to never reliably support triggering a file download —
several different approaches were tried and each failed identically (or
couldn't be confirmed working) across real Outlook clients. Doing the
actual plan viewing/tracing/export in an ordinary browser tab sidesteps
that entirely: downloads, drag-and-drop, and file pickers all just work.
The add-in's only job now is to hand off to that tab with a head start —
see `Office.context.ui.openBrowserWindow` in `commands.js`, an API
Microsoft ships specifically for "things the sandboxed add-in view can't
do, like downloading a file."

## Using the app on its own

Open `app/index.html` (or your hosted URL) in any browser. Drag a PDF plan
onto the page, or a zip file containing one (if it has more than one PDF
inside, you'll be asked which one to open) — no Outlook or email needed at
all. Everything else works the same as described below.

## What you need to host the add-in

- A Microsoft 365 work/school account.
- Somewhere to host static files over HTTPS. **GitHub Pages** is the
  easiest free option and is what's deployed here (see
  `.github/workflows/deploy-pages.yml`), but any static HTTPS host works.

## 1. Host the files

The GitHub Actions workflow in this repo deploys everything needed —
`manifest.xml`, `commands.html`, `commands.js`, `assets/`, `vendor/`, and
the whole `app/` folder — to GitHub Pages on every push to `main`. If
you're forking this for your own deployment, update every
`https://shop-star.github.io/falcon-outlook-addin` URL in `manifest.xml`
and `commands.js` (`APP_URL`) to your own base URL first.

## 2. Sideload the add-in into Outlook (work/school account)

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

Once added, open any email and you'll see a **Floor Areas** button in the
ribbon.

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

## Using it

1. Open the email with the plan attached, click **Floor Areas** in the
   ribbon — the app opens in a new browser tab, pre-filled with that
   email's subject, attachment names, and any links found in the body.
2. If there was exactly one PDF/zip attachment, it's already on your
   clipboard — click anywhere on the page and press Cmd/Ctrl+V to load it
   straight away. Otherwise (several attachments, or you followed a link
   instead), save the file from Outlook's own UI or the link, then drag it
   into the app (or a zip containing it).
3. Click **Set scale**, then click the two ends of a labelled dimension or
   the plan's scale bar, and enter its real length and unit when prompted.
4. Click **Trace room**, click around a room's outline corner by corner,
   then **Finish room** and give it a name. Repeat for each room (and each
   page, if the plan has multiple floors — you'll need to set the scale
   once per page). To fix up a room afterward, click its outline (when
   you're not actively tracing or setting scale) to select it, then drag
   its corner handles to reshape it — the area updates as you drag.
5. Read the running list and total in the **Floor areas** section (it
   updates live as you trace), and use **Copy results** to paste a
   tab-separated list into Excel or your quote, or **Download PDF with room
   data** for an annotated copy of the plan itself.

Your scale and traced rooms are saved automatically as you work, to your
browser's local storage for that exact file (by name and size) — reopening
the same file later in the same browser restores it. That's inherently
per-browser/per-device, unlike the old Outlook-item-storage approach — but
the **downloaded** PDF also carries its own progress embedded in it (see
above), so passing that file along (by email, OneDrive, wherever) is what
carries your progress across devices or to someone else. **Clear all**
wipes the saved progress for the current file, so use it when you actually
want to start over.

## Limitations to know about

- This is a **manual trace** tool, not automatic room detection — it won't
  guess wall lines for you. That trade-off is deliberate: automatic
  detection on real construction drawings (furniture, dimension lines,
  hatching) is unreliable, and for quoting you want areas you can trust.
- Multi-page plans need the scale set separately on each page (a ground
  floor and first floor page are usually drawn at different points on the
  sheet, so we don't assume they share a scale).
- The clipboard hand-off only ever covers **one** attachment (whichever is
  the sole plan-like attachment on the email) and skips anything over 8MB
  — ambiguous or oversized cases fall back to the plain name list, saved
  and dragged in manually. Clipboard write from inside the add-in's launch
  handler also depends on Outlook granting clipboard access at that point,
  which hasn't been battle-tested across every Outlook build — if the hint
  never shows up, the manual drag-and-drop path always works regardless.
- Email *links* never carry any bytes — there's no built-in Dropbox/OneDrive
  picker (yet), so links are just shown for a one-click open in your own
  browser, where your existing login to that service applies normally, and
  you still need to save the file from there and drop it into the app.
- If your organization has disabled add-in sideloading, you'll need your
  Microsoft 365 admin to deploy the manifest centrally (see note above).
