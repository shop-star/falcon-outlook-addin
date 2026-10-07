# Surface Takeoff

Two things live in this repo:

1. **A standalone browser app** (`app/`) — open a PDF plan (or a zip
   containing one), click two points on a known measurement to calibrate
   the scale, then trace each outline. Close a tracing as an **area** (its
   m²) or leave it open and **measure its length** (its m) instead — both
   kinds show up together in the results, with a running total for each.
   Results update live, you can copy them straight into a spreadsheet or
   quote, and download an annotated copy of the PDF with the tracings
   burned onto the pages and the underlying data embedded in the file
   itself (so reopening that downloaded copy restores your exact progress —
   no account or server involved).
2. **A minimal Outlook add-in** (`manifest.xml`, `commands.html`,
   `commands.js`) — adds a "Surface Takeoff" button to the ribbon when
   you're reading an email. Clicking it opens the app above in a real
   browser window, with the open email's subject, PDF/zip attachment names,
   and any links found in its body carried over so you have a head start.

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

### Plans that arrive as images

If a plan comes as photos or scans (JPG, PNG — also GIF, WebP or BMP)
instead of a PDF, pick or drop **one or more images at once** (or a zip
holding only images). The app combines them into a single new PDF, one
image per page, with pages ordered by file name (so `page 2` comes before
`page 10`). Phone photos are turned upright automatically. From then on
it's an ordinary PDF: set the scale, trace, and download as usual. The
downloaded copy carries your tracings like any other, and a refresh brings
it straight back. Picking the same images again later also restores your
saved progress.

Set the scale on an image by **measuring a known dimension**. A typed
scale like `1:100` only works if the image records its real scan
resolution (e.g. a 300 DPI scan). The app tells you when it doesn't.

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

Once added, open any email and you'll see a **Surface Takeoff** button in
the ribbon.

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

1. Open the email with the plan attached, click **Surface Takeoff** in the
   ribbon — the app opens in a new browser tab, pre-filled with that
   email's subject, attachment names, and any links found in the body.
2. Save the PDF attachment from Outlook's own UI (or download it from one
   of the listed links), then drag it into the app (or a zip containing
   it). Once a plan is open, section 1 also holds the page navigation
   (◀ ▶ and **Go to page**) and zoom controls, and shrinks down to its file
   name,
   a **Rename** button (the name is what your downloads are called, and
   your saved progress follows the file to its new name), and an **Open a
   different file** button.
3. In the **Set scale** section, either click the two ends of a labelled
   dimension or the plan's scale bar and enter its real length and unit
   when prompted, or — if you already know it — just type the page's
   printed scale (e.g. `1:100`) into the box below and click **Use this
   scale**, or click one of the **1:50 / 1:100 / 1:200** preset buttons for
   an instant one-click set. The typed-scale shortcut assumes the PDF's
   page size matches the real sheet size (i.e. it wasn't resized on
   export); if measurements come out wrong, use the click-and-measure
   method instead.
4. In **Add tracings**, click **Trace**, then click around an outline
   point by point. When you're done, either:
   - Click **Finish area** (needs at least 3 points) to close it as an
     area — its m² is computed from the enclosed shape, or
   - Click **Measure length** (needs at least 2 points) to leave it open
     and just measure the total length of the lines you drew, without
     closing it into a shape.

   Either way, it's added straight to the **Takeoff results** pane with
   the cursor already in its name box — type a name and press Enter (or
   carry on and name it later). Each one gets a different colour, which
   you can change from its swatch in the results pane. Repeat for each
   tracing (and each page, if the plan has multiple floors — you'll need
   to set the scale once per page). To fix up a tracing afterward, click
   its outline/line (when you're not actively tracing or setting scale) to
   select it, then drag its point handles to reshape it — its measurement
   updates as you drag — or drag the circle on its label to reposition the
   label itself (even away from the tracing, which draws a thin line back
   to it so it's still clearly that tracing's label). Both a tracing's
   drawing on the plan and its label always show its measurement (e.g.
   "Bedroom — 16.59 m²" or "Wall run A — 4.70 m"), and you can change a
   tracing's name or colour anytime directly in the results table.
5. Read the running list and totals in the **Takeoff results** pane on the
   right (it updates live as you trace, ordered by page, and also shows
   each tracing's page and the scale that was set on it) — click anywhere
   on a tracing's row to jump straight to its page in the viewer and
   select it for editing, or use the **Go to page** dropdown to jump to
   any page directly. Drag a row by its ⠿ handle to reorder the list (or
   click the handle and use the arrow keys) — new tracings are added
   after the others on their page, and **Copy results** and the PDF
   summary page follow whatever order you set. Areas and lengths are
   totalled separately (only
   shown when you actually have one of that kind). Use **Copy results**
   to paste a tab-separated list into Excel or your quote.
6. The **Download** section is always there, no button to open it first —
   pick which pages to include (just the current page, the whole
   document, or only the pages with a tracing on them) and whether to
   append a summary page listing every tracing and its measurement (on by
   default), then click **Download**.

If the numbered steps feel cramped stacked on top of each other, the
**Layout** control at the top of the page switches to putting all four
numbered sections in a narrow column on the left, with the plan
itself filling the middle — the **Takeoff results** pane always stays on
the right, at the same width, in either case. Pick
whichever suits your screen; it's remembered next time you open the app.

Your scale and tracings are saved automatically as you work, to your
browser's local storage for that exact file (by name and size), and the
open file itself is also kept (in IndexedDB, which allows much more room
than local storage) — so simply refreshing the page, or accidentally
closing and reopening the tab, brings back exactly where you left off with
no need to re-pick the file. That's inherently per-browser/per-device,
unlike the old Outlook-item-storage approach — but the **downloaded** PDF
also carries its own progress embedded in it (see above), so passing that
file along (by email, OneDrive, wherever) is what carries your progress
across devices or to someone else. **Clear all** wipes the saved
tracing/scale progress for the current file (not the remembered-open-file
part), so use it when you actually want to start over.

## Limitations to know about

- This is a **manual trace** tool, not automatic detection — it won't
  guess wall lines for you. That trade-off is deliberate: automatic
  detection on real construction drawings (furniture, dimension lines,
  hatching) is unreliable, and for quoting you want measurements you can
  trust.
- Multi-page plans need the scale set separately on each page (a ground
  floor and first floor page are usually drawn at different points on the
  sheet, so we don't assume they share a scale).
- The add-in only ever passes along attachment *names* and email *links* —
  never file bytes — so you still save/download the actual PDF yourself
  and drop it into the app. There's no built-in Dropbox/OneDrive picker
  (yet); links are just shown for a one-click open in your own browser,
  where your existing login to that service applies normally.
- If your organization has disabled add-in sideloading, you'll need your
  Microsoft 365 admin to deploy the manifest centrally (see note above).
