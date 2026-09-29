/* Surface Takeoff — standalone browser app
 * Replaces the old Outlook task pane + pop-up dialog pair with a single
 * ordinary page: no Office.js, no embedded WebView, no cross-window
 * messaging. A real browser tab handles file downloads and drag-and-drop
 * natively, which is the whole reason this exists (see README) — the
 * Outlook add-in's embedded dialog WebView never reliably supported either.
 *
 * Ingestion is a plain PDF or ZIP file, picked or dragged in. A companion
 * Outlook add-in (see manifest.xml) can open this page pre-filled with the
 * open email's subject, attachment names, and any links found in its body,
 * via ?subject=&attachment=&link= query params — see initFromEmailPanel.
 */

// ---- Lazy-loaded, self-hosted libraries ------------------------------------
// Same reasoning throughout: self-hosted rather than pulled from a public
// CDN, since some corporate networks block CDN domains; loaded on demand
// rather than eagerly, since most sessions only ever need pdf.js.
function loadScript(src, globalName) {
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = src;
    script.onload = () => {
      if (!window[globalName]) {
        reject(new Error(`${src} loaded but window.${globalName} was not set`));
        return;
      }
      resolve(window[globalName]);
    };
    script.onerror = () => reject(new Error(`Failed to load ${src}`));
    document.head.appendChild(script);
  });
}

let pdfjsLib = null;
let pdfjsLoadPromise = null;
function loadPdfJs() {
  if (!pdfjsLoadPromise) {
    pdfjsLoadPromise = loadScript("../vendor/pdfjs/pdf.min.js?v=__CACHEBUST__", "pdfjsLib")
      .then((lib) => {
        lib.GlobalWorkerOptions.workerSrc = "../vendor/pdfjs/pdf.worker.min.js?v=__CACHEBUST__";
        pdfjsLib = lib;
        return lib;
      })
      .catch((err) => {
        pdfjsLoadPromise = null;
        throw err;
      });
  }
  return pdfjsLoadPromise;
}

let pdfLibLoadPromise = null;
function loadPdfLib() {
  if (!pdfLibLoadPromise) {
    pdfLibLoadPromise = loadScript("../vendor/pdf-lib/pdf-lib.min.js?v=__CACHEBUST__", "PDFLib").catch(
      (err) => {
        pdfLibLoadPromise = null;
        throw err;
      }
    );
  }
  return pdfLibLoadPromise;
}

let jsZipLoadPromise = null;
function loadJSZip() {
  if (!jsZipLoadPromise) {
    jsZipLoadPromise = loadScript("../vendor/jszip/jszip.min.js?v=__CACHEBUST__", "JSZip").catch((err) => {
      jsZipLoadPromise = null;
      throw err;
    });
  }
  return jsZipLoadPromise;
}

const UNIT_TO_M = { mm: 0.001, cm: 0.01, m: 1, ft: 0.3048, in: 0.0254 };

// A tracing's colour is picked from this palette by default (cycling by how
// many tracings already exist), and can always be overridden — either in
// the naming form when first drawing it, or from the swatch in the results
// table afterward.
const DEFAULT_TRACING_COLOR = "#15655c";
const TRACING_COLOR_PALETTE = [
  "#15655c",
  "#c2185b",
  "#e07b00",
  "#3f51b5",
  "#7b1fa2",
  "#00838f",
  "#558b2f",
  "#ad1457",
];
function paletteColor(index) {
  return TRACING_COLOR_PALETTE[index % TRACING_COLOR_PALETTE.length];
}

// ---- State -----------------------------------------------------------
let currentPdf = null;
let currentPageNum = 1;
let renderScale = 1.5;
let currentFileName = null;
// Kept separately from whatever pdf.js does with its own copy of the bytes
// (getDocument() can transfer/detach the buffer it's given) so the download
// button always has a pristine, untouched original to build from.
let originalBytes = null;

/** pageGeometry[pageNum] = { calibration: {p1,p2,metersPerUnit,label} | null, tracings: [{id,name,points,kind}] } */
let pageGeometry = {};

/** flat list mirrored to the results table. kind: "area" (has areaM2) or "length" (has lengthM) */
let tracings = []; // { id, page, name, color, kind, areaM2? , lengthM? }
let nextTracingId = 1;

let mode = "idle"; // idle | calibrate | trace
let calibTemp = { p1: null, p2: null };
let traceTemp = { page: null, points: [] };
// Set right before opening the naming form (by the "Finish area" or
// "Measure length" button), so the confirm handler knows which kind of
// tracing it's actually completing.
let pendingTraceKind = null; // "area" | "length"

// Clicking an already-drawn tracing's outline while idle selects it for
// editing — its points are drawn as draggable handles until you click
// elsewhere (or start calibrating/tracing) to deselect.
let editingTracing = null; // tracing id, or null
let vertexDragState = null; // { tracingId, pointIndex, dragged }
let labelDragState = null; // { tracingId, dragged }

// ---- DOM refs ----------------------------------------------------------
const statusBarEl = document.getElementById("statusBar");
const fileNameHeadingEl = document.getElementById("fileNameHeading");

const mainLayoutEl = document.getElementById("mainLayout");
const layoutStackedBtn = document.getElementById("layoutStackedBtn");
const layoutSidebarBtn = document.getElementById("layoutSidebarBtn");

const intakeSectionEl = document.getElementById("intakeSection");
const dropZone = document.getElementById("dropZone");
const pickFileBtn = document.getElementById("pickFileBtn");
const fileInput = document.getElementById("fileInput");
const zipPicker = document.getElementById("zipPicker");
const zipPickerList = document.getElementById("zipPickerList");
const currentFileInfo = document.getElementById("currentFileInfo");
const currentFileNameEl = document.getElementById("currentFileNameEl");
const changeFileBtn = document.getElementById("changeFileBtn");

const scaleSectionEl = document.getElementById("scaleSection");
const viewerSectionEl = document.getElementById("viewerSection");
const resultsPaneEl = document.getElementById("resultsPane");

const canvasScroller = document.getElementById("canvasScroller");
const pdfCanvas = document.getElementById("pdfCanvas");
const overlayCanvas = document.getElementById("overlayCanvas");
const pdfCtx = pdfCanvas.getContext("2d");
const overlayCtx = overlayCanvas.getContext("2d");

const prevPageBtn = document.getElementById("prevPageBtn");
const nextPageBtn = document.getElementById("nextPageBtn");
const pageIndicatorEl = document.getElementById("pageIndicator");
const zoomOutBtn = document.getElementById("zoomOutBtn");
const zoomInBtn = document.getElementById("zoomInBtn");
const zoomIndicatorEl = document.getElementById("zoomIndicator");

const setScaleBtn = document.getElementById("setScaleBtn");
const cancelScaleBtn = document.getElementById("cancelScaleBtn");
const traceBtn = document.getElementById("traceBtn");
const undoPointBtn = document.getElementById("undoPointBtn");
const finishAreaBtn = document.getElementById("finishAreaBtn");
const measureLengthBtn = document.getElementById("measureLengthBtn");
const cancelTraceBtn = document.getElementById("cancelTraceBtn");

const scaleRatioForm = document.getElementById("scaleRatioForm");
const scaleRatioInput = document.getElementById("scaleRatioInput");
const scaleRatioConfirmBtn = document.getElementById("scaleRatioConfirmBtn");

const calibrationForm = document.getElementById("calibrationForm");
const calibLengthInput = document.getElementById("calibLengthInput");
const calibUnitSelect = document.getElementById("calibUnitSelect");
const calibConfirmBtn = document.getElementById("calibConfirmBtn");
const calibCancelBtn = document.getElementById("calibCancelBtn");

const tracingNameForm = document.getElementById("tracingNameForm");
const tracingNameInput = document.getElementById("tracingNameInput");
const tracingColorInput = document.getElementById("tracingColorInput");
const tracingNameConfirmBtn = document.getElementById("tracingNameConfirmBtn");
const tracingNameCancelBtn = document.getElementById("tracingNameCancelBtn");

const scaleInfoEl = document.getElementById("scaleInfo");

const resultsPageSelect = document.getElementById("resultsPageSelect");
const resultsBody = document.getElementById("resultsBody");
const totalMeasurementEl = document.getElementById("totalMeasurement");
const copyResultsBtn = document.getElementById("copyResultsBtn");
const downloadSectionEl = document.getElementById("downloadSection");
const downloadIncludeSummary = document.getElementById("downloadIncludeSummary");
const downloadConfirmBtn = document.getElementById("downloadConfirmBtn");
const clearAllBtn = document.getElementById("clearAllBtn");
const copyFallback = document.getElementById("copyFallback");

const fromEmailPanel = document.getElementById("fromEmailPanel");
const fromEmailSubjectEl = document.getElementById("fromEmailSubject");
const fromEmailAttachmentsEl = document.getElementById("fromEmailAttachments");
const fromEmailLinksEl = document.getElementById("fromEmailLinks");

function setStatus(msg, isError) {
  statusBarEl.textContent = msg || "";
  statusBarEl.style.color = isError ? "#b3261e" : "";
}

function describeError(err) {
  if (!err) return "unknown error";
  const name = err.name || "Error";
  const message = err.message || String(err);
  return `${name}: ${message}`;
}

function isBenignRenderRace(text) {
  return typeof text === "string" && text.includes("Cannot use the same canvas");
}

window.addEventListener("error", (evt) => {
  const message = evt.message || String(evt.error);
  if (isBenignRenderRace(message)) return;
  setStatus("Something went wrong: " + message, true);
});
window.addEventListener("unhandledrejection", (evt) => {
  const reason = evt.reason && evt.reason.message ? evt.reason.message : evt.reason;
  if (isBenignRenderRace(String(reason))) {
    evt.preventDefault();
    return;
  }
  setStatus("Something went wrong: " + reason, true);
});

// ---- "From your email" panel, filled in by the launcher add-in ------------
// Everything here comes from an email someone else could have sent, so it's
// rendered as text (never innerHTML with the raw value) and link hrefs are
// restricted to http(s) — a crafted "link" query param is not a good enough
// reason to let a javascript: URI onto a real, clickable anchor.
function initFromEmailPanel() {
  const params = new URLSearchParams(window.location.search);
  const subject = params.get("subject");
  const attachments = params.getAll("attachment").filter(Boolean);
  const links = params.getAll("link").filter((u) => /^https?:\/\//i.test(u));

  if (!subject && attachments.length === 0 && links.length === 0) return;
  fromEmailPanel.hidden = false;

  if (subject) {
    fromEmailSubjectEl.textContent = `Subject: ${subject}`;
    fromEmailSubjectEl.hidden = false;
  }

  if (attachments.length > 0) {
    const intro = document.createElement("p");
    intro.className = "muted";
    intro.textContent = "Attachments on that email — save them from Outlook, then drop them below:";
    const ul = document.createElement("ul");
    attachments.forEach((name) => {
      const li = document.createElement("li");
      li.textContent = name;
      ul.appendChild(li);
    });
    fromEmailAttachmentsEl.append(intro, ul);
  }

  if (links.length > 0) {
    const intro = document.createElement("p");
    intro.className = "muted";
    intro.textContent = "Links found in that email:";
    const ul = document.createElement("ul");
    links.forEach((url) => {
      const li = document.createElement("li");
      const a = document.createElement("a");
      a.href = url;
      a.target = "_blank";
      a.rel = "noopener noreferrer";
      a.textContent = url;
      li.appendChild(a);
      ul.appendChild(li);
    });
    fromEmailLinksEl.append(intro, ul);
  }
}
initFromEmailPanel();

// ---- Full-file persistence (survives a page refresh) -----------------------
// The autosave below (see saveAutosave) only ever held the scale/tracing
// *geometry* — small enough for localStorage, keyed by file name+size. The
// PDF's actual bytes were never persisted, so refreshing the page always
// meant re-picking the file from disk, even though the tracings would have
// come right back once you did. IndexedDB has a much higher storage quota
// than localStorage (built for exactly this kind of blob), so the
// currently-open file's bytes are stashed there too, and checked for on
// startup — turning a refresh into a full, automatic restore.
//
// These storage keys/names keep their original "floorAreaTakeoff" spelling
// even after the app's rename to Surface Takeoff — they're invisible
// implementation details, and changing them would silently orphan anyone's
// already-saved file/progress instead of restoring it.
const FILE_DB_NAME = "floorAreaTakeoffFiles";
const FILE_STORE_NAME = "currentFile";
const FILE_DB_KEY = "current";

function openFileDb() {
  return new Promise((resolve, reject) => {
    let req;
    try {
      req = indexedDB.open(FILE_DB_NAME, 1);
    } catch (e) {
      reject(e);
      return;
    }
    req.onupgradeneeded = () => {
      req.result.createObjectStore(FILE_STORE_NAME);
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

// Best-effort: if IndexedDB isn't available (private browsing, quota,
// disabled storage), a refresh just falls back to today's behavior —
// nothing here needs to succeed for the app to keep working.
function persistCurrentFile(fileName, bytes) {
  openFileDb()
    .then((db) => {
      db.transaction(FILE_STORE_NAME, "readwrite").objectStore(FILE_STORE_NAME).put({ fileName, bytes }, FILE_DB_KEY);
    })
    .catch(() => {});
}

function loadPersistedFile() {
  return openFileDb().then(
    (db) =>
      new Promise((resolve, reject) => {
        const req = db.transaction(FILE_STORE_NAME, "readonly").objectStore(FILE_STORE_NAME).get(FILE_DB_KEY);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => reject(req.error);
      })
  );
}

loadPersistedFile()
  .then((entry) => {
    // If the user has already opened something (dragged/picked a file
    // themselves) by the time this IndexedDB read resolves, don't clobber
    // it — this check is what keeps a slow auto-restore from racing a
    // manual file pick that happens right after a page load.
    if (currentPdf) return;
    if (!entry || !entry.fileName || !entry.bytes) return;
    const bytes = entry.bytes instanceof Uint8Array ? entry.bytes : new Uint8Array(entry.bytes);
    openPdfFromBytes(entry.fileName, bytes);
  })
  .catch(() => {});

// ---- Layout toggle: stacked (default) vs. "Columns" (1, 2 & 4 narrowed) ---
const LAYOUT_STORAGE_KEY = "floorAreaTakeoff:layout";

function applyLayout(layout) {
  mainLayoutEl.classList.toggle("layout-sidebar", layout === "sidebar");
  layoutStackedBtn.classList.toggle("active", layout !== "sidebar");
  layoutSidebarBtn.classList.toggle("active", layout === "sidebar");
}

function setLayout(layout) {
  applyLayout(layout);
  try {
    localStorage.setItem(LAYOUT_STORAGE_KEY, layout);
  } catch (e) {
    // Storage unavailable — the choice just won't persist past this session.
  }
}

layoutStackedBtn.addEventListener("click", () => setLayout("stacked"));
layoutSidebarBtn.addEventListener("click", () => setLayout("sidebar"));

(function initLayout() {
  let saved = null;
  try {
    saved = localStorage.getItem(LAYOUT_STORAGE_KEY);
  } catch (e) {
    // Ignore — falls back to the default below.
  }
  applyLayout(saved === "sidebar" ? "sidebar" : "stacked");
})();

// ---- 1. File intake: drag-and-drop, file picker, and zip extraction -------
["dragenter", "dragover"].forEach((evtName) =>
  dropZone.addEventListener(evtName, (evt) => {
    evt.preventDefault();
    dropZone.classList.add("drag-over");
  })
);
["dragleave", "drop"].forEach((evtName) =>
  dropZone.addEventListener(evtName, (evt) => {
    evt.preventDefault();
    dropZone.classList.remove("drag-over");
  })
);
dropZone.addEventListener("drop", (evt) => {
  const file = evt.dataTransfer.files && evt.dataTransfer.files[0];
  if (file) handleFile(file);
});
pickFileBtn.addEventListener("click", () => fileInput.click());
changeFileBtn.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", () => {
  const file = fileInput.files && fileInput.files[0];
  if (file) handleFile(file);
  fileInput.value = "";
});

function handleFile(file) {
  const name = (file.name || "").toLowerCase();
  zipPicker.hidden = true;
  if (name.endsWith(".zip") || file.type === "application/zip") {
    handleZipFile(file);
  } else if (name.endsWith(".pdf") || file.type === "application/pdf") {
    file
      .arrayBuffer()
      .then((buf) => openPdfFromBytes(file.name, new Uint8Array(buf)))
      .catch((e) => setStatus(`Couldn't read "${file.name}" — ` + describeError(e), true));
  } else {
    setStatus(`"${file.name}" doesn't look like a PDF or a ZIP file.`, true);
  }
}

function handleZipFile(file) {
  setStatus(`Reading "${file.name}"…`);
  loadJSZip()
    .then((JSZipLib) => file.arrayBuffer().then((buf) => JSZipLib.loadAsync(buf)))
    .then(
      (zip) => {
        const pdfEntries = Object.values(zip.files).filter(
          (f) => !f.dir && /\.pdf$/i.test(f.name)
        );
        if (pdfEntries.length === 0) {
          setStatus(`No PDF files found inside "${file.name}".`, true);
          return;
        }
        if (pdfEntries.length === 1) {
          loadZipEntry(pdfEntries[0]);
          return;
        }
        showZipPicker(pdfEntries);
      },
      (err) => setStatus(`Couldn't read "${file.name}" as a zip — ` + describeError(err), true)
    );
}

function loadZipEntry(entry) {
  setStatus(`Extracting "${entry.name}"…`);
  entry.async("uint8array").then(
    (bytes) => openPdfFromBytes(entry.name.split("/").pop(), bytes),
    (err) => setStatus(`Couldn't extract "${entry.name}" — ` + describeError(err), true)
  );
}

function showZipPicker(entries) {
  zipPickerList.innerHTML = "";
  entries.forEach((entry) => {
    const btn = document.createElement("button");
    btn.textContent = entry.name;
    btn.addEventListener("click", () => {
      zipPicker.hidden = true;
      loadZipEntry(entry);
    });
    zipPickerList.appendChild(btn);
  });
  zipPicker.hidden = false;
  setStatus("That zip has more than one PDF — pick which one to open.");
}

// ---- 2. Opening a PDF -------------------------------------------------------
// Normally only ever triggered by one exclusive user action at a time (drag,
// file picker, zip-entry pick) — but the IndexedDB auto-restore above can
// also call this on startup, independently of the user, which makes genuine
// overlap possible for the first time (e.g. auto-restore is still loading
// an old file when the user drags in a new one right after the page loads).
// openRequestToken makes only the *latest* call's result ever get applied,
// regardless of which one's async chain happens to finish last.
let openRequestToken = 0;
function openPdfFromBytes(fileName, bytes) {
  const myToken = ++openRequestToken;
  setStatus(`Loading "${fileName}"…`);
  loadPdfJs().then(
    () => {
      if (myToken !== openRequestToken) return;
      // A copy for pdf.js — it can transfer/detach the buffer it's given,
      // and `bytes` needs to stay pristine for the download button to embed
      // later (see the doc comment on originalBytes above).
      pdfjsLib.getDocument({ data: bytes.slice() }).promise.then(
        (pdf) => {
          if (myToken !== openRequestToken) return;
          originalBytes = bytes;
          currentFileName = fileName;
          currentPdf = pdf;
          currentPageNum = 1;
          pageGeometry = {};
          tracings = [];
          nextTracingId = 1;
          renderScale = 1.5;
          resetToolState();
          setScaleBtn.disabled = false;
          fileNameHeadingEl.textContent = `3. Add tracings — "${fileName}"`;
          scaleSectionEl.hidden = false;
          viewerSectionEl.hidden = false;
          resultsPaneEl.hidden = false;
          downloadSectionEl.hidden = false;
          dropZone.hidden = true;
          currentFileInfo.hidden = false;
          currentFileNameEl.textContent = fileName;
          populatePageSelect(pdf.numPages);
          renderPage();
          updateResultsTable();
          setStatus(`Loaded "${fileName}". Set the scale, then trace each outline.`);
          restorePreviousProgress(pdf, fileName, bytes.length);
          persistCurrentFile(fileName, bytes);
        },
        (err) => {
          if (myToken !== openRequestToken) return;
          setStatus(`Couldn't open "${fileName}" — ` + describeError(err), true);
        }
      );
    },
    (err) => {
      if (myToken !== openRequestToken) return;
      setStatus("Couldn't load the PDF viewer library — " + describeError(err), true);
    }
  );
}

// Two sources of previously-saved progress, checked in order: this exact
// file already opened once in this browser (autosaved to localStorage,
// keyed by name+size — see saveAutosave), or the file itself is a
// previously-downloaded round-trip copy with its own embedded scale/tracing
// data (see buildDownloadPdfBytes). Either way, restoring means the flat
// tracings list gets recomputed from the geometry rather than trusted as
// saved, so it can never drift from what's actually drawn.
function restorePreviousProgress(pdf, fileName, byteLength) {
  const autosaved = loadAutosave(fileName, byteLength);
  if (autosaved && autosaved.pageGeometry && Object.keys(autosaved.pageGeometry).length > 0) {
    applyRestoredGeometry(autosaved.pageGeometry, autosaved.nextTracingId || autosaved.nextRoomId);
    setStatus("Restored your previous scale and tracings for this plan.");
    return;
  }
  pdf.getAttachments().then(
    (attachments) => {
      const embedded = attachments && attachments["floor-area-takeoff.json"];
      if (!embedded || !embedded.content) return;
      try {
        const parsed = JSON.parse(new TextDecoder().decode(embedded.content));
        if (parsed && parsed.pageGeometry && Object.keys(parsed.pageGeometry).length > 0) {
          applyRestoredGeometry(parsed.pageGeometry, parsed.nextTracingId || parsed.nextRoomId);
          setStatus("Restored the scale and tracings embedded in this PDF.");
        }
      } catch (e) {
        // Not our own embedded data, or corrupted — ignore, the PDF still
        // opened normally either way.
      }
    },
    () => {}
  );
}

function applyRestoredGeometry(geometry, nextId) {
  pageGeometry = geometry || {};
  nextTracingId = nextId || 1;
  tracings = [];
  editingTracing = null;
  vertexDragState = null;
  Object.keys(pageGeometry).forEach((pageNumKey) => {
    const pageNum = Number(pageNumKey);
    const geo = pageGeometry[pageNum];
    // Older saved/downloaded files used `rooms` instead of `tracings`, and
    // never had a `kind` (everything was an area) — normalize both here so
    // restoring an old file behaves exactly as it always did.
    if (!geo.tracings && geo.rooms) {
      geo.tracings = geo.rooms;
      delete geo.rooms;
    }
    geo.tracings = geo.tracings || [];
    geo.tracings.forEach((t) => {
      if (!t.kind) t.kind = "area";
      const color = t.color || DEFAULT_TRACING_COLOR;
      if (t.kind === "length") {
        const lengthPageUnits = polylineLengthPageUnits(t.points);
        const lengthM = geo.calibration ? lengthPageUnits * geo.calibration.metersPerUnit : 0;
        tracings.push({ id: t.id, page: pageNum, name: t.name, color, kind: "length", lengthM });
      } else {
        const areaPageUnits = polygonAreaPageUnits(t.points);
        const areaM2 = geo.calibration
          ? areaPageUnits * geo.calibration.metersPerUnit * geo.calibration.metersPerUnit
          : 0;
        tracings.push({ id: t.id, page: pageNum, name: t.name, color, kind: "area", areaM2 });
      }
    });
  });
  redrawOverlay();
  updateResultsTable();
}

// ---- Per-file autosave (localStorage) --------------------------------------
// A real browser tab keeps localStorage reliably, unlike the Outlook
// add-in's cross-window messaging — this is just "save progress for this
// exact file", keyed on name+size as a cheap fingerprint.
function autosaveKey(fileName, byteLength) {
  return `floorAreaTakeoff:${fileName}:${byteLength}`;
}

function loadAutosave(fileName, byteLength) {
  try {
    const raw = localStorage.getItem(autosaveKey(fileName, byteLength));
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

let autosaveTimer = null;
function writeAutosaveNow() {
  if (!currentFileName || !originalBytes) return;
  try {
    localStorage.setItem(
      autosaveKey(currentFileName, originalBytes.length),
      JSON.stringify({ pageGeometry, nextTracingId })
    );
  } catch (e) {
    // Storage full or unavailable (private browsing etc.) — this session
    // still works fine, it just won't be there next time.
  }
}

function saveAutosave() {
  if (!currentFileName || !originalBytes) return;
  clearTimeout(autosaveTimer);
  autosaveTimer = setTimeout(writeAutosaveNow, 500);
}

// The 500ms debounce above is fine for normal typing/dragging, but a reload
// or tab close right after the last edit would cancel that pending timer
// (page navigation kills JS timers) and silently lose it — exactly the kind
// of loss the refresh-persistence feature above is trying to prevent. Flush
// immediately, synchronously, the moment the page might be going away.
function flushAutosave() {
  if (!autosaveTimer) return;
  clearTimeout(autosaveTimer);
  autosaveTimer = null;
  writeAutosaveNow();
}
window.addEventListener("pagehide", flushAutosave);
window.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "hidden") flushAutosave();
});

function clearAutosave() {
  if (!currentFileName || !originalBytes) return;
  try {
    localStorage.removeItem(autosaveKey(currentFileName, originalBytes.length));
  } catch (e) {
    // Nothing more to do.
  }
}

// ---- Rendering --------------------------------------------------------
// A mouse-wheel zoom gesture fires many events in quick succession, each
// calling renderPage() — pdf.js refuses to start a render() on a canvas
// that still has one in flight ("Cannot use the same canvas during
// multiple render() operations"), so the previous render task is
// cancelled first. A cancelled render's promise rejects with
// RenderingCancelledException, which is expected/harmless here, not a
// real error to report.
let currentRenderTask = null;

// onResized, if given, runs right after the canvas is resized to the new
// scale but before the (async) redraw — used by wheel-zoom to restore the
// scroll position so the point under the cursor stays put.
function renderPage(onResized) {
  if (!currentPdf) return;
  currentPdf.getPage(currentPageNum).then((page) => {
    const viewport = page.getViewport({ scale: renderScale });
    pdfCanvas.width = overlayCanvas.width = Math.ceil(viewport.width);
    pdfCanvas.height = overlayCanvas.height = Math.ceil(viewport.height);

    if (onResized) onResized();

    if (currentRenderTask) {
      currentRenderTask.cancel();
    }
    currentRenderTask = page.render({ canvasContext: pdfCtx, viewport });
    currentRenderTask.promise.then(
      () => {
        currentRenderTask = null;
        redrawOverlay();
      },
      (err) => {
        currentRenderTask = null;
        if (err && err.name === "RenderingCancelledException") return;
        if (isBenignRenderRace(describeError(err))) return;
        setStatus("Couldn't render this page — " + describeError(err), true);
      }
    );

    pageIndicatorEl.textContent = `Page ${currentPageNum} / ${currentPdf.numPages}`;
    resultsPageSelect.value = String(currentPageNum);
    zoomIndicatorEl.textContent = `${Math.round((renderScale / 1.5) * 100)}%`;
    prevPageBtn.disabled = currentPageNum <= 1;
    nextPageBtn.disabled = currentPageNum >= currentPdf.numPages;
  });
}

// Populated once per opened file (numPages is fixed for that file) so
// section 3's "Go to page" selector always lists every real page, not just
// the ones with a tracing on them.
function populatePageSelect(numPages) {
  resultsPageSelect.innerHTML = "";
  for (let i = 1; i <= numPages; i++) {
    const opt = document.createElement("option");
    opt.value = String(i);
    opt.textContent = `Page ${i}`;
    resultsPageSelect.appendChild(opt);
  }
}

resultsPageSelect.addEventListener("change", () => {
  const pageNum = Number(resultsPageSelect.value);
  if (!currentPdf || pageNum === currentPageNum) return;
  resetToolState();
  currentPageNum = pageNum;
  renderPage();
});

function currentGeometry() {
  if (!pageGeometry[currentPageNum]) {
    pageGeometry[currentPageNum] = { calibration: null, tracings: [] };
  }
  return pageGeometry[currentPageNum];
}

function toCanvas([x, y]) {
  return [x * renderScale, y * renderScale];
}

function redrawOverlay() {
  overlayCtx.clearRect(0, 0, overlayCanvas.width, overlayCanvas.height);
  const geo = currentGeometry();

  // A scale entered directly as a ratio (see scaleRatioConfirmBtn) has no
  // drawn line to show — there's nothing clicked to draw.
  if (geo.calibration && geo.calibration.p1 && geo.calibration.p2) {
    drawLine(geo.calibration.p1, geo.calibration.p2, "#e07b00", 2, true, geo.calibration.label);
  }
  if (mode === "calibrate" && calibTemp.p1) {
    if (calibTemp.p2) {
      drawLine(calibTemp.p1, calibTemp.p2, "#e07b00", 2, true);
    } else {
      drawPoint(calibTemp.p1, "#e07b00");
    }
  }

  geo.tracings.forEach((tracing) => {
    const isEditing = tracing.id === editingTracing;
    const color = tracing.color || DEFAULT_TRACING_COLOR;
    const label = tracingLabelText(tracing);
    if (tracing.kind === "length") {
      drawMeasuredLine(tracing.points, color, label, isEditing ? 3 : 2, tracing.labelPos);
    } else {
      drawPolygon(tracing.points, color, label, isEditing ? 3 : 2, tracing.labelPos);
    }
    if (isEditing) {
      tracing.points.forEach((pt) => drawHandle(pt, color));
      drawLabelHandle(tracing.labelPos || polygonCentroid(tracing.points), color);
    }
  });

  if (mode === "trace" && traceTemp.page === currentPageNum && traceTemp.points.length > 0) {
    drawPolyline(traceTemp.points, "#c2185b");
  }

  traceBtn.disabled = !geo.calibration;
  setScaleBtn.textContent = geo.calibration ? "Re-set scale" : "Set scale";
  if (geo.calibration) {
    scaleInfoEl.hidden = false;
    const mmPerUnit = (geo.calibration.metersPerUnit * 1000).toFixed(2);
    scaleInfoEl.textContent = geo.calibration.pixelDist
      ? `Scale on this page: ${geo.calibration.label} = ${geo.calibration.pixelDist.toFixed(
          1
        )} plan units (1 unit ≈ ${mmPerUnit} mm).`
      : `Scale on this page: ${geo.calibration.label} (1 unit ≈ ${mmPerUnit} mm).`;
  } else {
    scaleInfoEl.hidden = true;
  }
}

function drawPoint(p, color) {
  const [cx, cy] = toCanvas(p);
  overlayCtx.beginPath();
  overlayCtx.arc(cx, cy, 4, 0, Math.PI * 2);
  overlayCtx.fillStyle = color;
  overlayCtx.fill();
}

// A bigger, hollow square rather than drawPoint's filled dot — reads as a
// "drag this" handle on a selected tracing's points rather than just a
// marker.
function drawHandle(p, color) {
  const [cx, cy] = toCanvas(p);
  const size = 9;
  overlayCtx.fillStyle = "#ffffff";
  overlayCtx.fillRect(cx - size / 2, cy - size / 2, size, size);
  overlayCtx.strokeStyle = color;
  overlayCtx.lineWidth = 2;
  overlayCtx.strokeRect(cx - size / 2, cy - size / 2, size, size);
}

// A hollow circle rather than drawHandle's square, so a selected tracing's
// label handle reads as a distinct kind of drag target from its points.
function drawLabelHandle(p, color) {
  const [cx, cy] = toCanvas(p);
  const r = 6;
  overlayCtx.beginPath();
  overlayCtx.arc(cx, cy, r, 0, Math.PI * 2);
  overlayCtx.fillStyle = "#ffffff";
  overlayCtx.fill();
  overlayCtx.strokeStyle = color;
  overlayCtx.lineWidth = 2;
  overlayCtx.stroke();
}

function drawLine(p1, p2, color, width, withMarkers, label) {
  const [x1, y1] = toCanvas(p1);
  const [x2, y2] = toCanvas(p2);
  overlayCtx.beginPath();
  overlayCtx.moveTo(x1, y1);
  overlayCtx.lineTo(x2, y2);
  overlayCtx.strokeStyle = color;
  overlayCtx.lineWidth = width;
  overlayCtx.stroke();
  if (withMarkers) {
    drawPoint(p1, color);
    drawPoint(p2, color);
  }
  if (label) {
    overlayCtx.fillStyle = color;
    overlayCtx.font = "13px Segoe UI, Arial, sans-serif";
    overlayCtx.fillText(label, (x1 + x2) / 2 + 4, (y1 + y2) / 2 - 4);
  }
}

function drawPolyline(points, color) {
  if (points.length === 0) return;
  overlayCtx.beginPath();
  const [x0, y0] = toCanvas(points[0]);
  overlayCtx.moveTo(x0, y0);
  for (let i = 1; i < points.length; i++) {
    const [x, y] = toCanvas(points[i]);
    overlayCtx.lineTo(x, y);
  }
  overlayCtx.strokeStyle = color;
  overlayCtx.lineWidth = 2;
  overlayCtx.stroke();
  points.forEach((p) => drawPoint(p, color));
}

function hexToRgba(hex, alpha) {
  const clean = (hex || "").replace("#", "");
  const r = parseInt(clean.substring(0, 2), 16);
  const g = parseInt(clean.substring(2, 4), 16);
  const b = parseInt(clean.substring(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

// labelPos, if given, overrides the default centroid placement — set by
// dragging a tracing's label handle (see labelDragState). When that moved
// position falls outside the tracing's own outline, a thin dashed line
// connects it back so it's still obviously that tracing's label.
function drawPolygon(points, color, label, lineWidth, labelPos) {
  if (points.length < 3) return;
  overlayCtx.beginPath();
  const [x0, y0] = toCanvas(points[0]);
  overlayCtx.moveTo(x0, y0);
  for (let i = 1; i < points.length; i++) {
    const [x, y] = toCanvas(points[i]);
    overlayCtx.lineTo(x, y);
  }
  overlayCtx.closePath();
  overlayCtx.fillStyle = hexToRgba(color, 0.15);
  overlayCtx.fill();
  overlayCtx.strokeStyle = color;
  overlayCtx.lineWidth = lineWidth || 2;
  overlayCtx.stroke();

  const centroid = polygonCentroid(points);
  const anchor = labelPos || centroid;

  if (labelPos && !pointInPolygon(labelPos, points)) {
    const [ax, ay] = toCanvas(anchor);
    const [ccx, ccy] = toCanvas(centroid);
    overlayCtx.beginPath();
    overlayCtx.moveTo(ax, ay);
    overlayCtx.lineTo(ccx, ccy);
    overlayCtx.strokeStyle = color;
    overlayCtx.lineWidth = 1;
    overlayCtx.setLineDash([4, 3]);
    overlayCtx.stroke();
    overlayCtx.setLineDash([]);
  }

  const [cx, cy] = toCanvas(anchor);
  overlayCtx.fillStyle = color;
  overlayCtx.font = "13px Segoe UI, Arial, sans-serif";
  overlayCtx.textAlign = "center";
  overlayCtx.fillText(label, cx, cy);
  overlayCtx.textAlign = "left";
}

// Same labelPos/leader-line convention as drawPolygon, for an open
// (unclosed, unfilled) length tracing — a polyline has no "inside" to test
// a dragged label against, so any custom labelPos at all is treated as
// "moved away from the default", and always gets a leader line back to the
// line's centroid.
function drawMeasuredLine(points, color, label, lineWidth, labelPos) {
  if (points.length < 2) return;
  overlayCtx.beginPath();
  const [x0, y0] = toCanvas(points[0]);
  overlayCtx.moveTo(x0, y0);
  for (let i = 1; i < points.length; i++) {
    const [x, y] = toCanvas(points[i]);
    overlayCtx.lineTo(x, y);
  }
  overlayCtx.strokeStyle = color;
  overlayCtx.lineWidth = lineWidth || 2;
  overlayCtx.stroke();

  const centroid = polygonCentroid(points);
  const anchor = labelPos || centroid;

  if (labelPos) {
    const [ax, ay] = toCanvas(anchor);
    const [ccx, ccy] = toCanvas(centroid);
    overlayCtx.beginPath();
    overlayCtx.moveTo(ax, ay);
    overlayCtx.lineTo(ccx, ccy);
    overlayCtx.strokeStyle = color;
    overlayCtx.lineWidth = 1;
    overlayCtx.setLineDash([4, 3]);
    overlayCtx.stroke();
    overlayCtx.setLineDash([]);
  }

  const [cx, cy] = toCanvas(anchor);
  overlayCtx.fillStyle = color;
  overlayCtx.font = "13px Segoe UI, Arial, sans-serif";
  overlayCtx.textAlign = "center";
  overlayCtx.fillText(label, cx, cy);
  overlayCtx.textAlign = "left";
}

// ---- Geometry helpers -----------------------------------------------------
function distance(p1, p2) {
  return Math.hypot(p2[0] - p1[0], p2[1] - p1[1]);
}

function polygonCentroid(points) {
  return points.reduce(
    (acc, p) => [acc[0] + p[0] / points.length, acc[1] + p[1] / points.length],
    [0, 0]
  );
}

function polygonAreaPageUnits(points) {
  let sum = 0;
  const n = points.length;
  for (let i = 0; i < n; i++) {
    const [x1, y1] = points[i];
    const [x2, y2] = points[(i + 1) % n];
    sum += x1 * y2 - x2 * y1;
  }
  return Math.abs(sum) / 2;
}

// Sum of consecutive segment lengths — deliberately does NOT add a closing
// segment back to the first point, unlike an area tracing's polygon: a
// length tracing measures the lines actually drawn, not a shape's perimeter.
function polylineLengthPageUnits(points) {
  let sum = 0;
  for (let i = 0; i < points.length - 1; i++) {
    sum += distance(points[i], points[i + 1]);
  }
  return sum;
}

// Shortest distance from point p to the segment a-b — used to hit-test a
// click against a length tracing's line, which (unlike an area's polygon)
// has no interior for a simple point-in-shape test to work against.
function distanceToSegment(p, a, b) {
  const [px, py] = p;
  const [ax, ay] = a;
  const [bx, by] = b;
  const dx = bx - ax;
  const dy = by - ay;
  const lenSq = dx * dx + dy * dy;
  if (lenSq === 0) return distance(p, a);
  let t = ((px - ax) * dx + (py - ay) * dy) / lenSq;
  t = Math.max(0, Math.min(1, t));
  return distance(p, [ax + t * dx, ay + t * dy]);
}

function pointNearPolyline(p, points, thresholdPageUnits) {
  for (let i = 0; i < points.length - 1; i++) {
    if (distanceToSegment(p, points[i], points[i + 1]) <= thresholdPageUnits) return true;
  }
  return false;
}

// Shared by the canvas overlay and the PDF exports — a tracing's label
// always shows its name and computed measurement, looked up from the flat
// `tracings` list (the source of truth for areaM2/lengthM) rather than
// recomputed here.
function tracingLabelText(geoTracing) {
  const flat = tracings.find((t) => t.id === geoTracing.id);
  if (!flat) return geoTracing.name;
  return geoTracing.kind === "length"
    ? `${geoTracing.name} — ${flat.lengthM.toFixed(2)} m`
    : `${geoTracing.name} — ${flat.areaM2.toFixed(2)} m²`;
}

function canvasPointFromEvent(evt) {
  const rect = overlayCanvas.getBoundingClientRect();
  const cx = evt.clientX - rect.left;
  const cy = evt.clientY - rect.top;
  return [cx / renderScale, cy / renderScale];
}

// Ray-casting point-in-polygon test, used to pick which area tracing a
// click landed on (page-space coordinates on both sides).
function pointInPolygon(p, points) {
  let inside = false;
  const [x, y] = p;
  for (let i = 0, j = points.length - 1; i < points.length; j = i++) {
    const [xi, yi] = points[i];
    const [xj, yj] = points[j];
    const intersect = yi > y !== yj > y && x < ((xj - xi) * (y - yi)) / (yj - yi) + xi;
    if (intersect) inside = !inside;
  }
  return inside;
}

// A fixed on-screen hit radius (converted to page-space via renderScale) so
// handles are just as easy to grab whether you're zoomed in or out.
const VERTEX_HIT_RADIUS_PX = 10;

function hitTestVertex(tracing, p) {
  const thresholdPageUnits = VERTEX_HIT_RADIUS_PX / renderScale;
  for (let i = 0; i < tracing.points.length; i++) {
    if (distance(tracing.points[i], p) <= thresholdPageUnits) return i;
  }
  return -1;
}

// A bit more forgiving than a vertex's, since the label handle sits over
// actual text rather than a single point.
const LABEL_HIT_RADIUS_PX = 14;

function hitTestLabel(tracing, p) {
  const anchor = tracing.labelPos || polygonCentroid(tracing.points);
  return distance(anchor, p) <= LABEL_HIT_RADIUS_PX / renderScale;
}

// Selecting a length tracing by clicking "inside" it doesn't make sense (an
// open line has no interior), so clicks are tested against its line itself
// within this radius instead — a bit more forgiving than a vertex handle's,
// since a thin line is a smaller target than a filled area's polygon.
const LINE_SELECT_RADIUS_PX = 8;

// ---- Mode / tool state ----------------------------------------------------
function resetToolState() {
  mode = "idle";
  calibTemp = { p1: null, p2: null };
  traceTemp = { page: null, points: [] };
  pendingTraceKind = null;
  editingTracing = null;
  vertexDragState = null;
  labelDragState = null;
  scaleRatioForm.hidden = true;
  calibrationForm.hidden = true;
  tracingNameForm.hidden = true;
  undoPointBtn.hidden = true;
  finishAreaBtn.hidden = true;
  measureLengthBtn.hidden = true;
  cancelScaleBtn.hidden = true;
  cancelTraceBtn.hidden = true;
}

overlayCanvas.addEventListener("click", (evt) => {
  // A drag (panning the view, or dragging a tracing's vertex handle — both
  // start on mousedown/mousemove on this same canvas) still ends with a
  // "click" event on mouseup even though the pointer moved, because native
  // click semantics don't care about distance travelled, only that
  // mousedown/mouseup landed on the same element. Left unchecked, that
  // turned every pan into an unwanted extra trace/calibration point. The
  // mouseup handler below sets this flag whenever the gesture it just
  // finished was actually a drag, so this one click event gets skipped.
  if (suppressNextClick) {
    suppressNextClick = false;
    return;
  }

  const p = canvasPointFromEvent(evt);

  if (mode === "calibrate") {
    if (!calibTemp.p1) {
      calibTemp.p1 = p;
      setStatus("Now click the other end of that same measurement.");
      redrawOverlay();
    } else if (!calibTemp.p2) {
      calibTemp.p2 = p;
      calibrationForm.hidden = false;
      calibLengthInput.focus();
      setStatus("Enter the real-world length of the line you just drew.");
      redrawOverlay();
    }
    return;
  }

  if (mode === "trace") {
    traceTemp.points.push(p);
    finishAreaBtn.disabled = traceTemp.points.length < 3;
    measureLengthBtn.disabled = traceTemp.points.length < 2;
    redrawOverlay();
    return;
  }

  if (mode === "idle") {
    const geo = currentGeometry();
    const thresholdPageUnits = LINE_SELECT_RADIUS_PX / renderScale;
    const hit = geo.tracings
      .slice()
      .reverse()
      .find((t) =>
        t.kind === "length" ? pointNearPolyline(p, t.points, thresholdPageUnits) : pointInPolygon(p, t.points)
      );
    editingTracing = hit ? hit.id : null;
    redrawOverlay();
    setStatus(
      hit
        ? `Editing "${hit.name}" — drag its point handles to reshape it, or the circle on its label to move it.`
        : ""
    );
  }
});

// ---- Pan (click-drag) and zoom (mouse wheel) -------------------------------
// Dragging a tracing's selected vertex handle (set up in mousedown below)
// takes priority over panning for that gesture; anywhere else on the canvas
// still pans as before.
let panState = null;
let suppressNextClick = false;

overlayCanvas.addEventListener("mousedown", (evt) => {
  if (mode === "idle" && editingTracing) {
    const geo = currentGeometry();
    const tracing = geo.tracings.find((t) => t.id === editingTracing);
    if (tracing) {
      const p = canvasPointFromEvent(evt);
      // Checked before the vertices — the label handle usually sits well
      // clear of the points, so there's little real ambiguity in practice.
      if (hitTestLabel(tracing, p)) {
        labelDragState = { tracingId: tracing.id, dragged: false };
        return;
      }
      const idx = hitTestVertex(tracing, p);
      if (idx !== -1) {
        vertexDragState = { tracingId: tracing.id, pointIndex: idx, dragged: false };
        return;
      }
    }
  }

  panState = {
    startX: evt.clientX,
    startY: evt.clientY,
    startScrollLeft: canvasScroller.scrollLeft,
    startScrollTop: canvasScroller.scrollTop,
    dragged: false,
  };
});

window.addEventListener("mousemove", (evt) => {
  if (labelDragState) {
    labelDragState.dragged = true;
    const geo = currentGeometry();
    const tracing = geo.tracings.find((t) => t.id === labelDragState.tracingId);
    if (tracing) {
      tracing.labelPos = canvasPointFromEvent(evt);
      redrawOverlay();
    }
    return;
  }

  if (vertexDragState) {
    vertexDragState.dragged = true;
    const geo = currentGeometry();
    const tracing = geo.tracings.find((t) => t.id === vertexDragState.tracingId);
    if (tracing) {
      tracing.points[vertexDragState.pointIndex] = canvasPointFromEvent(evt);
      redrawOverlay();
    }
    return;
  }

  if (!panState) return;
  const dx = evt.clientX - panState.startX;
  const dy = evt.clientY - panState.startY;
  if (!panState.dragged && Math.hypot(dx, dy) > 4) {
    panState.dragged = true;
    overlayCanvas.style.cursor = "grabbing";
  }
  if (panState.dragged) {
    canvasScroller.scrollLeft = panState.startScrollLeft - dx;
    canvasScroller.scrollTop = panState.startScrollTop - dy;
  }
});

window.addEventListener("mouseup", () => {
  if (labelDragState) {
    if (labelDragState.dragged) {
      saveAutosave();
      suppressNextClick = true;
    }
    labelDragState = null;
    return;
  }

  if (vertexDragState) {
    if (vertexDragState.dragged) {
      recalcAllMeasurementsForPage(currentPageNum);
      suppressNextClick = true;
    }
    vertexDragState = null;
    return;
  }

  if (panState && panState.dragged) {
    overlayCanvas.style.cursor = "crosshair";
    suppressNextClick = true;
  }
  panState = null;
});

// Zooms around the point under the cursor (rather than the canvas's
// top-left corner) so the plan doesn't visually jump while zooming in on
// a specific detail.
canvasScroller.addEventListener(
  "wheel",
  (evt) => {
    if (!currentPdf) return;
    evt.preventDefault();
    const rect = canvasScroller.getBoundingClientRect();
    const offsetX = evt.clientX - rect.left;
    const offsetY = evt.clientY - rect.top;
    const pdfX = (canvasScroller.scrollLeft + offsetX) / renderScale;
    const pdfY = (canvasScroller.scrollTop + offsetY) / renderScale;

    const oldScale = renderScale;
    renderScale = evt.deltaY < 0 ? Math.min(4, renderScale * 1.1) : Math.max(0.5, renderScale / 1.1);
    if (renderScale === oldScale) return;

    renderPage(() => {
      canvasScroller.scrollLeft = pdfX * renderScale - offsetX;
      canvasScroller.scrollTop = pdfY * renderScale - offsetY;
    });
  },
  { passive: false }
);

setScaleBtn.addEventListener("click", () => {
  resetToolState();
  mode = "calibrate";
  cancelScaleBtn.hidden = false;
  scaleRatioForm.hidden = false;
  setStatus(
    "Click one end of a known measurement on the plan (a scale bar or a labelled dimension), or enter the drawing's printed scale below."
  );
});

traceBtn.addEventListener("click", () => {
  resetToolState();
  mode = "trace";
  traceTemp = { page: currentPageNum, points: [] };
  undoPointBtn.hidden = false;
  finishAreaBtn.hidden = false;
  finishAreaBtn.disabled = true;
  measureLengthBtn.hidden = false;
  measureLengthBtn.disabled = true;
  cancelTraceBtn.hidden = false;
  setStatus(
    "Click each point of the outline in order, then click “Finish area” to close it as an area, or “Measure length” to just measure the lines drawn so far."
  );
});

undoPointBtn.addEventListener("click", () => {
  traceTemp.points.pop();
  finishAreaBtn.disabled = traceTemp.points.length < 3;
  measureLengthBtn.disabled = traceTemp.points.length < 2;
  redrawOverlay();
});

function cancelCurrentAction() {
  resetToolState();
  setStatus("Cancelled.");
  redrawOverlay();
}
cancelScaleBtn.addEventListener("click", cancelCurrentAction);
cancelTraceBtn.addEventListener("click", cancelCurrentAction);

// A printed scale like "1:100" means 1 PDF point on the page (1/72 inch)
// represents 100 times that in real life — this only holds if the PDF's
// page size actually matches the real sheet size (true for a properly
// exported/unresized architectural PDF, which is the assumption called out
// in the form itself). Accepts "1:100", "1/100", or just a bare "100".
function parseScaleRatio(text) {
  const trimmed = (text || "").trim();
  const match =
    /^1\s*[:/]\s*(\d+(?:\.\d+)?)$/.exec(trimmed) || /^(\d+(?:\.\d+)?)$/.exec(trimmed);
  if (!match) return null;
  const n = parseFloat(match[1]);
  return n > 0 ? n : null;
}
const METERS_PER_POINT = 0.0254 / 72;

// Shared by the free-text ratio input and the one-click preset buttons
// (1:50/1:100/1:200) below — both just need to apply a given ratio.
function applyScaleRatio(n) {
  const geo = currentGeometry();
  geo.calibration = {
    p1: null,
    p2: null,
    pixelDist: null,
    metersPerUnit: METERS_PER_POINT * n,
    label: `1:${n}`,
  };
  scaleRatioInput.value = "";
  resetToolState();
  redrawOverlay();
  recalcAllMeasurementsForPage(currentPageNum);
  saveAutosave();
  setStatus(`Scale set to 1:${n} for this page. Click “Trace” to start your first tracing.`);
}

scaleRatioConfirmBtn.addEventListener("click", () => {
  const n = parseScaleRatio(scaleRatioInput.value);
  if (!n) {
    setStatus("Enter a scale like 1:100.", true);
    return;
  }
  applyScaleRatio(n);
});

document.querySelectorAll(".scale-preset-btn").forEach((btn) => {
  btn.addEventListener("click", () => applyScaleRatio(Number(btn.dataset.scale)));
});

calibConfirmBtn.addEventListener("click", () => {
  const value = parseFloat(calibLengthInput.value);
  const unit = calibUnitSelect.value;
  if (!value || value <= 0) {
    setStatus("Enter a positive length first.", true);
    return;
  }
  const lengthM = value * UNIT_TO_M[unit];
  const pixelDist = distance(calibTemp.p1, calibTemp.p2);
  if (pixelDist === 0) {
    setStatus("Those two points were the same — try again.", true);
    return;
  }
  const geo = currentGeometry();
  geo.calibration = {
    p1: calibTemp.p1,
    p2: calibTemp.p2,
    pixelDist,
    metersPerUnit: lengthM / pixelDist,
    label: `${value}${unit}`,
  };
  resetToolState();
  redrawOverlay();
  recalcAllMeasurementsForPage(currentPageNum);
  saveAutosave();
  setStatus("Scale set for this page. Click “Trace” to start your first tracing.");
});

calibCancelBtn.addEventListener("click", () => {
  resetToolState();
  redrawOverlay();
});

finishAreaBtn.addEventListener("click", () => {
  if (traceTemp.points.length < 3) return;
  pendingTraceKind = "area";
  openTracingNameForm();
});

measureLengthBtn.addEventListener("click", () => {
  if (traceTemp.points.length < 2) return;
  pendingTraceKind = "length";
  openTracingNameForm();
});

function openTracingNameForm() {
  tracingNameForm.hidden = false;
  tracingNameInput.placeholder = pendingTraceKind === "length" ? "e.g. Wall run" : "e.g. Living room";
  tracingNameInput.value = `Tracing ${tracings.length + 1}`;
  tracingColorInput.value = paletteColor(tracings.length);
  tracingNameInput.focus();
  tracingNameInput.select();
}

tracingNameConfirmBtn.addEventListener("click", () => {
  const geo = currentGeometry();
  if (!geo.calibration) {
    setStatus("Scale isn't set for this page anymore — set it again first.", true);
    return;
  }
  const name = tracingNameInput.value.trim() || `Tracing ${tracings.length + 1}`;
  const color = tracingColorInput.value || DEFAULT_TRACING_COLOR;
  const kind = pendingTraceKind || "area";
  const id = nextTracingId++;
  const points = traceTemp.points.slice();

  geo.tracings.push({ id, name, points, color, kind });

  let measurementText;
  if (kind === "length") {
    const lengthPageUnits = polylineLengthPageUnits(points);
    const lengthM = lengthPageUnits * geo.calibration.metersPerUnit;
    tracings.push({ id, page: currentPageNum, name, color, kind, lengthM });
    measurementText = `${lengthM.toFixed(2)} m`;
  } else {
    const areaPageUnits = polygonAreaPageUnits(points);
    const areaM2 = areaPageUnits * geo.calibration.metersPerUnit * geo.calibration.metersPerUnit;
    tracings.push({ id, page: currentPageNum, name, color, kind, areaM2 });
    measurementText = `${areaM2.toFixed(2)} m²`;
  }

  tracingNameForm.hidden = true;
  resetToolState();
  redrawOverlay();
  updateResultsTable();
  saveAutosave();
  setStatus(`Added "${name}" — ${measurementText}. Trace another outline, or move to the next page.`);
});

tracingNameCancelBtn.addEventListener("click", () => {
  tracingNameForm.hidden = true;
});

// None of these inline forms are real <form> elements (a plain <div>, same
// reasoning as everywhere else in this app — no accidental page navigation
// from an implicit submit), which also means Enter in a text field does
// nothing by default. Pressing Enter to submit a single-field form is a
// reasonable enough expectation that it's worth wiring up explicitly rather
// than silently doing nothing and leaving whoever typed a scale or a name
// wondering why the button they didn't click never fired.
function submitOnEnter(input, btn) {
  input.addEventListener("keydown", (evt) => {
    if (evt.key === "Enter") {
      evt.preventDefault();
      btn.click();
    }
  });
}
submitOnEnter(scaleRatioInput, scaleRatioConfirmBtn);
submitOnEnter(calibLengthInput, calibConfirmBtn);
submitOnEnter(tracingNameInput, tracingNameConfirmBtn);

// If the scale is re-set on a page, existing tracings on that page keep
// their drawn points but their area/length is recalculated against the new
// scale.
function recalcAllMeasurementsForPage(pageNum) {
  const geo = pageGeometry[pageNum];
  if (!geo || !geo.calibration) return;
  geo.tracings.forEach((t) => {
    const flat = tracings.find((x) => x.id === t.id);
    if (!flat) return;
    if (t.kind === "length") {
      const lengthPageUnits = polylineLengthPageUnits(t.points);
      flat.lengthM = lengthPageUnits * geo.calibration.metersPerUnit;
    } else {
      const areaPageUnits = polygonAreaPageUnits(t.points);
      flat.areaM2 = areaPageUnits * geo.calibration.metersPerUnit * geo.calibration.metersPerUnit;
    }
  });
  updateResultsTable();
}

// ---- Page nav / zoom -------------------------------------------------------
prevPageBtn.addEventListener("click", () => {
  if (currentPageNum > 1) {
    resetToolState();
    currentPageNum -= 1;
    renderPage();
  }
});
nextPageBtn.addEventListener("click", () => {
  if (currentPdf && currentPageNum < currentPdf.numPages) {
    resetToolState();
    currentPageNum += 1;
    renderPage();
  }
});
zoomInBtn.addEventListener("click", () => {
  renderScale = Math.min(4, renderScale * 1.25);
  renderPage();
});
zoomOutBtn.addEventListener("click", () => {
  renderScale = Math.max(0.5, renderScale / 1.25);
  renderPage();
});

// Clicking a tracing's row (see updateResultsTable) jumps the viewer to
// that tracing's page and selects it, same as clicking its outline on the
// canvas — resetToolState() first so any in-progress trace/calibration on
// the page being left doesn't linger, and editingTracing is set *after* it
// (which would otherwise clear it) so the tracing is already selected by
// the time the new page's render finishes and calls redrawOverlay().
function goToTracing(id, pageNum) {
  resetToolState();
  if (pageNum !== currentPageNum) {
    currentPageNum = pageNum;
    editingTracing = id;
    renderPage();
  } else {
    editingTracing = id;
    redrawOverlay();
  }
  const flat = tracings.find((t) => t.id === id);
  setStatus(
    flat
      ? `Editing "${flat.name}" — drag its point handles to reshape it, or the circle on its label to move it.`
      : ""
  );
}

// ---- 3. Results table ----------------------------------------------------
function updateResultsTable() {
  resultsBody.innerHTML = "";
  let totalAreaM2 = 0;
  let totalLengthM = 0;
  let hasArea = false;
  let hasLength = false;
  // Sort is stable, so tracings on the same page keep their original
  // (creation) order relative to each other — only the page grouping
  // changes. `tracings` itself stays in insertion order for everything
  // else that reads it (naming defaults, copy/export), only this render
  // is page-ordered.
  const displayTracings = tracings.slice().sort((a, b) => a.page - b.page);
  displayTracings.forEach((t) => {
    const tr = document.createElement("tr");

    const colorTd = document.createElement("td");
    const colorInput = document.createElement("input");
    colorInput.type = "color";
    colorInput.value = t.color || DEFAULT_TRACING_COLOR;
    colorInput.title = "Tracing colour";
    colorInput.addEventListener("input", () => updateTracingColor(t.id, colorInput.value));
    colorTd.appendChild(colorInput);

    const nameTd = document.createElement("td");
    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.className = "tracing-name-input";
    nameInput.value = t.name;
    nameInput.title = "Tracing name";
    nameInput.addEventListener("change", () => {
      nameInput.value = updateTracingName(t.id, nameInput.value);
    });
    nameInput.addEventListener("keydown", (evt) => {
      if (evt.key === "Enter") nameInput.blur();
    });
    nameTd.appendChild(nameInput);

    const pageTd = document.createElement("td");
    pageTd.textContent = t.page;

    // Read live off that page's own calibration rather than stored on the
    // tracing, so it can't go stale if the page's scale is ever re-set.
    const scaleTd = document.createElement("td");
    const pageGeo = pageGeometry[t.page];
    scaleTd.textContent = pageGeo && pageGeo.calibration ? pageGeo.calibration.label : "—";

    const measurementTd = document.createElement("td");
    if (t.kind === "length") {
      hasLength = true;
      totalLengthM += t.lengthM;
      measurementTd.textContent = `${t.lengthM.toFixed(2)} m`;
    } else {
      hasArea = true;
      totalAreaM2 += t.areaM2;
      measurementTd.textContent = `${t.areaM2.toFixed(2)} m²`;
    }

    const delTd = document.createElement("td");
    const delBtn = document.createElement("button");
    delBtn.className = "del-btn";
    delBtn.textContent = "✕";
    delBtn.title = "Remove this tracing";
    delBtn.addEventListener("click", () => removeTracing(t.id));
    delTd.appendChild(delBtn);

    tr.appendChild(colorTd);
    tr.appendChild(nameTd);
    tr.appendChild(pageTd);
    tr.appendChild(scaleTd);
    tr.appendChild(measurementTd);
    tr.appendChild(delTd);
    // Ignore clicks that landed on one of the row's own interactive
    // controls (colour swatch, name input, delete button) — those already
    // have their own handlers, and jumping pages out from under a click
    // meant to edit the name/colour would be surprising.
    tr.addEventListener("click", (evt) => {
      if (evt.target.closest("input, button")) return;
      goToTracing(t.id, t.page);
    });
    resultsBody.appendChild(tr);
  });

  // Show each kind's total only if there's actually one of that kind —
  // otherwise (including the very first "nothing traced yet" state) fall
  // back to a single "0.00 m²" placeholder, matching this app's original
  // always-area-only zero state.
  let totalLines = [];
  if (hasArea) totalLines.push(`${totalAreaM2.toFixed(2)} m²`);
  if (hasLength) totalLines.push(`${totalLengthM.toFixed(2)} m`);
  if (totalLines.length === 0) totalLines.push("0.00 m²");
  totalMeasurementEl.innerHTML = totalLines.map((l) => `<strong>${l}</strong>`).join("<br>");
}

function removeTracing(id) {
  tracings = tracings.filter((t) => t.id !== id);
  Object.values(pageGeometry).forEach((geo) => {
    geo.tracings = geo.tracings.filter((t) => t.id !== id);
  });
  if (editingTracing === id) editingTracing = null;
  redrawOverlay();
  updateResultsTable();
  saveAutosave();
}

// Kept in sync on both the flat `tracings` list (what the results table's
// swatch reflects) and the per-page geometry (what drawing/export reads) —
// deliberately doesn't call updateResultsTable(), which would rebuild the
// row out from under the very <input type="color"> the user is still
// interacting with.
function updateTracingColor(id, color) {
  const flat = tracings.find((t) => t.id === id);
  if (flat) flat.color = color;
  const geo = pageGeometry[flat ? flat.page : currentPageNum];
  const geoTracing = geo && geo.tracings.find((t) => t.id === id);
  if (geoTracing) geoTracing.color = color;
  redrawOverlay();
  saveAutosave();
}

// Same pattern as updateTracingColor — kept in sync on both the flat list
// and the per-page geometry, and deliberately skips updateResultsTable() so
// the row isn't rebuilt out from under the input the user just typed into.
function updateTracingName(id, name) {
  const trimmed = name.trim() || "Tracing";
  const flat = tracings.find((t) => t.id === id);
  if (flat) flat.name = trimmed;
  const geo = pageGeometry[flat ? flat.page : currentPageNum];
  const geoTracing = geo && geo.tracings.find((t) => t.id === id);
  if (geoTracing) geoTracing.name = trimmed;
  redrawOverlay();
  saveAutosave();
  return trimmed;
}

copyResultsBtn.addEventListener("click", () => {
  let text = "Tracing\tPage\tMeasurement\n";
  let totalAreaM2 = 0;
  let totalLengthM = 0;
  let hasArea = false;
  let hasLength = false;
  tracings.forEach((t) => {
    if (t.kind === "length") {
      hasLength = true;
      totalLengthM += t.lengthM;
      text += `${t.name}\t${t.page}\t${t.lengthM.toFixed(2)} m\n`;
    } else {
      hasArea = true;
      totalAreaM2 += t.areaM2;
      text += `${t.name}\t${t.page}\t${t.areaM2.toFixed(2)} m²\n`;
    }
  });
  if (hasArea || !hasLength) text += `Total area\t\t${totalAreaM2.toFixed(2)} m²\n`;
  if (hasLength) text += `Total length\t\t${totalLengthM.toFixed(2)} m\n`;

  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(
      () => setStatus("Results copied — paste into your quote/spreadsheet."),
      () => showCopyFallback(text)
    );
  } else {
    showCopyFallback(text);
  }
});

function showCopyFallback(text) {
  copyFallback.hidden = false;
  copyFallback.value = text;
  copyFallback.focus();
  copyFallback.select();
  setStatus("Couldn't copy automatically — the text is selected below, press Ctrl/Cmd+C.");
}

clearAllBtn.addEventListener("click", () => {
  tracings = [];
  nextTracingId = 1;
  pageGeometry = {};
  resetToolState();
  redrawOverlay();
  updateResultsTable();
  clearAutosave();
  setStatus("Cleared all tracings and scale settings on every page.");
});

// ---- 4. Download PDF with tracing data embedded ----------------------------
// Builds a modified copy of the original PDF — never the in-memory copy
// pdf.js is using, since getDocument() can transfer/detach that buffer —
// with the tracings/labels drawn directly onto the pages (so the
// measurements are visible in any ordinary PDF viewer), optionally a
// summary page listing every tracing and its measurement, and the raw
// geometry embedded as a JSON file attachment so re-opening this exact
// file restores the exact editable state — whether it kept every original
// page or only the traced ones (see buildDownloadPdfBytes for how the
// embedded page numbers get re-keyed to match whatever pages actually
// ended up in it).
function suggestDownloadName(name, pagesMode, targetPageNum) {
  const base = (name || "floor-plan").replace(/\.pdf$/i, "");
  if (pagesMode === "traced") return `${base}-traced-pages.pdf`;
  if (pagesMode === "current") return `${base}-page-${targetPageNum}.pdf`;
  return `${base}-with-tracings.pdf`;
}

function hexToPdfRgb(lib, hex) {
  const clean = (hex || "").replace("#", "");
  const r = parseInt(clean.substring(0, 2), 16) / 255;
  const g = parseInt(clean.substring(2, 4), 16) / 255;
  const b = parseInt(clean.substring(4, 6), 16) / 255;
  return lib.rgb(r, g, b);
}

// Draws every tracing on `geo` onto `pdfLibPage`, each in its own chosen
// colour, using `viewportAtScale1` to map our stored points (pdf.js's
// scale-1 viewport space, see canvasPointFromEvent) back to the PDF's own
// coordinate space (bottom-left origin, correctly accounting for page
// rotation), which is what pdf-lib's drawing calls expect.
function drawTracingAnnotations(lib, pdfLibPage, geo, viewportAtScale1, font) {
  geo.tracings.forEach((tracing) => {
    const isLength = tracing.kind === "length";
    if (isLength ? tracing.points.length < 2 : tracing.points.length < 3) return;
    const color = hexToPdfRgb(lib, tracing.color || DEFAULT_TRACING_COLOR);
    const toPdfPoint = (p) => {
      const [x, y] = viewportAtScale1.convertToPdfPoint(p[0], p[1]);
      return { x, y };
    };
    const pts = tracing.points.map(toPdfPoint);
    if (isLength) {
      for (let i = 0; i < pts.length - 1; i++) {
        pdfLibPage.drawLine({ start: pts[i], end: pts[i + 1], thickness: 1.5, color });
      }
    } else {
      for (let i = 0; i < pts.length; i++) {
        const a = pts[i];
        const b = pts[(i + 1) % pts.length];
        pdfLibPage.drawLine({ start: a, end: b, thickness: 1.5, color });
      }
    }

    // Same labelPos/leader-line logic as the canvas overlay (drawPolygon /
    // drawMeasuredLine) — tracing.points and tracing.labelPos are both in
    // pdf.js's scale-1 page space, so pointInPolygon works the same way
    // here as it does there.
    const centroidPage = polygonCentroid(tracing.points);
    const anchorPage = tracing.labelPos || centroidPage;
    const anchor = toPdfPoint(anchorPage);
    const showLeader = isLength
      ? !!tracing.labelPos
      : tracing.labelPos && !pointInPolygon(tracing.labelPos, tracing.points);
    if (showLeader) {
      pdfLibPage.drawLine({
        start: anchor,
        end: toPdfPoint(centroidPage),
        thickness: 1,
        color,
        dashArray: [4, 3],
      });
    }

    const label = tracingLabelText(tracing);
    pdfLibPage.drawText(label, {
      x: anchor.x - label.length * 2.3,
      y: anchor.y,
      size: 9,
      font,
      color,
    });
  });
}

// A plain, paginating table of the given tracings and their measurements
// appended to the end of the document — header row, one row per tracing, a
// total row, starting a fresh page whenever the current one runs out of
// room. Uses the same page size as the plan itself so it sits consistently
// alongside it. `summaryTracings` is every tracing for a whole-document/
// traced-pages download, or just the ones on the kept page for a
// current-page-only download — listing tracings that aren't even in this
// file would be confusing.
async function addSummaryPages(lib, outDoc, pageWidth, pageHeight, summaryTracings) {
  const font = await outDoc.embedFont(lib.StandardFonts.Helvetica);
  const boldFont = await outDoc.embedFont(lib.StandardFonts.HelveticaBold);
  const black = lib.rgb(0, 0, 0);
  const grey = lib.rgb(0.6, 0.6, 0.6);
  const margin = 40;
  const rowHeight = 18;
  const colX = { name: margin, page: pageWidth - 170, measurement: pageWidth - 90 };

  let page = null;
  let y = 0;

  function drawHeaderRow() {
    page.drawText("Tracing", { x: colX.name, y, size: 11, font: boldFont, color: black });
    page.drawText("Page", { x: colX.page, y, size: 11, font: boldFont, color: black });
    page.drawText("Measurement", { x: colX.measurement, y, size: 11, font: boldFont, color: black });
    y -= 6;
    page.drawLine({ start: { x: margin, y }, end: { x: pageWidth - margin, y }, thickness: 1, color: grey });
    y -= rowHeight;
  }

  function startPage(withTitle) {
    page = outDoc.addPage([pageWidth, pageHeight]);
    y = pageHeight - margin;
    if (withTitle) {
      page.drawText("Takeoff Summary", { x: margin, y, size: 16, font: boldFont, color: black });
      y -= 28;
    }
    drawHeaderRow();
  }

  startPage(true);

  let totalArea = 0;
  let totalLength = 0;
  let hasArea = false;
  let hasLength = false;
  summaryTracings.forEach((t) => {
    if (y < margin + rowHeight * 2) startPage(false);
    const measureText = t.kind === "length" ? `${t.lengthM.toFixed(2)} m` : `${t.areaM2.toFixed(2)} m²`;
    page.drawText(t.name, { x: colX.name, y, size: 10, font, color: black });
    page.drawText(String(t.page), { x: colX.page, y, size: 10, font, color: black });
    page.drawText(measureText, { x: colX.measurement, y, size: 10, font, color: black });
    if (t.kind === "length") {
      hasLength = true;
      totalLength += t.lengthM;
    } else {
      hasArea = true;
      totalArea += t.areaM2;
    }
    y -= rowHeight;
  });

  y -= 4;
  page.drawLine({
    start: { x: margin, y: y + rowHeight - 4 },
    end: { x: pageWidth - margin, y: y + rowHeight - 4 },
    thickness: 1.5,
    color: black,
  });
  // Same kind-symmetric rule as the on-screen totals: don't show a "Total
  // area: 0.00 m²" line when nothing here is actually an area (and vice
  // versa), except in the genuinely-empty case, which keeps the original
  // always-area-only placeholder.
  if (hasArea || !hasLength) {
    page.drawText("Total area", { x: colX.name, y, size: 11, font: boldFont, color: black });
    page.drawText(`${totalArea.toFixed(2)} m²`, { x: colX.measurement, y, size: 11, font: boldFont, color: black });
    if (hasLength) y -= rowHeight;
  }
  if (hasLength) {
    page.drawText("Total length", { x: colX.name, y, size: 11, font: boldFont, color: black });
    page.drawText(`${totalLength.toFixed(2)} m`, { x: colX.measurement, y, size: 11, font: boldFont, color: black });
  }
}

// pagesMode: "all" keeps every original page; "traced" builds a brand-new
// document with just the pages that actually have a tracing on them. Either
// way the embedded round-trip geometry JSON is keyed to *this output
// document's own page numbers*, not the original file's — when pages are
// dropped, page 1 of the trimmed copy might be page 5 of the original, so
// the geometry saved under "page 5" gets re-keyed to "page 1" here. That's
// what makes a reopened copy of either kind restore correctly: its embedded
// data always describes the pages actually in that file. includeSummary
// appends the takeoff summary table as extra page(s) either way.
async function buildDownloadPdfBytes({ pagesMode, includeSummary, targetPageNum }) {
  const lib = await loadPdfLib();
  let outDoc;
  let pageEntries; // [{ origPageNum, pdfLibPage }], in the order they end up in outDoc
  let geometryForEmbed;
  let summaryTracings = tracings;

  if (pagesMode === "traced") {
    const tracedPageNums = Object.keys(pageGeometry)
      .map(Number)
      .filter((pageNum) => {
        const geo = pageGeometry[pageNum];
        return geo && geo.tracings && geo.tracings.length > 0;
      })
      .sort((a, b) => a - b);

    if (tracedPageNums.length === 0) {
      throw new Error("No tracings yet — add at least one first.");
    }

    const srcDoc = await lib.PDFDocument.load(originalBytes.slice());
    outDoc = await lib.PDFDocument.create();
    const copiedPages = await outDoc.copyPages(srcDoc, tracedPageNums.map((n) => n - 1));
    copiedPages.forEach((page) => outDoc.addPage(page));

    pageEntries = tracedPageNums.map((origPageNum, i) => ({ origPageNum, pdfLibPage: copiedPages[i] }));
    geometryForEmbed = {};
    pageEntries.forEach((entry, i) => {
      geometryForEmbed[i + 1] = pageGeometry[entry.origPageNum];
    });
  } else if (pagesMode === "current") {
    const origPageNum = targetPageNum;
    const srcDoc = await lib.PDFDocument.load(originalBytes.slice());
    outDoc = await lib.PDFDocument.create();
    const [copiedPage] = await outDoc.copyPages(srcDoc, [origPageNum - 1]);
    outDoc.addPage(copiedPage);

    pageEntries = [{ origPageNum, pdfLibPage: copiedPage }];
    geometryForEmbed = pageGeometry[origPageNum] ? { 1: pageGeometry[origPageNum] } : {};
    summaryTracings = tracings.filter((t) => t.page === origPageNum);
  } else {
    outDoc = await lib.PDFDocument.load(originalBytes.slice());
    pageEntries = outDoc.getPages().map((pdfLibPage, i) => ({ origPageNum: i + 1, pdfLibPage }));
    geometryForEmbed = pageGeometry;
  }

  const geometryJson = JSON.stringify({ pageGeometry: geometryForEmbed, nextTracingId }, null, 2);
  await outDoc.attach(new TextEncoder().encode(geometryJson), "floor-area-takeoff.json", {
    mimeType: "application/json",
    description: "Surface Takeoff — scale calibration and tracings",
  });

  const font = await outDoc.embedFont(lib.StandardFonts.Helvetica);
  for (const { origPageNum, pdfLibPage } of pageEntries) {
    const geo = pageGeometry[origPageNum];
    if (!geo || !geo.tracings || geo.tracings.length === 0) continue;
    // currentPdf is pdf.js's parse of the *original* file, so this lookup
    // always needs the original page number, regardless of where that
    // page ended up (or whether it was renumbered) in outDoc.
    const pdfjsPage = await currentPdf.getPage(origPageNum);
    const viewportAtScale1 = pdfjsPage.getViewport({ scale: 1 });
    drawTracingAnnotations(lib, pdfLibPage, geo, viewportAtScale1, font);
  }

  if (includeSummary) {
    const refPage = outDoc.getPages()[0];
    await addSummaryPages(lib, outDoc, refPage.getWidth(), refPage.getHeight(), summaryTracings);
  }

  return outDoc.save();
}

// A real browser tab (unlike the Outlook add-in's embedded dialog WebView
// this replaced) handles a Blob URL + a programmatically-clicked <a
// download> perfectly normally, even after the async build work below —
// this is an ordinary, well-supported web pattern here, not the two-phase
// dance the add-in needed.
downloadConfirmBtn.addEventListener("click", () => {
  if (!currentPdf || !originalBytes || downloadConfirmBtn.dataset.busy === "1") return;

  const pagesMode = document.querySelector('input[name="downloadPages"]:checked').value;
  const includeSummary = downloadIncludeSummary.checked;
  // Captured now, synchronously, rather than read back out of currentPageNum
  // once the async build below finishes — the user could navigate pages
  // while "Preparing…" is showing, and this download is about the page
  // they were on when they clicked, not whatever page is current later.
  const targetPageNum = currentPageNum;

  const originalText = downloadConfirmBtn.textContent;
  downloadConfirmBtn.dataset.busy = "1";
  downloadConfirmBtn.textContent = "Preparing…";
  setStatus("Preparing PDF…");

  buildDownloadPdfBytes({ pagesMode, includeSummary, targetPageNum }).then(
    (bytes) => {
      const blob = new Blob([bytes], { type: "application/pdf" });
      const url = URL.createObjectURL(blob);
      const filename = suggestDownloadName(currentFileName, pagesMode, targetPageNum);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30000);
      downloadConfirmBtn.textContent = originalText;
      downloadConfirmBtn.dataset.busy = "";
      setStatus(`Downloaded "${filename}".`);
    },
    (err) => {
      downloadConfirmBtn.textContent = originalText;
      downloadConfirmBtn.dataset.busy = "";
      setStatus("Couldn't build the download — " + describeError(err), true);
    }
  );
});
