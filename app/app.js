/* Floor Area Takeoff — standalone browser app
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

// ---- State -----------------------------------------------------------
let currentPdf = null;
let currentPageNum = 1;
let renderScale = 1.5;
let currentFileName = null;
// Kept separately from whatever pdf.js does with its own copy of the bytes
// (getDocument() can transfer/detach the buffer it's given) so the download
// button always has a pristine, untouched original to build from.
let originalBytes = null;

/** pageGeometry[pageNum] = { calibration: {p1,p2,metersPerUnit,label} | null, rooms: [{id,name,points}] } */
let pageGeometry = {};

/** flat list mirrored to the results table */
let rooms = []; // { id, page, name, areaM2 }
let nextRoomId = 1;

let mode = "idle"; // idle | calibrate | trace
let calibTemp = { p1: null, p2: null };
let traceTemp = { page: null, points: [] };

// Clicking an already-traced room's outline while idle selects it for
// editing — its corner points are drawn as draggable handles until you
// click elsewhere (or start calibrating/tracing) to deselect.
let editingRoom = null; // room id, or null
let vertexDragState = null; // { roomId, pointIndex, dragged }

// ---- DOM refs ----------------------------------------------------------
const statusBarEl = document.getElementById("statusBar");
const fileNameHeadingEl = document.getElementById("fileNameHeading");

const intakeSectionEl = document.getElementById("intakeSection");
const dropZone = document.getElementById("dropZone");
const pickFileBtn = document.getElementById("pickFileBtn");
const fileInput = document.getElementById("fileInput");
const zipPicker = document.getElementById("zipPicker");
const zipPickerList = document.getElementById("zipPickerList");

const viewerSectionEl = document.getElementById("viewerSection");
const resultsSectionEl = document.getElementById("resultsSection");

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
const traceRoomBtn = document.getElementById("traceRoomBtn");
const undoPointBtn = document.getElementById("undoPointBtn");
const finishRoomBtn = document.getElementById("finishRoomBtn");
const cancelActionBtn = document.getElementById("cancelActionBtn");

const scaleRatioForm = document.getElementById("scaleRatioForm");
const scaleRatioInput = document.getElementById("scaleRatioInput");
const scaleRatioConfirmBtn = document.getElementById("scaleRatioConfirmBtn");

const calibrationForm = document.getElementById("calibrationForm");
const calibLengthInput = document.getElementById("calibLengthInput");
const calibUnitSelect = document.getElementById("calibUnitSelect");
const calibConfirmBtn = document.getElementById("calibConfirmBtn");
const calibCancelBtn = document.getElementById("calibCancelBtn");

const roomNameForm = document.getElementById("roomNameForm");
const roomNameInput = document.getElementById("roomNameInput");
const roomNameConfirmBtn = document.getElementById("roomNameConfirmBtn");
const roomNameCancelBtn = document.getElementById("roomNameCancelBtn");

const scaleInfoEl = document.getElementById("scaleInfo");

const resultsBody = document.getElementById("resultsBody");
const totalM2El = document.getElementById("totalM2");
const copyResultsBtn = document.getElementById("copyResultsBtn");
const downloadPdfBtn = document.getElementById("downloadPdfBtn");
const downloadTracedPagesBtn = document.getElementById("downloadTracedPagesBtn");
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
function openPdfFromBytes(fileName, bytes) {
  setStatus(`Loading "${fileName}"…`);
  loadPdfJs().then(
    () => {
      // A copy for pdf.js — it can transfer/detach the buffer it's given,
      // and `bytes` needs to stay pristine for the download button to embed
      // later (see the doc comment on originalBytes above).
      pdfjsLib.getDocument({ data: bytes.slice() }).promise.then(
        (pdf) => {
          originalBytes = bytes;
          currentFileName = fileName;
          currentPdf = pdf;
          currentPageNum = 1;
          pageGeometry = {};
          rooms = [];
          nextRoomId = 1;
          renderScale = 1.5;
          resetToolState();
          setScaleBtn.disabled = false;
          fileNameHeadingEl.textContent = `2. Set scale, then trace rooms — "${fileName}"`;
          viewerSectionEl.hidden = false;
          resultsSectionEl.hidden = false;
          downloadPdfBtn.hidden = false;
          downloadTracedPagesBtn.hidden = false;
          renderPage();
          updateResultsTable();
          setStatus(`Loaded "${fileName}". Set the scale, then trace each room.`);
          restorePreviousProgress(pdf, fileName, bytes.length);
        },
        (err) => setStatus(`Couldn't open "${fileName}" — ` + describeError(err), true)
      );
    },
    (err) => setStatus("Couldn't load the PDF viewer library — " + describeError(err), true)
  );
}

// Two sources of previously-saved progress, checked in order: this exact
// file already opened once in this browser (autosaved to localStorage,
// keyed by name+size — see saveAutosave), or the file itself is a
// previously-downloaded round-trip copy with its own embedded scale/room
// data (see buildAnnotatedPdfBytes). Either way, restoring means the flat
// room list gets recomputed from the polygons rather than trusted as saved,
// so it can never drift from what's actually drawn.
function restorePreviousProgress(pdf, fileName, byteLength) {
  const autosaved = loadAutosave(fileName, byteLength);
  if (autosaved && autosaved.pageGeometry && Object.keys(autosaved.pageGeometry).length > 0) {
    applyRestoredGeometry(autosaved.pageGeometry, autosaved.nextRoomId);
    setStatus("Restored your previous scale and room traces for this plan.");
    return;
  }
  pdf.getAttachments().then(
    (attachments) => {
      const embedded = attachments && attachments["floor-area-takeoff.json"];
      if (!embedded || !embedded.content) return;
      try {
        const parsed = JSON.parse(new TextDecoder().decode(embedded.content));
        if (parsed && parsed.pageGeometry && Object.keys(parsed.pageGeometry).length > 0) {
          applyRestoredGeometry(parsed.pageGeometry, parsed.nextRoomId);
          setStatus("Restored the scale and room traces embedded in this PDF.");
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
  nextRoomId = nextId || 1;
  rooms = [];
  editingRoom = null;
  vertexDragState = null;
  Object.keys(pageGeometry).forEach((pageNumKey) => {
    const pageNum = Number(pageNumKey);
    const geo = pageGeometry[pageNum];
    (geo.rooms || []).forEach((r) => {
      const areaPageUnits = polygonAreaPageUnits(r.points);
      const areaM2 = geo.calibration
        ? areaPageUnits * geo.calibration.metersPerUnit * geo.calibration.metersPerUnit
        : 0;
      rooms.push({ id: r.id, page: pageNum, name: r.name, areaM2 });
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
function saveAutosave() {
  if (!currentFileName || !originalBytes) return;
  clearTimeout(autosaveTimer);
  autosaveTimer = setTimeout(() => {
    try {
      localStorage.setItem(
        autosaveKey(currentFileName, originalBytes.length),
        JSON.stringify({ pageGeometry, nextRoomId })
      );
    } catch (e) {
      // Storage full or unavailable (private browsing etc.) — this session
      // still works fine, it just won't be there next time.
    }
  }, 500);
}

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
    zoomIndicatorEl.textContent = `${Math.round((renderScale / 1.5) * 100)}%`;
    prevPageBtn.disabled = currentPageNum <= 1;
    nextPageBtn.disabled = currentPageNum >= currentPdf.numPages;
  });
}

function currentGeometry() {
  if (!pageGeometry[currentPageNum]) {
    pageGeometry[currentPageNum] = { calibration: null, rooms: [] };
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

  geo.rooms.forEach((room) => {
    const isEditing = room.id === editingRoom;
    drawPolygon(room.points, isEditing ? "#c2185b" : "#15655c", room.name);
    if (isEditing) {
      room.points.forEach((pt) => drawHandle(pt, "#c2185b"));
    }
  });

  if (mode === "trace" && traceTemp.page === currentPageNum && traceTemp.points.length > 0) {
    drawPolyline(traceTemp.points, "#c2185b");
  }

  traceRoomBtn.disabled = !geo.calibration;
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
// "drag this" handle on a selected room's corners rather than just a marker.
function drawHandle(p, color) {
  const [cx, cy] = toCanvas(p);
  const size = 9;
  overlayCtx.fillStyle = "#ffffff";
  overlayCtx.fillRect(cx - size / 2, cy - size / 2, size, size);
  overlayCtx.strokeStyle = color;
  overlayCtx.lineWidth = 2;
  overlayCtx.strokeRect(cx - size / 2, cy - size / 2, size, size);
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

function drawPolygon(points, color, label) {
  if (points.length < 3) return;
  overlayCtx.beginPath();
  const [x0, y0] = toCanvas(points[0]);
  overlayCtx.moveTo(x0, y0);
  for (let i = 1; i < points.length; i++) {
    const [x, y] = toCanvas(points[i]);
    overlayCtx.lineTo(x, y);
  }
  overlayCtx.closePath();
  overlayCtx.fillStyle = "rgba(21, 101, 92, 0.15)";
  overlayCtx.fill();
  overlayCtx.strokeStyle = color;
  overlayCtx.lineWidth = 2;
  overlayCtx.stroke();

  const centroid = points.reduce(
    (acc, p) => [acc[0] + p[0] / points.length, acc[1] + p[1] / points.length],
    [0, 0]
  );
  const [cx, cy] = toCanvas(centroid);
  overlayCtx.fillStyle = "#0e453e";
  overlayCtx.font = "13px Segoe UI, Arial, sans-serif";
  overlayCtx.textAlign = "center";
  overlayCtx.fillText(label, cx, cy);
  overlayCtx.textAlign = "left";
}

// ---- Geometry helpers -----------------------------------------------------
function distance(p1, p2) {
  return Math.hypot(p2[0] - p1[0], p2[1] - p1[1]);
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

function canvasPointFromEvent(evt) {
  const rect = overlayCanvas.getBoundingClientRect();
  const cx = evt.clientX - rect.left;
  const cy = evt.clientY - rect.top;
  return [cx / renderScale, cy / renderScale];
}

// Ray-casting point-in-polygon test, used to pick which room a click landed
// on (page-space coordinates on both sides).
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

function hitTestVertex(room, p) {
  const thresholdPageUnits = VERTEX_HIT_RADIUS_PX / renderScale;
  for (let i = 0; i < room.points.length; i++) {
    if (distance(room.points[i], p) <= thresholdPageUnits) return i;
  }
  return -1;
}

// ---- Mode / tool state ----------------------------------------------------
function resetToolState() {
  mode = "idle";
  calibTemp = { p1: null, p2: null };
  traceTemp = { page: null, points: [] };
  editingRoom = null;
  vertexDragState = null;
  scaleRatioForm.hidden = true;
  calibrationForm.hidden = true;
  roomNameForm.hidden = true;
  undoPointBtn.hidden = true;
  finishRoomBtn.hidden = true;
  cancelActionBtn.hidden = true;
}

overlayCanvas.addEventListener("click", (evt) => {
  // A drag (panning the view, or dragging a room's vertex handle — both
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
    finishRoomBtn.disabled = traceTemp.points.length < 3;
    redrawOverlay();
    return;
  }

  if (mode === "idle") {
    const geo = currentGeometry();
    const hit = geo.rooms.slice().reverse().find((r) => pointInPolygon(p, r.points));
    editingRoom = hit ? hit.id : null;
    redrawOverlay();
    setStatus(hit ? `Editing "${hit.name}" — drag its corner handles to reshape it.` : "");
  }
});

// ---- Pan (click-drag) and zoom (mouse wheel) -------------------------------
// Dragging a room's selected vertex handle (set up in mousedown below) takes
// priority over panning for that gesture; anywhere else on the canvas still
// pans as before.
let panState = null;
let suppressNextClick = false;

overlayCanvas.addEventListener("mousedown", (evt) => {
  if (mode === "idle" && editingRoom) {
    const geo = currentGeometry();
    const room = geo.rooms.find((r) => r.id === editingRoom);
    if (room) {
      const idx = hitTestVertex(room, canvasPointFromEvent(evt));
      if (idx !== -1) {
        vertexDragState = { roomId: room.id, pointIndex: idx, dragged: false };
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
  if (vertexDragState) {
    vertexDragState.dragged = true;
    const geo = currentGeometry();
    const room = geo.rooms.find((r) => r.id === vertexDragState.roomId);
    if (room) {
      room.points[vertexDragState.pointIndex] = canvasPointFromEvent(evt);
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
  if (vertexDragState) {
    if (vertexDragState.dragged) {
      recalcAllAreasForPage(currentPageNum);
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
  cancelActionBtn.hidden = false;
  scaleRatioForm.hidden = false;
  setStatus(
    "Click one end of a known measurement on the plan (a scale bar or a labelled dimension), or enter the drawing's printed scale below."
  );
});

traceRoomBtn.addEventListener("click", () => {
  resetToolState();
  mode = "trace";
  traceTemp = { page: currentPageNum, points: [] };
  undoPointBtn.hidden = false;
  finishRoomBtn.hidden = false;
  finishRoomBtn.disabled = true;
  cancelActionBtn.hidden = false;
  setStatus("Click each corner of the room in order, then click “Finish room”.");
});

undoPointBtn.addEventListener("click", () => {
  traceTemp.points.pop();
  finishRoomBtn.disabled = traceTemp.points.length < 3;
  redrawOverlay();
});

cancelActionBtn.addEventListener("click", () => {
  resetToolState();
  setStatus("Cancelled.");
  redrawOverlay();
});

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

scaleRatioConfirmBtn.addEventListener("click", () => {
  const n = parseScaleRatio(scaleRatioInput.value);
  if (!n) {
    setStatus("Enter a scale like 1:100.", true);
    return;
  }
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
  recalcAllAreasForPage(currentPageNum);
  saveAutosave();
  setStatus(`Scale set to 1:${n} for this page. Click “Trace room” to start on the first room.`);
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
  recalcAllAreasForPage(currentPageNum);
  saveAutosave();
  setStatus("Scale set for this page. Click “Trace room” to start on the first room.");
});

calibCancelBtn.addEventListener("click", () => {
  resetToolState();
  redrawOverlay();
});

finishRoomBtn.addEventListener("click", () => {
  if (traceTemp.points.length < 3) return;
  roomNameForm.hidden = false;
  roomNameInput.value = `Room ${rooms.length + 1}`;
  roomNameInput.focus();
  roomNameInput.select();
});

roomNameConfirmBtn.addEventListener("click", () => {
  const geo = currentGeometry();
  if (!geo.calibration) {
    setStatus("Scale isn't set for this page anymore — set it again first.", true);
    return;
  }
  const name = roomNameInput.value.trim() || `Room ${rooms.length + 1}`;
  const areaPageUnits = polygonAreaPageUnits(traceTemp.points);
  const areaM2 = areaPageUnits * geo.calibration.metersPerUnit * geo.calibration.metersPerUnit;

  const id = nextRoomId++;
  geo.rooms.push({ id, name, points: traceTemp.points.slice() });
  rooms.push({ id, page: currentPageNum, name, areaM2 });

  roomNameForm.hidden = true;
  resetToolState();
  redrawOverlay();
  updateResultsTable();
  saveAutosave();
  setStatus(`Added "${name}" — ${areaM2.toFixed(2)} m². Trace another room, or move to the next page.`);
});

roomNameCancelBtn.addEventListener("click", () => {
  roomNameForm.hidden = true;
});

// If the scale is re-set on a page, existing rooms on that page keep their
// traced outlines but their areas are recalculated against the new scale.
function recalcAllAreasForPage(pageNum) {
  const geo = pageGeometry[pageNum];
  if (!geo || !geo.calibration) return;
  geo.rooms.forEach((r) => {
    const areaPageUnits = polygonAreaPageUnits(r.points);
    const areaM2 = areaPageUnits * geo.calibration.metersPerUnit * geo.calibration.metersPerUnit;
    const flat = rooms.find((x) => x.id === r.id);
    if (flat) flat.areaM2 = areaM2;
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

// ---- 3. Results table ----------------------------------------------------
function updateResultsTable() {
  resultsBody.innerHTML = "";
  let totalM2 = 0;
  rooms.forEach((r) => {
    totalM2 += r.areaM2;
    const tr = document.createElement("tr");

    const nameTd = document.createElement("td");
    nameTd.textContent = r.name;

    const pageTd = document.createElement("td");
    pageTd.textContent = r.page;

    const m2Td = document.createElement("td");
    m2Td.textContent = r.areaM2.toFixed(2);

    const delTd = document.createElement("td");
    const delBtn = document.createElement("button");
    delBtn.className = "del-btn";
    delBtn.textContent = "✕";
    delBtn.title = "Remove this room";
    delBtn.addEventListener("click", () => removeRoom(r.id));
    delTd.appendChild(delBtn);

    tr.appendChild(nameTd);
    tr.appendChild(pageTd);
    tr.appendChild(m2Td);
    tr.appendChild(delTd);
    resultsBody.appendChild(tr);
  });

  totalM2El.innerHTML = `<strong>${totalM2.toFixed(2)}</strong>`;
}

function removeRoom(id) {
  rooms = rooms.filter((r) => r.id !== id);
  Object.values(pageGeometry).forEach((geo) => {
    geo.rooms = geo.rooms.filter((r) => r.id !== id);
  });
  if (editingRoom === id) editingRoom = null;
  redrawOverlay();
  updateResultsTable();
  saveAutosave();
}

copyResultsBtn.addEventListener("click", () => {
  let text = "Room\tPage\tm²\n";
  rooms.forEach((r) => {
    text += `${r.name}\t${r.page}\t${r.areaM2.toFixed(2)}\n`;
  });
  const totalM2 = rooms.reduce((s, r) => s + r.areaM2, 0);
  text += `Total\t\t${totalM2.toFixed(2)}\n`;

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
  rooms = [];
  nextRoomId = 1;
  pageGeometry = {};
  resetToolState();
  redrawOverlay();
  updateResultsTable();
  clearAutosave();
  setStatus("Cleared all rooms and scale settings on every page.");
});

// ---- 4. Download PDF with room data embedded -------------------------------
// Builds a modified copy of the original PDF — never the in-memory copy
// pdf.js is using, since getDocument() can transfer/detach that buffer —
// with two things added: the raw geometry as a JSON file attachment (so
// re-opening this exact file restores the exact editable state, no account
// or storage needed) and the traced outlines/labels drawn directly onto the
// pages (so the measurements are visible in any ordinary PDF viewer).
function suggestDownloadName(name) {
  const base = (name || "floor-plan").replace(/\.pdf$/i, "");
  return `${base}-with-rooms.pdf`;
}

// Shared by both build functions below — draws every traced room on `geo`
// onto `pdfLibPage`, using `viewportAtScale1` to map our stored points
// (pdf.js's scale-1 viewport space, see canvasPointFromEvent) back to the
// PDF's own coordinate space (bottom-left origin, correctly accounting for
// page rotation), which is what pdf-lib's drawing calls expect.
function drawRoomAnnotations(pdfLibPage, geo, viewportAtScale1, font, outlineColor) {
  geo.rooms.forEach((room) => {
    if (room.points.length < 3) return;
    const pts = room.points.map((p) => {
      const [x, y] = viewportAtScale1.convertToPdfPoint(p[0], p[1]);
      return { x, y };
    });
    for (let i = 0; i < pts.length; i++) {
      const a = pts[i];
      const b = pts[(i + 1) % pts.length];
      pdfLibPage.drawLine({ start: a, end: b, thickness: 1.5, color: outlineColor });
    }
    const cx = pts.reduce((sum, p) => sum + p.x, 0) / pts.length;
    const cy = pts.reduce((sum, p) => sum + p.y, 0) / pts.length;
    const flatRoom = rooms.find((r) => r.id === room.id);
    const label = flatRoom ? `${room.name} — ${flatRoom.areaM2.toFixed(2)} m²` : room.name;
    pdfLibPage.drawText(label, {
      x: cx - label.length * 2.3,
      y: cy,
      size: 9,
      font,
      color: outlineColor,
    });
  });
}

async function buildAnnotatedPdfBytes() {
  const lib = await loadPdfLib();
  const pdfDoc = await lib.PDFDocument.load(originalBytes.slice());

  const geometryJson = JSON.stringify({ pageGeometry, nextRoomId }, null, 2);
  await pdfDoc.attach(new TextEncoder().encode(geometryJson), "floor-area-takeoff.json", {
    mimeType: "application/json",
    description: "Floor Area Takeoff — scale calibration and traced room outlines",
  });

  const pdfLibPages = pdfDoc.getPages();
  const font = await pdfDoc.embedFont(lib.StandardFonts.Helvetica);
  const outlineColor = lib.rgb(0.08, 0.4, 0.36);

  for (const pageNumKey of Object.keys(pageGeometry)) {
    const pageNum = Number(pageNumKey);
    const geo = pageGeometry[pageNum];
    if (!geo.rooms || geo.rooms.length === 0) continue;
    if (pageNum < 1 || pageNum > pdfLibPages.length) continue;

    const pdfjsPage = await currentPdf.getPage(pageNum);
    const viewportAtScale1 = pdfjsPage.getViewport({ scale: 1 });
    drawRoomAnnotations(pdfLibPages[pageNum - 1], geo, viewportAtScale1, font, outlineColor);
  }

  return pdfDoc.save();
}

function suggestTracedPagesDownloadName(name) {
  const base = (name || "floor-plan").replace(/\.pdf$/i, "");
  return `${base}-traced-pages.pdf`;
}

// A lighter export than buildAnnotatedPdfBytes above: a brand-new document
// containing only the pages that actually have a traced room, for sharing
// just the relevant sheets instead of the whole plan set. Deliberately
// doesn't embed the geometry JSON the way the full export does — once pages
// are dropped, page numbers no longer match the original file, so restoring
// progress from this copy would apply the wrong page's rooms to the wrong
// page. This one is for reading, not round-tripping.
async function buildTracedPagesOnlyPdfBytes() {
  const tracedPageNums = Object.keys(pageGeometry)
    .map(Number)
    .filter((pageNum) => {
      const geo = pageGeometry[pageNum];
      return geo && geo.rooms && geo.rooms.length > 0;
    })
    .sort((a, b) => a - b);

  if (tracedPageNums.length === 0) {
    throw new Error("No traced rooms yet — trace at least one room first.");
  }

  const lib = await loadPdfLib();
  const srcDoc = await lib.PDFDocument.load(originalBytes.slice());
  const outDoc = await lib.PDFDocument.create();
  const copiedPages = await outDoc.copyPages(srcDoc, tracedPageNums.map((n) => n - 1));
  copiedPages.forEach((page) => outDoc.addPage(page));

  const font = await outDoc.embedFont(lib.StandardFonts.Helvetica);
  const outlineColor = lib.rgb(0.08, 0.4, 0.36);

  for (let i = 0; i < tracedPageNums.length; i++) {
    const pageNum = tracedPageNums[i];
    const pdfjsPage = await currentPdf.getPage(pageNum);
    const viewportAtScale1 = pdfjsPage.getViewport({ scale: 1 });
    drawRoomAnnotations(copiedPages[i], pageGeometry[pageNum], viewportAtScale1, font, outlineColor);
  }

  return outDoc.save();
}

// Shared by both download buttons. A real browser tab (unlike the Outlook
// add-in's embedded dialog WebView this replaced) handles a Blob URL + a
// programmatically-clicked <a download> perfectly normally, even after the
// async build work below — this is an ordinary, well-supported web pattern
// here, not the two-phase dance the add-in needed.
function wireDownloadButton(btn, buildFn, filenameFn, preparingStatus) {
  btn.addEventListener("click", (evt) => {
    evt.preventDefault();
    if (!currentPdf || !originalBytes || btn.dataset.busy === "1") return;

    const originalText = btn.textContent;
    btn.dataset.busy = "1";
    btn.textContent = "Preparing…";
    setStatus(preparingStatus);

    buildFn().then(
      (bytes) => {
        const blob = new Blob([bytes], { type: "application/pdf" });
        const url = URL.createObjectURL(blob);
        const filename = filenameFn();
        const a = document.createElement("a");
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 30000);
        btn.textContent = originalText;
        btn.dataset.busy = "";
        setStatus(`Downloaded "${filename}".`);
      },
      (err) => {
        btn.textContent = originalText;
        btn.dataset.busy = "";
        setStatus("Couldn't build the download — " + describeError(err), true);
      }
    );
  });
}

wireDownloadButton(
  downloadPdfBtn,
  buildAnnotatedPdfBytes,
  () => suggestDownloadName(currentFileName),
  "Preparing PDF with room data…"
);
wireDownloadButton(
  downloadTracedPagesBtn,
  buildTracedPagesOnlyPdfBytes,
  () => suggestTracedPagesDownloadName(currentFileName),
  "Preparing traced pages…"
);
