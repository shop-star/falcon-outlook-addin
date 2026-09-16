/* Floor Area Takeoff — Outlook task pane add-in
 * Loads a PDF plan attachment, lets the user calibrate scale from a known
 * measurement, then trace room outlines to get floor areas. Everything runs
 * client-side in the task pane; nothing is uploaded anywhere.
 */

// pdf.js is self-hosted under vendor/pdfjs/ (same origin as this add-in)
// rather than pulled from a public CDN: some corporate networks block
// CDN domains like cdnjs.cloudflare.com from inside the Outlook webview,
// which silently broke PDF loading even though the add-in itself loaded
// fine (it's served from the same origin we already trust). It's also
// loaded lazily, only when a PDF is actually opened, rather than via a
// static top-level import: a failing top-level import would throw before
// Office.onReady ever gets registered, freezing the whole task pane on
// its initial "Looking for PDF attachments…" state with no visible error.
let pdfjsLib = null;
let pdfjsLoadPromise = null;

function loadPdfJs() {
  if (!pdfjsLoadPromise) {
    pdfjsLoadPromise = import("./vendor/pdfjs/pdf.min.mjs").then((mod) => {
      mod.GlobalWorkerOptions.workerSrc = "./vendor/pdfjs/pdf.worker.min.mjs";
      pdfjsLib = mod;
      return mod;
    });
  }
  return pdfjsLoadPromise;
}

const UNIT_TO_M = { mm: 0.001, cm: 0.01, m: 1, ft: 0.3048, in: 0.0254 };
const M2_TO_FT2 = 10.76391;

// ---- State -----------------------------------------------------------
let currentPdf = null;
let currentPageNum = 1;
let renderScale = 1.5;

/** pageGeometry[pageNum] = { calibration: {p1,p2,metersPerUnit,label} | null, rooms: [{id,name,points}] } */
let pageGeometry = {};

/** flat list used for the results table/total, mirrors pageGeometry rooms */
let rooms = []; // { id, page, name, areaM2 }
let nextRoomId = 1;

let mode = "idle"; // idle | calibrate | trace
let calibTemp = { p1: null, p2: null };
let traceTemp = { page: null, points: [] };

// ---- DOM refs ----------------------------------------------------------
const attachmentListEl = document.getElementById("attachmentList");
const statusBarEl = document.getElementById("statusBar");
const viewerSectionEl = document.getElementById("viewerSection");
const resultsSectionEl = document.getElementById("resultsSection");

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
const totalFt2El = document.getElementById("totalFt2");
const copyResultsBtn = document.getElementById("copyResultsBtn");
const clearAllBtn = document.getElementById("clearAllBtn");
const copyFallback = document.getElementById("copyFallback");

// ---- Office.js bootstrap ------------------------------------------------
// Surface any otherwise-silent script error in the status bar instead of
// leaving the task pane stuck on its initial static text with no clue why.
window.addEventListener("error", (evt) => {
  setStatus("Something went wrong: " + (evt.message || evt.error), true);
});
window.addEventListener("unhandledrejection", (evt) => {
  const reason = evt.reason && evt.reason.message ? evt.reason.message : evt.reason;
  setStatus("Something went wrong: " + reason, true);
});

Office.onReady((info) => {
  try {
    if (info.host !== Office.HostType.Outlook) {
      setStatus("This add-in only works inside Outlook.");
      return;
    }
    loadAttachments();
  } catch (e) {
    setStatus("Failed to start: " + e.message, true);
  }
});

function setStatus(msg, isError) {
  statusBarEl.textContent = msg || "";
  statusBarEl.style.color = isError ? "#b3261e" : "";
}

// ---- 1. Attachment list --------------------------------------------------
function loadAttachments() {
  const item = Office.context.mailbox.item;
  if (!item || !item.attachments) {
    attachmentListEl.innerHTML =
      '<p class="muted">No attachments found on this message.</p>';
    return;
  }

  const pdfAttachments = item.attachments.filter((a) => {
    const name = (a.name || "").toLowerCase();
    const type = (a.contentType || "").toLowerCase();
    return a.attachmentType !== "item" && (name.endsWith(".pdf") || type.includes("pdf"));
  });

  if (pdfAttachments.length === 0) {
    attachmentListEl.innerHTML =
      '<p class="muted">No PDF attachments found on this message.</p>';
    return;
  }

  attachmentListEl.innerHTML = "";
  pdfAttachments.forEach((att) => {
    const row = document.createElement("div");
    row.className = "attachment-item";
    const nameSpan = document.createElement("span");
    nameSpan.className = "name";
    nameSpan.textContent = att.name;
    const openBtn = document.createElement("button");
    openBtn.className = "primary";
    openBtn.textContent = "Open";
    openBtn.addEventListener("click", () => openAttachment(att));
    row.appendChild(nameSpan);
    row.appendChild(openBtn);
    attachmentListEl.appendChild(row);
  });
}

function openAttachment(att) {
  setStatus(`Loading "${att.name}"…`);
  loadPdfJs().catch(() => {
    setStatus(
      "Couldn't load the PDF viewer library. Try reloading the add-in.",
      true
    );
  });
  Office.context.mailbox.item.getAttachmentContentAsync(att.id, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus(
        `Couldn't read "${att.name}". If it's a cloud/OneDrive attachment, try downloading and re-attaching the actual PDF file. (${
          result.error ? result.error.message : "unknown error"
        })`,
        true
      );
      return;
    }
    const { format, content } = result.value;
    if (format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
      setStatus("Unexpected attachment format — couldn't read this file as a PDF.", true);
      return;
    }
    loadPdfJs().then(
      () => {
        try {
          const bytes = base64ToUint8Array(content);
          pdfjsLib.getDocument({ data: bytes }).promise.then(
            (pdf) => {
              currentPdf = pdf;
              currentPageNum = 1;
              pageGeometry = {};
              rooms = [];
              nextRoomId = 1;
              renderScale = 1.5;
              viewerSectionEl.hidden = false;
              resultsSectionEl.hidden = false;
              resetToolState();
              renderPage();
              updateResultsTable();
              setStatus(`Loaded "${att.name}". Set the scale, then trace each room.`);
            },
            (err) => setStatus("Couldn't open this PDF: " + err.message, true)
          );
        } catch (e) {
          setStatus("Couldn't decode this attachment as a PDF.", true);
        }
      },
      () => {
        setStatus(
          "Couldn't load the PDF viewer library. Try reloading the add-in.",
          true
        );
      }
    );
  });
}

function base64ToUint8Array(base64) {
  const binaryString = atob(base64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) bytes[i] = binaryString.charCodeAt(i);
  return bytes;
}

// ---- 2. Rendering --------------------------------------------------------
function renderPage() {
  if (!currentPdf) return;
  currentPdf.getPage(currentPageNum).then((page) => {
    const viewport = page.getViewport({ scale: renderScale });
    pdfCanvas.width = overlayCanvas.width = Math.ceil(viewport.width);
    pdfCanvas.height = overlayCanvas.height = Math.ceil(viewport.height);

    page.render({ canvasContext: pdfCtx, viewport }).promise.then(() => {
      redrawOverlay();
    });

    pageIndicatorEl.textContent = `Page ${currentPageNum} / ${currentPdf.numPages}`;
    zoomIndicatorEl.textContent = `${Math.round(renderScale / 1.5 * 100)}%`;
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

  // Calibration line
  if (geo.calibration) {
    drawLine(geo.calibration.p1, geo.calibration.p2, "#e07b00", 2, true, geo.calibration.label);
  }
  if (mode === "calibrate" && calibTemp.p1 && !calibTemp.p2) {
    drawPoint(calibTemp.p1, "#e07b00");
  }

  // Completed rooms
  geo.rooms.forEach((room) => drawPolygon(room.points, "#15655c", room.name));

  // In-progress trace
  if (mode === "trace" && traceTemp.page === currentPageNum && traceTemp.points.length > 0) {
    drawPolyline(traceTemp.points, "#c2185b");
  }

  traceRoomBtn.disabled = !geo.calibration;
  setScaleBtn.textContent = geo.calibration ? "Re-set scale" : "Set scale";
  if (geo.calibration) {
    scaleInfoEl.hidden = false;
    scaleInfoEl.textContent = `Scale on this page: ${geo.calibration.label} = ${geo.calibration.pixelDist.toFixed(
      1
    )} plan units (1 unit ≈ ${(geo.calibration.metersPerUnit * 1000).toFixed(2)} mm).`;
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
    overlayCtx.font = "12px Segoe UI, Arial, sans-serif";
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
  overlayCtx.font = "12px Segoe UI, Arial, sans-serif";
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

// ---- Mode / tool state ----------------------------------------------------
function resetToolState() {
  mode = "idle";
  calibTemp = { p1: null, p2: null };
  traceTemp = { page: null, points: [] };
  calibrationForm.hidden = true;
  roomNameForm.hidden = true;
  undoPointBtn.hidden = true;
  finishRoomBtn.hidden = true;
  cancelActionBtn.hidden = true;
}

overlayCanvas.addEventListener("click", (evt) => {
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
  }
});

setScaleBtn.addEventListener("click", () => {
  resetToolState();
  mode = "calibrate";
  cancelActionBtn.hidden = false;
  setStatus("Click one end of a known measurement on the plan (a scale bar or a labelled dimension).");
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

// ---- Results table ----------------------------------------------------------
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

    const ft2Td = document.createElement("td");
    ft2Td.textContent = (r.areaM2 * M2_TO_FT2).toFixed(2);

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
    tr.appendChild(ft2Td);
    tr.appendChild(delTd);
    resultsBody.appendChild(tr);
  });

  totalM2El.innerHTML = `<strong>${totalM2.toFixed(2)}</strong>`;
  totalFt2El.innerHTML = `<strong>${(totalM2 * M2_TO_FT2).toFixed(2)}</strong>`;
}

function removeRoom(id) {
  rooms = rooms.filter((r) => r.id !== id);
  Object.values(pageGeometry).forEach((geo) => {
    geo.rooms = geo.rooms.filter((r) => r.id !== id);
  });
  updateResultsTable();
  redrawOverlay();
}

copyResultsBtn.addEventListener("click", () => {
  let text = "Room\tPage\tm²\tft²\n";
  rooms.forEach((r) => {
    text += `${r.name}\t${r.page}\t${r.areaM2.toFixed(2)}\t${(r.areaM2 * M2_TO_FT2).toFixed(2)}\n`;
  });
  const totalM2 = rooms.reduce((s, r) => s + r.areaM2, 0);
  text += `Total\t\t${totalM2.toFixed(2)}\t${(totalM2 * M2_TO_FT2).toFixed(2)}\n`;

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
  updateResultsTable();
  redrawOverlay();
  setStatus("Cleared all rooms and scale settings on every page.");
});
