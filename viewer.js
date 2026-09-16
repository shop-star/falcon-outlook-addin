/* Floor Area Takeoff — plan viewer dialog
 * Opened by taskpane.js via Office.context.ui.displayDialogAsync so the PDF
 * has a full window to work with instead of the cramped task pane. The task
 * pane sends the PDF's base64 content over in chunks (see the message
 * protocol below); this window renders it, handles scale calibration and
 * room tracing, and reports the current room list back to the task pane
 * after every change so its results table stays in sync live.
 */

// Same self-hosted, older pdf.js build as previously used successfully in
// the task pane (see taskpane.js for why: newer releases, even their
// "legacy" compatibility builds, hit JS-engine incompatibilities in
// Outlook for Mac's embedded WebKit).
let pdfjsLib = null;
let pdfjsLoadPromise = null;

function loadPdfJs() {
  if (!pdfjsLoadPromise) {
    pdfjsLoadPromise = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = "./vendor/pdfjs/pdf.min.js?v=__CACHEBUST__";
      script.onload = () => {
        if (!window.pdfjsLib) {
          reject(new Error("pdf.min.js loaded but window.pdfjsLib was not set"));
          return;
        }
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = "./vendor/pdfjs/pdf.worker.min.js?v=__CACHEBUST__";
        pdfjsLib = window.pdfjsLib;
        resolve(pdfjsLib);
      };
      script.onerror = () => reject(new Error("Failed to load pdf.min.js"));
      document.head.appendChild(script);
    }).catch((err) => {
      pdfjsLoadPromise = null;
      throw err;
    });
  }
  return pdfjsLoadPromise;
}

const UNIT_TO_M = { mm: 0.001, cm: 0.01, m: 1, ft: 0.3048, in: 0.0254 };

// ---- State -----------------------------------------------------------
let currentPdf = null;
let currentPageNum = 1;
let renderScale = 1.5;

/** pageGeometry[pageNum] = { calibration: {p1,p2,metersPerUnit,label} | null, rooms: [{id,name,points}] } */
let pageGeometry = {};

/** flat list mirrored to the task pane's results table */
let rooms = []; // { id, page, name, areaM2 }
let nextRoomId = 1;

let mode = "idle"; // idle | calibrate | trace
let calibTemp = { p1: null, p2: null };
let traceTemp = { page: null, points: [] };

// ---- Incoming PDF data (chunked from the task pane) -----------------------
let incomingChunks = null; // { fileName, total, parts: [] }

// ---- DOM refs ----------------------------------------------------------
const statusBarEl = document.getElementById("statusBar");
const fileNameHeadingEl = document.getElementById("fileNameHeading");

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
const doneBtn = document.getElementById("doneBtn");

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

window.addEventListener("error", (evt) => {
  setStatus("Something went wrong: " + (evt.message || evt.error), true);
  notifyParentError(evt.message || String(evt.error));
});
window.addEventListener("unhandledrejection", (evt) => {
  const reason = evt.reason && evt.reason.message ? evt.reason.message : evt.reason;
  setStatus("Something went wrong: " + reason, true);
  notifyParentError(String(reason));
});

function notifyParentError(message) {
  try {
    Office.context.ui.messageParent(JSON.stringify({ type: "error", message }));
  } catch (e) {
    // If we can't even message the parent, there's nothing more to do.
  }
}

function notifyParentRooms() {
  Office.context.ui.messageParent(JSON.stringify({ type: "rooms", rooms }));
}

// ---- Office.js bootstrap / message handshake with the task pane -----------
Office.onReady(() => {
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onParentMessage);
  // Tell the task pane we're ready to receive the PDF content.
  Office.context.ui.messageParent(JSON.stringify({ type: "ready" }));
});

function onParentMessage(arg) {
  let msg;
  try {
    msg = JSON.parse(arg.message);
  } catch (e) {
    return;
  }

  if (msg.type === "start") {
    incomingChunks = {
      fileName: msg.fileName,
      total: msg.total,
      totalLength: msg.totalLength,
      parts: new Array(msg.total),
      receivedCount: 0,
    };
    fileNameHeadingEl.textContent = msg.fileName || "Floor plan";
    setStatus(`Loading "${msg.fileName}"…`);
  } else if (msg.type === "chunk") {
    if (!incomingChunks) return;
    if (msg.data.length !== msg.len) {
      const err = `Chunk ${msg.index} arrived corrupted (expected ${msg.len} characters, got ${msg.data.length}).`;
      setStatus(err, true);
      notifyParentError(err);
      incomingChunks = null;
      return;
    }
    let rawChunk;
    try {
      // Chunks arrive URL-encoded (see taskpane.js) specifically to keep
      // base64's +, /, = characters off the wire, since a bare "+" turning
      // into a space is a known length-preserving corruption mode.
      rawChunk = decodeURIComponent(msg.data);
    } catch (e) {
      const err = `Chunk ${msg.index} arrived corrupted (couldn't decode: ${e.message}).`;
      setStatus(err, true);
      notifyParentError(err);
      incomingChunks = null;
      return;
    }
    // A plain `new Array(n)` is sparse until every index is explicitly
    // assigned, and Array.prototype.every() silently SKIPS holes in a
    // sparse array rather than visiting them — so checking completeness
    // with parts.every(p => p !== undefined) would return true after just
    // the first chunk (the only "real" element `.every()` could see),
    // regardless of how many holes remained. An explicit counter avoids
    // relying on sparse-array iteration semantics entirely.
    if (incomingChunks.parts[msg.index] === undefined) {
      incomingChunks.receivedCount += 1;
    }
    incomingChunks.parts[msg.index] = rawChunk;
    Office.context.ui.messageParent(JSON.stringify({ type: "chunkAck", index: msg.index }));
    setStatus(
      `Loading "${incomingChunks.fileName}"… (${incomingChunks.receivedCount}/${incomingChunks.total} chunks)`
    );
    if (incomingChunks.receivedCount === incomingChunks.total) {
      const receivedCount = incomingChunks.receivedCount;
      const total = incomingChunks.total;
      const nonEmptyParts = incomingChunks.parts.filter((p) => p !== undefined).length;
      const base64Content = incomingChunks.parts.join("");
      const fileName = incomingChunks.fileName;
      const expectedLength = incomingChunks.totalLength;
      incomingChunks = null;
      if (base64Content.length !== expectedLength) {
        const err =
          `PDF data was corrupted in transit (expected ${expectedLength} characters, got ` +
          `${base64Content.length}; receivedCount=${receivedCount}/${total}, ` +
          `non-empty parts=${nonEmptyParts}/${total}, last chunk index=${msg.index}).`;
        setStatus(err, true);
        notifyParentError(err);
        return;
      }
      openPdfFromBase64(fileName, base64Content);
    }
  } else if (msg.type === "removeRoom") {
    removeRoom(msg.id);
  } else if (msg.type === "clear") {
    clearAll();
  }
}

function openPdfFromBase64(fileName, base64Content) {
  loadPdfJs().then(
    () => {
      try {
        const bytes = base64ToUint8Array(base64Content);
        pdfjsLib.getDocument({ data: bytes }).promise.then(
          (pdf) => {
            currentPdf = pdf;
            currentPageNum = 1;
            pageGeometry = {};
            rooms = [];
            nextRoomId = 1;
            renderScale = 1.5;
            resetToolState();
            setScaleBtn.disabled = false;
            renderPage();
            notifyParentRooms();
            setStatus(`Loaded "${fileName}". Set the scale, then trace each room.`);
          },
          (err) => {
            setStatus("Couldn't open this PDF — " + describeError(err), true);
            notifyParentError("Couldn't open this PDF — " + describeError(err));
          }
        );
      } catch (e) {
        setStatus("Couldn't decode this attachment as a PDF — " + describeError(e), true);
        notifyParentError("Couldn't decode this attachment as a PDF — " + describeError(e));
      }
    },
    (err) => {
      setStatus("Couldn't load the PDF viewer library — " + describeError(err), true);
      notifyParentError("Couldn't load the PDF viewer library — " + describeError(err));
    }
  );
}

function base64ToUint8Array(base64) {
  const binaryString = atob(base64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) bytes[i] = binaryString.charCodeAt(i);
  return bytes;
}

// ---- Rendering --------------------------------------------------------
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

  if (geo.calibration) {
    drawLine(geo.calibration.p1, geo.calibration.p2, "#e07b00", 2, true, geo.calibration.label);
  }
  if (mode === "calibrate" && calibTemp.p1 && !calibTemp.p2) {
    drawPoint(calibTemp.p1, "#e07b00");
  }

  geo.rooms.forEach((room) => drawPolygon(room.points, "#15655c", room.name));

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

// ---- Pan (click-drag) and zoom (mouse wheel) -------------------------------
// Click-to-place-a-point (calibration/tracing, above) and drag-to-pan share
// the same canvas without needing to coordinate explicitly: a real drag
// moves the pointer enough that the browser never fires the "click" event
// afterward, so the point-placing handler simply never sees drags, and a
// plain click never triggers this pan code (dragged stays false, so the
// scroll position is never touched).
let panState = null;

overlayCanvas.addEventListener("mousedown", (evt) => {
  panState = {
    startX: evt.clientX,
    startY: evt.clientY,
    startScrollLeft: canvasScroller.scrollLeft,
    startScrollTop: canvasScroller.scrollTop,
    dragged: false,
  };
});

window.addEventListener("mousemove", (evt) => {
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
  if (panState && panState.dragged) {
    overlayCanvas.style.cursor = "crosshair";
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
    renderScale =
      evt.deltaY < 0 ? Math.min(4, renderScale * 1.1) : Math.max(0.5, renderScale / 1.1);
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
  notifyParentRooms();
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
  notifyParentRooms();
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

// ---- Room removal / clear-all, driven by the task pane's results table ----
// The task pane owns the visible results list; it asks this window to
// remove a room (or clear everything) so the traced outline disappears
// from the overlay too, then this window reports the updated list back.
function removeRoom(id) {
  rooms = rooms.filter((r) => r.id !== id);
  Object.values(pageGeometry).forEach((geo) => {
    geo.rooms = geo.rooms.filter((r) => r.id !== id);
  });
  redrawOverlay();
  notifyParentRooms();
}

function clearAll() {
  rooms = [];
  nextRoomId = 1;
  pageGeometry = {};
  resetToolState();
  redrawOverlay();
  notifyParentRooms();
  setStatus("Cleared all rooms and scale settings on every page.");
}

// A dialog can't close itself: window.close() only works on windows opened
// by script (window.open()) in the same page, and a dialog opened by the
// host application isn't considered "script-opened" from its own point of
// view, so the call silently no-ops. The documented pattern is to ask the
// parent to close it — the task pane holds the real Dialog object (from
// displayDialogAsync's callback), which has a working .close() method.
function requestClose() {
  Office.context.ui.messageParent(JSON.stringify({ type: "closeRequest" }));
}

doneBtn.addEventListener("click", requestClose);

// This pop-up staying on top of other applications when you switch away
// is a documented macOS-specific limitation of the Office.js Dialog API
// itself (tracked upstream in Microsoft's office-js repo) — there's no
// parameter or workaround available from the add-in's own code to fix
// the window layering directly. Auto-closing when the window loses focus
// sidesteps it: once you switch to something else, the window gets out
// of the way instead of floating above everything indefinitely.
//
// Both signals are wired up since this environment's WebKit build has
// repeatedly turned out to support standard web APIs inconsistently
// (missing globals, incorrect array iteration, etc. — see taskpane.js):
// window "blur" alone didn't fire on switching applications, so the Page
// Visibility API is added too, as a differently-implemented alternative
// that some embedded webviews support more reliably than raw focus events
// for OS-level app switching specifically.
window.addEventListener("blur", requestClose);
document.addEventListener("visibilitychange", () => {
  if (document.hidden) requestClose();
});
