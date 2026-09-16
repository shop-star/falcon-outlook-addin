/* Floor Area Takeoff — plan viewer dialog
 * Opened by taskpane.js via Office.context.ui.displayDialogAsync so the PDF
 * has a full window to work with instead of the cramped task pane. The task
 * pane sends the PDF's base64 content over in chunks (see the message
 * protocol below); this window renders it, handles scale calibration and
 * room tracing, and reports the current room list back to the task pane
 * after every change so its results table stays in sync live.
 */

function withTimeout(promise, ms, label) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(new Error(`${label} timed out after ${Math.round(ms / 1000)}s`));
    }, ms);
    promise.then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      (err) => {
        clearTimeout(timer);
        reject(err);
      }
    );
  });
}

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

// pdf.js only reads PDFs — pdf-lib is the write side, used to embed the room
// data and burn the traced outlines into a downloadable copy. Self-hosted
// for the same reason as pdf.js (see above): CDN domains get blocked on
// some corporate networks this add-in runs on.
let pdfLib = null;
let pdfLibLoadPromise = null;

function loadPdfLib() {
  if (!pdfLibLoadPromise) {
    pdfLibLoadPromise = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = "./vendor/pdf-lib/pdf-lib.min.js?v=__CACHEBUST__";
      script.onload = () => {
        if (!window.PDFLib) {
          reject(new Error("pdf-lib.min.js loaded but window.PDFLib was not set"));
          return;
        }
        pdfLib = window.PDFLib;
        resolve(pdfLib);
      };
      script.onerror = () => reject(new Error("Failed to load pdf-lib.min.js"));
      document.head.appendChild(script);
    }).catch((err) => {
      pdfLibLoadPromise = null;
      throw err;
    });
  }
  return pdfLibLoadPromise;
}

const UNIT_TO_M = { mm: 0.001, cm: 0.01, m: 1, ft: 0.3048, in: 0.0254 };

// ---- State -----------------------------------------------------------
let currentPdf = null;
let currentPageNum = 1;
let renderScale = 1.5;

// Kept separately from whatever pdf.js does with its own copy of the bytes
// (getDocument() can transfer/detach the buffer it's given) so "Download
// PDF" always has a pristine, untouched original to build from.
let originalBase64Content = null;
let currentFileName = "floor-plan.pdf";

/** pageGeometry[pageNum] = { calibration: {p1,p2,metersPerUnit,label} | null, rooms: [{id,name,points}] } */
let pageGeometry = {};

/** flat list mirrored to the task pane's results table */
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

// ---- Incoming PDF data (chunked from the task pane) -----------------------
let incomingChunks = null; // { fileName, total, parts: [] }

// A restoreGeometry message (sent right after the last PDF chunk) can arrive
// before openPdfFromBase64's async PDF decode has finished setting up fresh
// state — stash it and apply once currentPdf is actually ready.
let pendingRestoreGeometry = null;

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
const downloadPdfBtn = document.getElementById("downloadPdfBtn");
const downloadPdfLink = document.getElementById("downloadPdfLink");

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

// Fast wheel-zooming can still occasionally race two pdf.js render() calls
// on the same canvas despite cancelling the previous task (cancellation
// isn't immediate) — harmless in practice, the next render just takes
// over, but the resulting error was still escaping as an unhandled
// rejection here (not through renderPage's own, more targeted handling)
// and showing a scary-looking message for something that isn't actually
// a problem. Filtered out rather than chasing the underlying race further.
function isBenignRenderRace(text) {
  return typeof text === "string" && text.includes("Cannot use the same canvas");
}

window.addEventListener("error", (evt) => {
  const message = evt.message || String(evt.error);
  if (isBenignRenderRace(message)) return;
  setStatus("Something went wrong: " + message, true);
  notifyParentError(message);
});
window.addEventListener("unhandledrejection", (evt) => {
  const reason = evt.reason && evt.reason.message ? evt.reason.message : evt.reason;
  if (isBenignRenderRace(String(reason))) {
    evt.preventDefault();
    return;
  }
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

// Sends the flat room list (for the task pane's results table) together
// with the full per-page geometry (calibration + traced polygon points) —
// the task pane persists the geometry so scale/traces survive closing and
// reopening this window, and saves it on the email itself so they survive
// closing the task pane entirely.
function notifyParentState() {
  // Any state change invalidates an already-prepared download — otherwise
  // a further edit after "Preparing…" finishes could hand out a save link
  // for a PDF that no longer matches what's on screen.
  resetPreparedDownload();
  Office.context.ui.messageParent(
    JSON.stringify({ type: "state", rooms, geometry: pageGeometry, nextRoomId })
  );
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
    currentFileName = msg.fileName || "floor-plan.pdf";
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
  } else if (msg.type === "downloadChunkAck") {
    if (pendingDownloadSend && msg.index === pendingDownloadSend.nextIndex) {
      pendingDownloadSend.nextIndex += 1;
      sendNextDownloadChunk();
    }
  } else if (msg.type === "restoreGeometry") {
    if (currentPdf) {
      applyRestoredGeometry(msg.geometry, msg.nextRoomId);
    } else {
      // The PDF hasn't finished decoding yet (this message always arrives
      // right after the last chunk, but openPdfFromBase64 resolves
      // asynchronously) — apply it once it has.
      pendingRestoreGeometry = { geometry: msg.geometry, nextRoomId: msg.nextRoomId };
    }
  }
}

// Rebuilds this window's state from a geometry blob the task pane saved
// from a previous time this same attachment was open (either earlier this
// session, or restored from the email itself). The flat room list is
// recomputed from the polygons rather than trusted as sent, so it can never
// drift from what's actually drawn.
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
  notifyParentState();
  if (rooms.length > 0 || Object.keys(pageGeometry).length > 0) {
    setStatus("Restored your previous scale and room traces for this plan.");
  }
}

function openPdfFromBase64(fileName, base64Content) {
  originalBase64Content = base64Content;
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
            downloadPdfBtn.disabled = false;
            renderPage();
            notifyParentState();
            setStatus(`Loaded "${fileName}". Set the scale, then trace each room.`);
            if (pendingRestoreGeometry) {
              const restore = pendingRestoreGeometry;
              pendingRestoreGeometry = null;
              applyRestoredGeometry(restore.geometry, restore.nextRoomId);
            } else {
              // No progress already known for this attachment (from this
              // session or the email's saved state) — check whether the
              // file itself is a previously-downloaded round-trip copy that
              // has its own embedded scale/room data, and restore from that
              // if so. Guarded on pageGeometry still being empty so this
              // never clobbers a restore that arrives in the moment between
              // this check and that promise settling.
              pdf.getAttachments().then(
                (attachments) => {
                  const embedded = attachments && attachments["floor-area-takeoff.json"];
                  if (!embedded || !embedded.content) return;
                  if (Object.keys(pageGeometry).length > 0) return;
                  try {
                    const parsed = JSON.parse(new TextDecoder().decode(embedded.content));
                    if (parsed && parsed.pageGeometry && Object.keys(parsed.pageGeometry).length > 0) {
                      applyRestoredGeometry(parsed.pageGeometry, parsed.nextRoomId);
                    }
                  } catch (e) {
                    // Not our own embedded data, or corrupted — ignore, the
                    // PDF still opens normally either way.
                  }
                },
                () => {}
              );
            }
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

function uint8ArrayToBase64(bytes) {
  // String.fromCharCode.apply on the whole array at once can blow the call
  // stack for a multi-MB PDF — build it up in smaller pieces instead.
  let binary = "";
  const pieceSize = 0x8000;
  for (let i = 0; i < bytes.length; i += pieceSize) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + pieceSize));
  }
  return btoa(binary);
}

// ---- Sending the finished PDF to the task pane ----------------------------
// Mirrors the task pane's own chunked, ack-based send of the original PDF
// (see taskpane.js) in reverse — same reasoning: no documented size/rate
// limit for cross-window Office.js messages, and firing one big message
// produced silent corruption in testing for the original transfer.
const DOWNLOAD_CHUNK_SIZE = 50000;
let pendingDownloadSend = null; // { base64, total, nextIndex }

function sendPdfToTaskPane(bytes, filename) {
  const base64 = uint8ArrayToBase64(bytes);
  const total = Math.max(1, Math.ceil(base64.length / DOWNLOAD_CHUNK_SIZE));
  pendingDownloadSend = { base64, total, nextIndex: 0 };
  Office.context.ui.messageParent(
    JSON.stringify({ type: "downloadStart", fileName: filename, total, totalLength: base64.length })
  );
  sendNextDownloadChunk();
}

function sendNextDownloadChunk() {
  if (!pendingDownloadSend) return;
  const { base64, total, nextIndex } = pendingDownloadSend;
  if (nextIndex >= total) {
    pendingDownloadSend = null;
    return;
  }
  const rawChunk = base64.slice(nextIndex * DOWNLOAD_CHUNK_SIZE, (nextIndex + 1) * DOWNLOAD_CHUNK_SIZE);
  const encoded = encodeURIComponent(rawChunk);
  Office.context.ui.messageParent(
    JSON.stringify({ type: "downloadChunk", index: nextIndex, data: encoded, len: encoded.length })
  );
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
        // describeError(), not err.message directly: if pdf.js rejects with
        // a plain string rather than an Error instance for this specific
        // failure, err.message would be undefined and this filter would
        // silently fail to match — describeError() already falls back to
        // String(err) the same way the two global handlers do.
        if (isBenignRenderRace(describeError(err))) return;
        setStatus("Couldn't render this page — " + describeError(err), true);
        notifyParentError("Couldn't render this page — " + describeError(err));
      }
    );

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
  if (mode === "calibrate" && calibTemp.p1) {
    if (calibTemp.p2) {
      // Both ends clicked — show the line while the length-entry form is
      // open, not just after confirming, so you can see what you measured.
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
  notifyParentState();
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
  notifyParentState();
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
  if (editingRoom === id) editingRoom = null;
  redrawOverlay();
  notifyParentState();
}

function clearAll() {
  rooms = [];
  nextRoomId = 1;
  pageGeometry = {};
  resetToolState();
  redrawOverlay();
  notifyParentState();
  setStatus("Cleared all rooms and scale settings on every page.");
}

// ---- Download PDF with room data embedded ----------------------------------
// Builds a modified copy of the original PDF — never the in-memory copy
// pdf.js is using, since getDocument() can transfer/detach that buffer —
// with two things added: the raw geometry as a JSON file attachment (so
// this tool, or a future companion web app, can restore the exact editable
// state from the file alone, no separate storage needed) and the traced
// outlines/labels drawn directly onto the pages (so the measurements are
// visible in any ordinary PDF viewer, not just here).
function suggestDownloadName(name) {
  const base = (name || "floor-plan").replace(/\.pdf$/i, "");
  return `${base}-with-rooms.pdf`;
}

async function buildAnnotatedPdfBytes() {
  const lib = await loadPdfLib();
  const bytes = base64ToUint8Array(originalBase64Content);
  const pdfDoc = await lib.PDFDocument.load(bytes);

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

    const pdfLibPage = pdfLibPages[pageNum - 1];
    // Our traced points are stored in pdf.js's scale-1 viewport space (see
    // canvasPointFromEvent) — convertToPdfPoint maps that back to the PDF's
    // own coordinate space (bottom-left origin, and correctly accounting
    // for page rotation), which is what pdf-lib's drawing calls expect.
    const pdfjsPage = await currentPdf.getPage(pageNum);
    const viewportAtScale1 = pdfjsPage.getViewport({ scale: 1 });

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
        x: cx - (label.length * 2.3),
        y: cy,
        size: 9,
        font,
        color: outlineColor,
      });
    });
  }

  return pdfDoc.save();
}

// Building the PDF (loading pdf-lib, parsing, embedding, saving) is all
// async, so it can't happen inside a real anchor's click — same-window
// navigation from a click handler with any await in front of it gets
// silently dropped. And a JS-synthesized a.click() on the result, it turns
// out, gets silently dropped too (confirmed by the task pane's identical
// blob-download attempt): the earlier "two fully synchronous clicks" fix
// didn't actually get a download out of Outlook's dialog WebView. What did
// work, in this same WebView, is a plain human click on a real, visible
// <a href> (see taskpane's debug test link) — so this button only builds
// the file; the resulting downloadPdfLink is a real anchor the user has to
// click themselves, with no script ever calling .click() on it.
let preparedDownload = null; // { url, filename }, only set while ready to save

function resetPreparedDownload() {
  if (!preparedDownload) return;
  URL.revokeObjectURL(preparedDownload.url);
  preparedDownload = null;
  downloadPdfLink.hidden = true;
  downloadPdfBtn.hidden = false;
  downloadPdfBtn.disabled = false;
  downloadPdfBtn.textContent = "Download PDF with room data";
}

downloadPdfBtn.addEventListener("click", () => {
  if (!currentPdf || !originalBase64Content) return;

  downloadPdfBtn.disabled = true;
  downloadPdfBtn.textContent = "Preparing…";
  setStatus("Preparing PDF with room data…");
  // If something in here hangs silently instead of rejecting (seen before
  // in this environment — e.g. a script tag whose load/error events never
  // fire), the button would otherwise sit on "Preparing…" forever with no
  // way to tell what went wrong. A hard timeout guarantees a visible error.
  withTimeout(buildAnnotatedPdfBytes(), 20000, "Building the PDF").then(
    (bytes) => {
      const blob = new Blob([bytes], { type: "application/pdf" });
      const filename = suggestDownloadName(currentFileName);
      preparedDownload = { url: URL.createObjectURL(blob), filename };
      downloadPdfLink.href = preparedDownload.url;
      downloadPdfLink.download = filename;
      downloadPdfLink.hidden = false;
      downloadPdfBtn.hidden = true;
      setStatus('PDF ready — click "Click here to save the PDF" to download it.');
      // The task pane is a different embedded surface than this pop-up —
      // worth trying the save from there too, in case it behaves
      // differently (it doesn't — same fix applies there too — but the
      // task pane save link is a useful second copy if this window closes).
      sendPdfToTaskPane(bytes, filename);
    },
    (err) => {
      downloadPdfBtn.disabled = false;
      downloadPdfBtn.textContent = "Download PDF with room data";
      setStatus("Couldn't build the PDF download — " + describeError(err), true);
      notifyParentError("Couldn't build the PDF download — " + describeError(err));
    }
  );
});

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
