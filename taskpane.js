/* Floor Area Takeoff — Outlook task pane add-in
 * Lists PDF plan attachments on the open email. Opening one fetches its
 * content, then hands off to a separate pop-up window (viewer.html/.js,
 * opened via the Office Dialog API) for the actual PDF viewing, scale
 * calibration, and room tracing — a task pane is too narrow to usefully
 * display a floor plan. This page keeps the attachment picker and the
 * live results table, which stay in sync with the pop-up via messages.
 */

// ---- State -----------------------------------------------------------
let rooms = []; // { id, page, name, areaM2 }
let currentDialog = null;
let lastOpenedAttachment = null; // { name, content } — for "Reopen window"

// ---- DOM refs ----------------------------------------------------------
const attachmentListEl = document.getElementById("attachmentList");
const statusBarEl = document.getElementById("statusBar");
const viewerSectionEl = document.getElementById("viewerSection");
const resultsSectionEl = document.getElementById("resultsSection");
const viewerFileNameEl = document.getElementById("viewerFileName");
const reopenViewerBtn = document.getElementById("reopenViewerBtn");

const resultsBody = document.getElementById("resultsBody");
const totalM2El = document.getElementById("totalM2");
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

// Dev tools/Inspect Element are unreliable in some Outlook clients (notably
// New Outlook for Mac), so error text needs to be readable directly from
// the status bar rather than assuming anyone can open a console.
function describeError(err) {
  if (!err) return "unknown error";
  const name = err.name || "Error";
  const message = err.message || String(err);
  return `${name}: ${message}`;
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

    rooms = [];
    updateResultsTable();
    lastOpenedAttachment = { name: att.name, content };
    viewerFileNameEl.textContent = att.name;
    viewerSectionEl.hidden = false;
    resultsSectionEl.hidden = false;
    openViewerDialog(att.name, content);
  });
}

// ---- 2. Pop-up plan viewer (Office Dialog API) -----------------------------
// The dialog can't be sent the PDF bytes as a URL parameter (way too big),
// and localStorage is documented as unreliable for sharing data between a
// task pane and a dialog specifically in Safari-based hosts (which is what
// Outlook for Mac uses) — so the content is sent over Office's own
// messageChild/messageParent channel instead, split into chunks since
// there's no documented upper size limit to rely on for a multi-MB PDF.
const CHUNK_SIZE = 50000; // characters per messageChild call
let pendingTransfer = null; // { dialog, base64Content, total, nextIndex }

function openViewerDialog(fileName, base64Content) {
  const url = new URL(`viewer.html?v=__CACHEBUST__`, window.location.href).href;
  setStatus(`Opening "${fileName}" in a new window…`);
  Office.context.ui.displayDialogAsync(
    url,
    { height: 80, width: 70, displayInIframe: false },
    (asyncResult) => {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus(
          "Couldn't open the plan viewer window — " + describeError(asyncResult.error),
          true
        );
        return;
      }
      const dialog = asyncResult.value;
      currentDialog = dialog;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) =>
        handleDialogMessage(dialog, arg, fileName, base64Content)
      );
      dialog.addEventHandler(Office.EventType.DialogEventReceived, handleDialogEvent);
    }
  );
}

function handleDialogMessage(dialog, arg, fileName, base64Content) {
  let msg;
  try {
    msg = JSON.parse(arg.message);
  } catch (e) {
    return;
  }

  if (msg.type === "ready") {
    sendPdfToDialog(dialog, fileName, base64Content);
  } else if (msg.type === "chunkAck") {
    if (pendingTransfer && msg.index === pendingTransfer.nextIndex) {
      pendingTransfer.nextIndex += 1;
      sendNextChunk();
    }
  } else if (msg.type === "rooms") {
    rooms = msg.rooms || [];
    updateResultsTable();
  } else if (msg.type === "error") {
    setStatus("Plan viewer — " + msg.message, true);
  }
}

// Chunks are sent one at a time, each waiting for the dialog to acknowledge
// the previous one, rather than firing them all at once — there's no
// documented size or rate limit for messageChild, and firing a burst of
// ~200KB messages produced silent, undetected corruption in testing
// (pdf.js failed with "Invalid PDF structure" even though every expected
// chunk index had arrived). One-at-a-time with small chunks and integrity
// checks on both ends trades a little latency for actually being reliable.
function sendPdfToDialog(dialog, fileName, base64Content) {
  const total = Math.max(1, Math.ceil(base64Content.length / CHUNK_SIZE));
  pendingTransfer = { dialog, base64Content, total, nextIndex: 0 };
  dialog.messageChild(
    JSON.stringify({ type: "start", fileName, total, totalLength: base64Content.length })
  );
  sendNextChunk();
}

function sendNextChunk() {
  if (!pendingTransfer) return;
  const { dialog, base64Content, total, nextIndex } = pendingTransfer;
  if (nextIndex >= total) {
    pendingTransfer = null;
    return;
  }
  const chunk = base64Content.slice(nextIndex * CHUNK_SIZE, (nextIndex + 1) * CHUNK_SIZE);
  dialog.messageChild(
    JSON.stringify({ type: "chunk", index: nextIndex, data: chunk, len: chunk.length })
  );
}

function handleDialogEvent(arg) {
  // 12006 = dialog closed by the user (the only one we need to react to).
  if (arg.error === 12006) {
    currentDialog = null;
    setStatus("Plan viewer window closed.");
  }
}

reopenViewerBtn.addEventListener("click", () => {
  if (!lastOpenedAttachment) return;
  openViewerDialog(lastOpenedAttachment.name, lastOpenedAttachment.content);
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

// Room deletion/clearing is asked of the open dialog (if any) so its traced
// outline disappears from the overlay too, rather than just vanishing from
// this table while a stale shape lingers in the pop-up.
function removeRoom(id) {
  if (currentDialog) {
    currentDialog.messageChild(JSON.stringify({ type: "removeRoom", id }));
    return;
  }
  rooms = rooms.filter((r) => r.id !== id);
  updateResultsTable();
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
  if (currentDialog) {
    currentDialog.messageChild(JSON.stringify({ type: "clear" }));
    return;
  }
  rooms = [];
  updateResultsTable();
  setStatus("Cleared all rooms and scale settings on every page.");
});
