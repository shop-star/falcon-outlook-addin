/* Floor Area Takeoff — ribbon command (the whole add-in)
 *
 * This used to be a task pane plus a pop-up plan viewer, both running
 * inside Outlook's embedded dialog/task-pane WebView. That WebView turned
 * out to never reliably support triggering a file download, across every
 * approach tried (see the project history) — and the same class of problem
 * would only get worse trying to add OAuth popups for Dropbox/OneDrive
 * picking. So the actual plan viewer, tracing, and export now live in a
 * standalone browser app (see app/), which a real browser tab handles
 * natively with none of those restrictions.
 *
 * This file's only job is to open that app in a real browser window via
 * Office.context.ui.openBrowserWindow — an API Microsoft ships specifically
 * for "things the sandboxed add-in WebView can't do, like downloading a
 * file" — and hand it a head start: the open email's subject, its PDF/zip
 * attachment names, and any links found in the body. When there's exactly
 * one plan-like attachment and it's a reasonable size, its bytes are also
 * put on the clipboard as a data: URI (see copyAttachmentToClipboard) so
 * the app can pick it up from a paste — skipping the "save from Outlook,
 * then drag into the browser" round trip for the common one-plan case.
 * Anything more ambiguous (several attachments, or a link) still needs a
 * manual save — reading attachment bytes here is fine (that side of
 * Office.js has always worked), it's only ever getting bytes *out* to an
 * arbitrary destination that this WebView can't do reliably.
 */

const APP_URL = "https://shop-star.github.io/falcon-outlook-addin/app/index.html";
const MAX_ATTACHMENTS = 10;
const MAX_LINKS = 10;
// Comfortably covers a typical floor plan PDF while keeping the clipboard
// payload (base64 runs ~33% bigger than the original) from ballooning into
// something that chokes the clipboard API or the page reading it back.
const MAX_CLIPBOARD_BYTES = 8 * 1024 * 1024;

Office.onReady();

Office.actions.associate("openFloorAreaApp", (event) => {
  const item = Office.context.mailbox.item;
  if (!item) {
    event.completed();
    return;
  }

  if (!Office.context.requirements.isSetSupported("OpenBrowserWindowApi", "1.1")) {
    showErrorNotification(
      `This version of Outlook can't open the Floor Area Takeoff app directly — open ${APP_URL} in your browser instead.`
    );
    event.completed();
    return;
  }

  const attachments = collectRelevantAttachments(item);

  item.body.getAsync(Office.CoercionType.Text, (result) => {
    const bodyText = result.status === Office.AsyncResultStatus.Succeeded ? result.value : "";
    const context = {
      subject: item.subject || "",
      attachments: attachments.map((a) => a.name),
      links: extractLinks(bodyText),
    };

    const single = attachments.length === 1 ? attachments[0] : null;
    if (single && single.size && single.size <= MAX_CLIPBOARD_BYTES) {
      copyAttachmentToClipboard(item, single, (ok) => {
        if (ok) context.clipboardReady = single.name;
        launchApp(context);
        event.completed();
      });
      return;
    }

    launchApp(context);
    event.completed();
  });
});

// Never more than one, and only when it's the sole plan-like attachment —
// with several attachments there's no good way to guess which one the
// clipboard hand-off should carry, so that case just falls back to the
// plain name list (still handled by launchApp below).
function collectRelevantAttachments(item) {
  if (!item.attachments) return [];
  return item.attachments
    .filter((a) => {
      if (a.attachmentType === "item") return false;
      const name = (a.name || "").toLowerCase();
      const type = (a.contentType || "").toLowerCase();
      return name.endsWith(".pdf") || name.endsWith(".zip") || type.includes("pdf") || type.includes("zip");
    })
    .filter((a) => a.name)
    .slice(0, MAX_ATTACHMENTS);
}

// A data: URI carries the file's own bytes, so the app can load it straight
// from a paste without needing to fetch anything or share storage across
// the Outlook WebView / real-browser boundary (the two don't share
// localStorage or a filesystem). Any failure here — read, permission,
// clipboard access — just means no clipboard hand-off; the caller still
// launches the app with the plain attachment-name list either way.
function copyAttachmentToClipboard(item, attachment, done) {
  if (!navigator.clipboard || !navigator.clipboard.writeText) {
    done(false);
    return;
  }
  item.getAttachmentContentAsync(attachment.id, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      done(false);
      return;
    }
    const { format, content } = result.value;
    if (format !== Office.MailboxEnums.AttachmentContentFormat.Base64 || !content) {
      done(false);
      return;
    }
    const mimeType = attachment.name.toLowerCase().endsWith(".zip") ? "application/zip" : "application/pdf";
    navigator.clipboard.writeText(`data:${mimeType};base64,${content}`).then(
      () => done(true),
      () => done(false)
    );
  });
}

// A plain http(s) scan — no attempt to guess which links are "file-sharing"
// links specifically, the app just lists them for a one-click open in the
// user's own browser (which carries their normal login session for
// SharePoint/Dropbox/etc., unlike anything fetched from inside Outlook).
const URL_PATTERN = /https?:\/\/[^\s<>"')]+/gi;
function extractLinks(text) {
  const found = text.match(URL_PATTERN) || [];
  const seen = new Set();
  const unique = [];
  for (const raw of found) {
    const url = raw.replace(/[.,;:!?]+$/, "");
    if (seen.has(url)) continue;
    seen.add(url);
    unique.push(url);
    if (unique.length >= MAX_LINKS) break;
  }
  return unique;
}

function launchApp({ subject, attachments, links, clipboardReady }) {
  const params = new URLSearchParams();
  if (subject) params.set("subject", subject);
  attachments.forEach((name) => params.append("attachment", name));
  links.forEach((url) => params.append("link", url));
  if (clipboardReady) params.set("clipboardReady", clipboardReady);

  const url = params.toString() ? `${APP_URL}?${params.toString()}` : APP_URL;
  Office.context.ui.openBrowserWindow(url);
}

// A function command has no task pane or dialog of its own to show status
// in — a notification bar on the reading pane is the one UI surface
// available for telling the user something went wrong.
function showErrorNotification(message) {
  try {
    Office.context.mailbox.item.notificationMessages.addAsync("floorAreaTakeoff", {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message,
    });
  } catch (e) {
    // No UI surface left to report this on.
  }
}
