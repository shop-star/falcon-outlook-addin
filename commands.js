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
 * attachment names (the user still saves these from Outlook's own UI and
 * drops them into the app — no attachment bytes flow through here), and any
 * links found in the email body.
 */

const APP_URL = "https://shop-star.github.io/falcon-outlook-addin/app/index.html";
const MAX_ATTACHMENTS = 10;
const MAX_LINKS = 10;

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

  const attachments = collectAttachmentNames(item);

  item.body.getAsync(Office.CoercionType.Text, (result) => {
    const bodyText = result.status === Office.AsyncResultStatus.Succeeded ? result.value : "";
    launchApp({
      subject: item.subject || "",
      attachments,
      links: extractLinks(bodyText),
    });
    event.completed();
  });
});

// Only names, never bytes — the user saves the actual file from Outlook's
// own attachment UI and drops it into the app themselves.
function collectAttachmentNames(item) {
  if (!item.attachments) return [];
  return item.attachments
    .filter((a) => {
      if (a.attachmentType === "item") return false;
      const name = (a.name || "").toLowerCase();
      const type = (a.contentType || "").toLowerCase();
      return name.endsWith(".pdf") || name.endsWith(".zip") || type.includes("pdf") || type.includes("zip");
    })
    .map((a) => a.name)
    .filter(Boolean)
    .slice(0, MAX_ATTACHMENTS);
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

function launchApp({ subject, attachments, links }) {
  const params = new URLSearchParams();
  if (subject) params.set("subject", subject);
  attachments.forEach((name) => params.append("attachment", name));
  links.forEach((url) => params.append("link", url));

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
