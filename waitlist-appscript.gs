// ============================================================================
// Kedge waitlist backend — Google Apps Script (hardened)
// ============================================================================
// Paste this over the ENTIRE contents of your Apps Script project, then follow
// WAITLIST_EMAIL_FIX.md. What this version does:
//   1. If sending the notification email fails, the error is written into the
//      email_status column instead of vanishing — so you can SEE why per signup.
//   2. sendTestEmail() forces Google's authorization prompt and confirms delivery.
//   3. Captures the newsletter opt-in as a SEPARATE CASL consent, with its own
//      timestamp, in its own column.
//
// SHEET HEADER ROW — set row 1 to exactly these 7 columns (left to right):
//   timestamp | name | email | source | newsletter | newsletter_consent_ts | email_status
// (An earlier version wrote email_status in column E; after this update it moves
//  to column G and E/F become the newsletter fields. A few early TEST rows may be
//  misaligned — safe to clear them, since the list is pre-launch.)
//
// Deliver notifications to whichever inbox you want:
// Alerts go to BOTH your Gmail (which you actively watch) and the kedgehealth
// inbox. Comma-separated recipients are valid. Drop jeff@kedgehealth.com once you
// confirm that mailbox reliably delivers to you.
const NOTIFY_EMAIL = "jeff.power1@gmail.com, jeff@kedgehealth.com";   // the "To"
// Send the alert FROM this address (the "From"). It MUST be a verified
// "Send mail as" alias on the account that runs this script (jeff@). If it
// isn't verified yet, the script falls back to the default sender and says so
// in column E, so nothing breaks. See WAITLIST_EMAIL_FIX / step below.
const SEND_FROM    = "hello@kedgehealth.com";

// ── Brevo push (newsletter drip) ────────────────────────────────────────────
// When someone opts into the newsletter, add them to a Brevo list so the welcome
// + weekly automation starts on its own.
//
// SECURITY — the API key is NOT stored in this file (it must never be committed to
// git, and this file lives in the repo). It lives in Script Properties instead:
//   Apps Script editor → Project Settings (⚙ gear, left sidebar) → Script Properties
//   → Add script property → name: BREVO_API_KEY → value: <your Brevo v3 key> → Save.
// Read it at runtime below. Leave it unset to disable the push (rows still save).
const BREVO_LIST_ID = 5;        // "Steady - newsletter" list (created 2026-07-27)

function getBrevoApiKey() {
  return PropertiesService.getScriptProperties().getProperty("BREVO_API_KEY") || "";
}

// Run once from the editor to confirm the key is set — WITHOUT printing it in full.
function checkBrevoConfigured() {
  var k = getBrevoApiKey();
  Logger.log(k
    ? "BREVO_API_KEY is set (" + k.length + " chars, ends …" + k.slice(-4) + ")."
    : "BREVO_API_KEY is NOT set — add it in Project Settings → Script Properties.");
}

function addNewsletterContactToBrevo(email, name) {
  var BREVO_API_KEY = getBrevoApiKey();
  if (!BREVO_API_KEY || !BREVO_LIST_ID) return "brevo skipped (not configured)";
  try {
    var res = UrlFetchApp.fetch("https://api.brevo.com/v3/contacts", {
      method: "post",
      contentType: "application/json",
      headers: { "api-key": BREVO_API_KEY, "accept": "application/json" },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        email: email,
        attributes: { FIRSTNAME: name },
        listIds: [BREVO_LIST_ID],
        updateEnabled: true   // idempotent: re-adding an existing contact is fine
      })
    });
    var code = res.getResponseCode();
    // 201 created, 204 updated — both success. Brevo uses double opt-in on the list.
    return (code === 201 || code === 204) ? "brevo ok (" + code + ")"
                                          : "brevo error " + code + ": " + res.getContentText().slice(0, 120);
  } catch (bErr) {
    return "brevo failed: " + bErr;
  }
}

function doPost(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    var name   = (p.name   || "").toString().slice(0, 200);
    var email  = (p.email  || "").toString().slice(0, 200);
    var source = (p.source || "").toString().slice(0, 200);
    var ts     = (p.ts     || new Date().toISOString());

    // Newsletter is a SEPARATE CASL consent from the waitlist. It is only "yes"
    // when the user actively ticked the (unticked-by-default) newsletter box, and
    // we stamp the moment they gave it, so the two consents are independently
    // auditable. Waitlist consent = they submitted the form; newsletter consent =
    // newsletter === "yes" with its own timestamp.
    var newsletter   = (String(p.newsletter || "no").toLowerCase() === "yes") ? "yes" : "no";
    var newsletterTs = (newsletter === "yes") ? ts : "";

    if (!email) { return _out({ ok: false, error: "no email" }); }

    // Try the email first so we can record its outcome alongside the row.
    var mailStatus = "sent";
    try {
      if (NOTIFY_EMAIL) {
        var body = "Name: " + name + "\nEmail: " + email + "\nSource: " + source + "\nTime: " + ts;
        // Use hello@ as the From if it's a verified send-as alias; else fall back.
        var aliases = GmailApp.getAliases();
        if (aliases.indexOf(SEND_FROM) !== -1) {
          GmailApp.sendEmail(NOTIFY_EMAIL, "New Kedge waitlist signup", body,
            { from: SEND_FROM, name: "Kedge Health" });
        } else {
          GmailApp.sendEmail(NOTIFY_EMAIL, "New Kedge waitlist signup", body,
            { name: "Kedge Health" });
          mailStatus = "sent (from default — " + SEND_FROM + " not a verified alias yet)";
        }
      } else {
        mailStatus = "no NOTIFY_EMAIL set";
      }
    } catch (mErr) {
      // The email failed — but we still want the signup saved. Record why.
      mailStatus = "MAIL FAILED: " + mErr;
    }

    // Column order (set the header row in the Sheet to match — see legend below):
    // A timestamp | B name | C email | D source | E newsletter | F newsletter_consent_ts | G email_status
    // Newsletter opt-in → push to Brevo so the drip starts. Best-effort; the row
    // still saves either way. Status folded into email_status so failures are visible.
    if (newsletter === "yes") {
      mailStatus = mailStatus + " | " + addNewsletterContactToBrevo(email, name);
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    sheet.appendRow([ts, name, email, source, newsletter, newsletterTs, mailStatus]);

    return _out({ ok: true, mail: mailStatus, newsletter: newsletter });
  } catch (err) {
    return _out({ ok: false, error: String(err) });
  }
}

function _out(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// RUN THIS ONCE from the editor (select sendTestEmail in the toolbar → Run).
// It does two things:
//   • Triggers Google's authorization screen, where you MUST grant the
//     "Send email as you" permission. This is the permission that was almost
//     certainly missing — which is why rows saved but no email came.
//   • Sends you a real test email so you can confirm delivery immediately.
// After it works, redeploy: Deploy → Manage deployments → edit (pencil) →
// Version: New version → Deploy. The live /exec URL does not change.
// ---------------------------------------------------------------------------
function sendTestEmail() {
  var body = "If you can read this in your inbox, the send-email permission is " +
    "granted and waitlist notifications will now arrive. You can delete this.";
  var aliases = GmailApp.getAliases();
  Logger.log("Verified send-as aliases on this account: " + JSON.stringify(aliases));
  if (aliases.indexOf(SEND_FROM) !== -1) {
    GmailApp.sendEmail(NOTIFY_EMAIL, "Kedge test — from " + SEND_FROM, body,
      { from: SEND_FROM, name: "Kedge Health" });
  } else {
    GmailApp.sendEmail(NOTIFY_EMAIL,
      "Kedge test — " + SEND_FROM + " NOT yet a verified alias", body,
      { name: "Kedge Health" });
  }
}
