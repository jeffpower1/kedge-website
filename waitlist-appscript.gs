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
const NOTIFY_EMAIL = "jeff@kedgehealth.com";   // where the alert is delivered (the "To")
// Send the alert FROM this address (the "From"). It MUST be a verified
// "Send mail as" alias on the account that runs this script (jeff@). If it
// isn't verified yet, the script falls back to the default sender and says so
// in column E, so nothing breaks. See WAITLIST_EMAIL_FIX / step below.
const SEND_FROM    = "hello@kedgehealth.com";

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
