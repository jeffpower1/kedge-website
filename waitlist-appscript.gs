// ============================================================================
// Kedge waitlist backend — Google Apps Script (hardened)
// ============================================================================
// Paste this over the ENTIRE contents of your Apps Script project, then follow
// WAITLIST_EMAIL_FIX.md. The key differences from the first version:
//   1. If sending the notification email fails, the error is written into a new
//      column E of the sheet ("email_status") instead of vanishing — so you can
//      SEE why no email arrived, per signup.
//   2. sendTestEmail() lets you force Google's authorization prompt and confirm
//      email delivery in one click, without waiting for a real signup.
//
// Deliver notifications to whichever inbox you want:
const NOTIFY_EMAIL = "jeff@kedgehealth.com";   // change if you prefer jeff.power1@gmail.com

function doPost(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    var name   = (p.name   || "").toString().slice(0, 200);
    var email  = (p.email  || "").toString().slice(0, 200);
    var source = (p.source || "").toString().slice(0, 200);
    var ts     = (p.ts     || new Date().toISOString());

    if (!email) { return _out({ ok: false, error: "no email" }); }

    // Try the email first so we can record its outcome alongside the row.
    var mailStatus = "sent";
    try {
      if (NOTIFY_EMAIL) {
        MailApp.sendEmail(
          NOTIFY_EMAIL,
          "New Kedge waitlist signup",
          "Name: " + name + "\nEmail: " + email + "\nSource: " + source + "\nTime: " + ts
        );
      } else {
        mailStatus = "no NOTIFY_EMAIL set";
      }
    } catch (mErr) {
      // The email failed — but we still want the signup saved. Record why.
      mailStatus = "MAIL FAILED: " + mErr;
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    sheet.appendRow([ts, name, email, source, mailStatus]);   // column E = email_status

    return _out({ ok: true, mail: mailStatus });
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
  MailApp.sendEmail(
    NOTIFY_EMAIL,
    "Kedge test email — MailApp is working",
    "If you can read this in your inbox, the send-email permission is granted " +
    "and waitlist notifications will now arrive. You can delete this."
  );
}
