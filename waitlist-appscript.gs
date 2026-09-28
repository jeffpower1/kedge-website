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
//   4. (25 Sep 2026) Accepts kind=survey from survey.html and writes those rows to
//      a "Survey" tab, leaving the waitlist tab and its columns unchanged.
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

// ── Survey (survey.html, catch-up funnel) ───────────────────────────────────
// Rows with kind=survey go to a "Survey" tab (created on first use). Header row
// is written automatically. No health information is asked on the survey page.
var SURVEY_FIELDS = ["name","email","doctor_status","how_long","went_first","story",
  "put_off","how_figure","searched","would_use","trust","who_run","pay","pay_amount","age_band",
  "community","work","gender","contact_ok","phone","anything","source"];

function handleSurvey(p) {
  var ts = (p.ts || new Date().toISOString());
  var email = (p.email || "").toString().slice(0, 200);
  if (!email) { return _out({ ok: false, error: "no email" }); }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Survey");
  if (!sheet) {
    sheet = ss.insertSheet("Survey");
    sheet.appendRow(["timestamp"].concat(SURVEY_FIELDS));
    sheet.setFrozenRows(1);
  }
  var row = [ts];
  for (var i = 0; i < SURVEY_FIELDS.length; i++) {
    row.push((p[SURVEY_FIELDS[i]] || "").toString().slice(0, 1500));
  }
  sheet.appendRow(row);
  var mailStatus = "sent";
  try {
    if (NOTIFY_EMAIL) {
      var body = "Email: " + email + "\nName: " + (p.name || "") + "\nStatus: " + (p.doctor_status || "") +
        "\nWould use: " + (p.would_use || "") + "\nContact OK: " + (p.contact_ok || "") +
        (p.phone ? "\nPhone: " + p.phone : "") + "\nTime: " + ts + "\n\nFull answers are on the Survey tab.";
      var aliases = GmailApp.getAliases();
      var opts = (aliases.indexOf(SEND_FROM) !== -1) ? { from: SEND_FROM, name: "Kedge Health" } : { name: "Kedge Health" };
      GmailApp.sendEmail(NOTIFY_EMAIL, "Kedge survey response" + (p.contact_ok && p.contact_ok !== "no" ? " (wants a call)" : ""), body, opts);
    }
  } catch (mErr) { mailStatus = "MAIL FAILED: " + mErr; }
  return _out({ ok: true, mail: mailStatus });
}

// ── Thank-you email to the registrant (25 Sep 2026) ─────────────────────────
// One email at sign-up, then nothing until there is news. Sent from hello@ when
// that alias is verified, otherwise from the account default. Failure is recorded
// in email_status; the row still saves.
function sendThankYou(email, name) {
  var first = (name || "").trim().split(/\s+/)[0];
  var subject = "You're on the list";
  var body =
    (first ? "Hi " + first + ",\n\n" : "Hi,\n\n") +
    "Thanks for registering with Kedge Health.\n\n" +
    "Here's what happens next. We're opening across Newfoundland and Labrador in stages, " +
    "and the first appointments go to people on this list. We'll email you when you can book, and not before. " +
    "If you told us we could call, the physician behind Kedge will be in touch in the next few weeks.\n\n" +
    "Nothing else to do for now.\n\n" +
    "One thing to say plainly: Kedge is for the routine check-ups that get put off when there's no one to order them. " +
    "If you're unwell today, call 811, or 911 in an emergency.\n\n" +
    "Kedge Health Limited\n" +
    "PO Box 29101 Torbay Rd RPO, St. John's, NL A1A 5B5\n" +
    "hello@kedgehealth.com\n\n" +
    "To come off the list, reply to this email with the word unsubscribe.";
  var aliases = GmailApp.getAliases();
  var opts = { name: "Kedge Health", replyTo: "hello@kedgehealth.com" };
  if (aliases.indexOf(SEND_FROM) !== -1) { opts.from = SEND_FROM; }
  GmailApp.sendEmail(email, subject, body, opts);
}

function doPost(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    if (String(p.kind || "") === "survey") { return handleSurvey(p); }
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

    // Thank-you to the registrant. Best-effort; outcome folded into email_status.
    try { sendThankYou(email, name); mailStatus = mailStatus + " | thank-you sent"; }
    catch (tErr) { mailStatus = mailStatus + " | THANK-YOU FAILED: " + tErr; }

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
