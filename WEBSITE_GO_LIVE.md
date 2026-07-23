# Website go-live — step by step

Getting **kedgehealth.com** live on GitHub Pages, and wiring the waitlist so signups
land in your own Google Workspace. Follow top to bottom. Nothing here needs code —
it's clicks, a copy-paste script, and DNS records.

---

## Part 1 — Put the site on GitHub (10 minutes)

1. **Create the repo.** On github.com (your `jeffpower1` account) → **New repository**
   → name it **`kedge-website`** → Public → **Create**. Don't add a README (this
   folder already has one).

2. **Push this folder.** In Terminal:
   ```
   cd ~/health_screening_app/kedge-website
   git remote add origin https://github.com/jeffpower1/kedge-website.git
   git branch -M main
   git push -u origin main
   ```
   *(The folder is already a git repo with commits — see the bottom of this doc.)*

3. **Enable GitHub Pages.** Repo → **Settings → Pages** →
   - **Source:** Deploy from a branch
   - **Branch:** `main`, folder `/ (root)` → **Save**.
   - Wait ~1 minute. It'll show a URL like `https://jeffpower1.github.io/kedge-website/`.
   - The `CNAME` file in the repo already tells Pages your custom domain is
     `kedgehealth.com`; once DNS (Part 2) resolves, Pages serves there.

---

## Part 2 — Point kedgehealth.com at GitHub Pages (DNS)

Do this at **whatever registrar holds kedgehealth.com** (GoDaddy, Namecheap, Google
Domains/Squarespace, Cloudflare — the record *values* are identical everywhere; only
the UI differs). You're adding **four A records** for the apex and **one CNAME** for
www.

**Apex domain `kedgehealth.com` — four A records** (GitHub Pages' IPs):

| Type | Host / Name | Value |
|---|---|---|
| A | `@` | `185.199.108.153` |
| A | `@` | `185.199.109.153` |
| A | `@` | `185.199.110.153` |
| A | `@` | `185.199.111.153` |

**www subdomain — one CNAME:**

| Type | Host / Name | Value |
|---|---|---|
| CNAME | `www` | `jeffpower1.github.io` |

Notes:
- Some registrars write the host as `@` for the apex, some leave it blank, some want
  the full `kedgehealth.com`. They all mean "the root domain."
- If your registrar is **Cloudflare**, set the records to **DNS only** (grey cloud,
  not proxied) for the initial Pages HTTPS handshake, then you can proxy later.
- Delete any old parking-page A records or a conflicting apex CNAME first.

**Back in GitHub → Settings → Pages → Custom domain:** type `kedgehealth.com` →
**Save**. Leave **Enforce HTTPS** unchecked *until* the certificate is issued (see
next).

---

## Part 3 — HTTPS

Once the A records resolve, GitHub automatically requests a free Let's Encrypt
certificate. This can take anywhere from a few minutes to ~24 hours the first time.
When Settings → Pages stops showing the "certificate being issued" notice, **tick
"Enforce HTTPS."** After that, `http://` visitors are redirected to `https://`.

**Propagation expectations:** DNS changes typically take 30 minutes to a few hours,
occasionally up to 24–48 hours to fully propagate. If the site 404s or shows a
certificate warning in the first hour, that's normal — wait and refresh. You can
check propagation at dnschecker.org (look up `kedgehealth.com`, type A).

---

## Part 4 — Waitlist backend (this is the one functional piece)

The form in `index.html` posts to a single constant, `WAITLIST_ENDPOINT` (near the
bottom of the file, currently blank). Point it at a backend that captures signups.

### Recommended: Google Apps Script → a Sheet in your Workspace (free, no third party)

Signups land in **your own Google Sheet**, and you get an **email** for each one.
Nothing leaves your Workspace.

**a. Make the Sheet.** In your Kedge Google Workspace, create a new Google Sheet,
name it `Kedge Waitlist`. In row 1 put headers: `timestamp`, `name`, `email`,
`source`.

**b. Add the script.** In that Sheet: **Extensions → Apps Script**. Delete the
placeholder and paste this:

```javascript
// Kedge waitlist backend. Deploy as a Web App (see steps below).
// Writes each signup to this Sheet and emails you a notification.
const NOTIFY_EMAIL = "jeff.power1@gmail.com"; // where signup alerts go

function doPost(e) {
  try {
    const p = (e && e.parameter) ? e.parameter : {};
    const name  = (p.name  || "").toString().slice(0, 200);
    const email = (p.email || "").toString().slice(0, 200);
    const source = (p.source || "").toString().slice(0, 200);
    const ts = (p.ts || new Date().toISOString());

    if (!email) { return _out({ ok: false, error: "no email" }); }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    sheet.appendRow([ts, name, email, source]);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(
        NOTIFY_EMAIL,
        "New Kedge waitlist signup",
        "Name: " + name + "\nEmail: " + email + "\nSource: " + source + "\nTime: " + ts
      );
    }
    return _out({ ok: true });
  } catch (err) {
    return _out({ ok: false, error: String(err) });
  }
}

function _out(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
```

**c. Deploy it (this is the 5-minute part).** In the Apps Script editor:
- **Deploy → New deployment** → gear icon → **Web app**.
- **Description:** `Kedge waitlist`.
- **Execute as:** **Me** (your account).
- **Who has access:** **Anyone**. *(This means anyone can POST a signup — which is
  the point of a public form. It does not expose your Sheet; only `doPost` runs.)*
- **Deploy.** Google asks you to **Authorize** — approve it (it'll warn it's an
  unverified app you made yourself; continue).
- Copy the **Web app URL** — it looks like
  `https://script.google.com/macros/s/AKfyc……/exec`.

**d. Wire it in.** Open `index.html`, find `var WAITLIST_ENDPOINT = "";`, and paste
your URL between the quotes. Commit and push. Done — signups now append to the Sheet
and email you.

> If you change the script later, use **Deploy → Manage deployments → edit → new
> version**, or the URL won't pick up your changes.

### Alternative: Formspree (zero setup, third-party)

If you'd rather not touch Apps Script: create a free form at formspree.io, and it
gives you an endpoint like `https://formspree.io/f/xxxxxxx`. Paste that into
`WAITLIST_ENDPOINT` instead. Signups then live in Formspree's dashboard (a third
party) rather than your own Workspace — the trade-off is convenience vs. keeping the
data in-house. The form code already sends `name` and `email`, which Formspree
accepts as-is.

---

## Privacy posture (already built into the page)

- The form shows a one-line consent note: joining means you may email them about the
  launch, and the list is **marketing data only, separate from any clinical system,
  with no health information**.
- Keep it that way: the waitlist Sheet is marketing data. Do not merge it with, or
  store it near, anything from the clinical platform. It is not PHI and must not
  become PHI-adjacent.

---

## Quick checklist

- [ ] `kedge-website` repo created and pushed
- [ ] Pages enabled (branch `main`, root)
- [ ] Four A records + one www CNAME set at the registrar
- [ ] Custom domain `kedgehealth.com` saved in Pages settings
- [ ] Certificate issued → **Enforce HTTPS** ticked
- [ ] Apps Script deployed, URL pasted into `WAITLIST_ENDPOINT`, pushed
- [ ] Tested: submit the form, confirm a row in the Sheet + an email
