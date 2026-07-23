# kedge-website

The public marketing site for **Kedge Health** — a single static page served from
GitHub Pages at **kedgehealth.com**.

Deliberately separate from the clinical repositories (`ClearPath`, `kedge-android`,
`PrimScreenNL`) for privacy and audit hygiene: nothing here touches patient data or
clinical code, and the public site has no access to any clinical system.

## Contents

| File | What it is |
|---|---|
| `index.html` | The whole site — one self-contained file (HTML + CSS + JS inline). |
| `CNAME` | Tells GitHub Pages to serve at `kedgehealth.com`. |
| `WEBSITE_GO_LIVE.md` | **Step-by-step**: create the repo, enable Pages, set DNS, wire the waitlist backend. |

## The one thing to set before launch

The waitlist form posts to `WAITLIST_ENDPOINT`, a single constant near the bottom of
`index.html` (currently blank). Set it to your deployed backend URL — see
`WEBSITE_GO_LIVE.md`. While it's blank the form still shows success but stores
nothing.

## Editing

It's one file. Open `index.html`, edit, commit, push — GitHub Pages redeploys in a
minute or two. No build step, no dependencies.

## Privacy note

The waitlist is **marketing data only** (first name + email). It is kept in the
founder's own Google Workspace, separate from every clinical system, and holds no
personal health information.
