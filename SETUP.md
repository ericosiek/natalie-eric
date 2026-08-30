# Setup & operating guide

Everything you need to run the site, in the order it has to happen.

---

## 1. GitHub Pages

Repo: `ericosiek/natalie-eric` · Branch: `main` · Folder: `/ (root)`

**Settings → Pages**
- Source: **Deploy from a branch**
- Branch: **main**, folder **/ (root)**
- Custom domain: `natalie-eric.website`
- Tick **Enforce HTTPS** (it stays greyed out until the DNS below resolves, usually
  within an hour, sometimes up to 24)

The `CNAME` file in the repo sets the custom domain too, so if the field ever
empties itself, pushing again restores it.

---

## 2. Namecheap DNS

**Domain List → natalie-eric.website → Manage → Advanced DNS**

Delete any parking records Namecheap added by default (usually a CNAME on `www`
pointing at `parkingpage.namecheap.com`, and sometimes an A record on `@`).

Then add exactly these six records:

| Type | Host | Value | TTL |
|---|---|---|---|
| A Record | `@` | `185.199.108.153` | Automatic |
| A Record | `@` | `185.199.109.153` | Automatic |
| A Record | `@` | `185.199.110.153` | Automatic |
| A Record | `@` | `185.199.111.153` | Automatic |
| CNAME Record | `www` | `ericosiek.github.io.` | Automatic |
| ALIAS/other | *(none needed)* | | |

Note the trailing dot on `ericosiek.github.io.` — Namecheap wants it.

Make sure **Nameservers** at the top of that page is set to **Namecheap BasicDNS**,
not "Custom DNS". Nothing else on the page needs touching.

DNS usually propagates in 15 to 60 minutes. `natalie-eric.website` and
`www.natalie-eric.website` will both work; GitHub redirects between them.

---

## 3. The RSVP spreadsheet

Sheet: **Natalie & Eric — Wedding RSVPs** (in Eric's Google Drive)

1. **Extensions → Apps Script**
2. Delete whatever is in `Code.gs`, paste the whole of `rsvp-backend.gs`, save.
3. Run `setup()` once. Google will ask for permission the first time: choose your
   account, click **Advanced → Go to (project) (unsafe)** — that warning is normal
   for a script you wrote yourself — then **Allow**.
4. **Deploy → New deployment → Web app**
   - Description: `RSVP API`
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Deploy, then copy the **Web app URL**. It ends in `/exec`.
5. Paste that URL into `CONFIG.apiUrl` in `index.html` and push.

### Redeploying after a script change

**Deploy → Manage deployments → (pencil icon) → Version: New version → Deploy.**
This keeps the same URL. Creating a *new deployment* instead gives you a new URL
and would break the site, so always edit the existing one.

---

## 4. Adding guests

In the **Guests** tab, one row per person:

| Party | First Name | Last Name | Email | Phone |
|---|---|---|---|---|
| 0037 | Wei | Chen | wei@example.com | |
| 0037 | Mei | Chen | mei@example.com | |
| 0037 | Lily | Chen | | |

Party numbers are text, so keep the leading zeros. Everyone sharing a number can
RSVP for each other.

Then **Wedding → Refresh parties & links** in the spreadsheet menu bar. Each party
gets a row in **Parties** with its own link in column E, and RSVP Open ticked.

### Sending the links

- **Email:** put addresses in the Email(s) column of Parties, then
  **Wedding → Email invites to new parties**. Only parties with an empty
  "Invite Sent" cell get mailed, so it is safe to re-run after adding more guests.
- **WhatsApp / iMessage:** **Wedding → Build invite text**. That creates an
  **Invite Text** tab with a ready-written message per party plus a one-click
  WhatsApp link and an iMessage/SMS link.

---

## 5. Opening and closing RSVPs

- **One party:** tick / untick **RSVP Open** on their row in the **Parties** tab.
- **Everyone:** set `RSVP Open` to `FALSE` on the **Config** tab.

Both take effect immediately. A guest with the page already open still cannot
submit once you close them; the server checks on every save, not just on load.

**Extending the deadline** is just leaving parties ticked, or ticking them back on.
The date on the Config tab is only what the website *says*; it does not lock
anything by itself.

**A second batch of invites later:** add the rows, refresh parties, email them.
Existing parties and their links are never disturbed.

---

## 6. Config tab reference

| Setting | What it does |
|---|---|
| RSVP Open | Master switch. FALSE closes RSVPs for everyone. |
| RSVP Deadline | Blank shows "at your earliest convenience". Otherwise `yyyy-mm-dd`. |
| Meal Options | Comma separated. Changes the dropdown on the site. |
| Site URL | Used to build the invite links. |
| Notify Email | Where RSVP alerts go. Blank falls back to Send As. |
| Notify On RSVP | FALSE stops the alert emails. |
| Send As | The Gmail alias every email goes out from: `natalie.eric.2027@gmail.com`. It must stay a verified alias on the account, otherwise mail silently falls back to the account's own address. |

---

## 7. Stripe registry

Each gift is a Payment Link. They are reusable, so ten people can each buy the
same $50 gift and it simply adds $50 ten times. The honeymoon fund uses Stripe's
customer-chooses-the-amount option.

Links live in the `CONFIG.registry` array in `index.html`. Replace each `"#"` with
the Payment Link URL and push; a tile with `"#"` shows as "Coming soon" rather
than breaking.

---

## 8. Email identity

Everything the site and the script send uses **natalie.eric.2027@gmail.com**, not Eric's
personal address:

- the mailto links on the site (RSVP help, FAQ, registry e-transfer, footer)
- the invitation emails from **Wedding → Email invites to new parties**
- the RSVP notification that lands when a guest replies

The script sends through `GmailApp` with the alias in the `from` field, which Gmail only
permits for a **verified alias**. If that alias is ever removed from the account, mail keeps
sending but reverts to the account's own address, so leave it in place under
**Gmail → Settings → Accounts → Send mail as**.

---

## 9. The registry page

The registry lives at **natalie-eric.website/registry** as its own page
(`registry/index.html`). The home page keeps a short teaser that links to it.

Every gift is defined in the `REGISTRY` object at the top of that file's `<script>`:

```js
{ id:"pot", name:"A cast iron dutch oven", amount:150, art:"pot",
  note:"For the slow Sunday cooking we keep promising ourselves.", url:"#" }
```

- `amount` — a number shows as `$150`; `null` shows as "Any amount" (the two funds).
- `art` — which line illustration to use. The drawings live in the `ART` object just below.
  Available: `plane house toast bed espresso pot knife plates linen crystal suitcase sun whisk`.
- `url` — the Stripe Payment Link. Left as `"#"` the tile shows **Coming soon** rather than a
  dead button, so it is always safe to publish before the links exist.

`REGISTRY.funds` renders as two wide cards at the top; `REGISTRY.gifts` fills the grid below.

### Wiring up Stripe

For each gift, create a **Payment Link** in Stripe (Product → one-time price → Payment link).
Two settings matter:

- **Honeymoon Fund and Home Fund**: turn on *"Let customers choose what they pay"* so any
  amount works.
- **Every link**: set the confirmation page to
  `https://natalie-eric.website/registry/?checkout=success`, which brings the guest back to the
  registry and shows the thank-you note.

Payment Links are reusable by design, so ten people can each buy the same $150 gift and it
simply adds $150 ten times. Nothing sells out.

### Pix

The Pix panel at the bottom is a placeholder. Send the Pix key and the account holder's name
and it becomes a proper panel with the key, a copy button and a QR code.

### Shared stylesheet

Both pages now load `/styles.css`. Editing a colour or a font there changes the whole site at
once. It is the only file both pages depend on, so do not rename it.
