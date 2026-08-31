# Natalie & Eric — how the website works

Everything lives in three places.

| Piece | Where | What it does |
|---|---|---|
| The website | GitHub repo `ericosiek/natalie-eric`, served free by GitHub Pages at **natalie-eric.website** | The pages guests see |
| The backend | A Google Apps Script attached to the RSVP spreadsheet | Answers the website, sends the emails, keeps the log |
| The data | That same Google Sheet | Guests, parties, schedule, FAQ, meals, settings, change log |

You never have to touch any of it directly. The admin portal covers day-to-day work,
and Claude covers changes to the pages themselves.

---

## The admin portal

**natalie-eric.website/admin** — username `wedding`, password `venables2026`.

A sign-in lasts 12 hours. Six wrong passwords in a row locks the door for 15 minutes.

### Guests & parties
Every household is a **party** with a four-digit number and its own private link.
Open a party card to edit names, emails, phones, RSVP answers, meals and dietary notes,
or to move someone into a different party by typing a different number in their Party box.

**Everything saves itself** a moment after you stop typing — there is no Save button. The label
on the right of each card header tells you where things stand: *Unsaved changes*, *Saving…*,
*All changes saved*. Adding a guest saves what is already on screen first, so nothing you have
typed is ever lost. The party's email box keeps itself in step with the guest emails as you type
them, unless you type something different in it yourself.

Each card also has:

- **Invite wave** — a number. Wave 1 goes out first, wave 2 later. Cosmetic grouping for you.
- **RSVP deadline** — optional, per party. Overrides the global one in Settings.
- **RSVP open** — the lock. Closed means the website will not accept their reply and,
  importantly, they cannot be found by the last-name search at all.
- **Private note** — never shown to guests.
- **Send invitation** / **Send reminder** — jumps to the Email tab with that party selected.

New parties are created **closed**. Sending the invitation is what opens them.

### Email
Pick a template (invitation, reminder, thank-you) or write your own. The tokens are:

- `{{names}}` first names of the party
- `{{name}}` party name
- `{{link}}` their own RSVP link
- `{{deadline}}` their deadline

Put `[[Open Invitation]]` on its own line and it becomes a button linking to their party.

The preview on the right is rendered by the server using the first party you have ticked,
so what you see is exactly what lands in their inbox — real names, real link.

Select parties with the filter buttons: all, none, not yet invited, invited but no reply,
missing an email.

Everything is sent from **natalie.eric.2027@gmail.com**.

### Site content
The schedule, the meal choices and the FAQ. Save, and the website picks them up on the
next page load. Untick **Show** to hide a row without deleting it.

The meal list here and **Meal Options** in Settings are the same list shown two ways — change it
in either place and the other follows.

### Change log
Every change to an RSVP, by a guest or by you, with what it was and what it became.

### Settings
Global switches. The one that matters most is **RSVP Open** — a master off switch that
closes RSVPs for everyone regardless of the individual party settings.

The admin password is deliberately **not** editable here. It lives in the Apps Script
project (Project Settings → Script properties → `ADMIN_PASSWORD`).

---

## Where the RSVP links come from

Every party has its own link, like `natalie-eric.website/?i=kthcxhnx`. The eight characters
after `?i=` are that party's **link code** — a random code, not a guessable number.

**They are made for you. You never create one.** A party gets its code the moment the party
row first exists, whichever way it came about:

- you clicked **Add a party** in the portal, or
- you typed a row into the **Parties** tab of the sheet, or
- you typed a guest into the **Guests** tab with a party number that had no row yet, and one
  was created for them.

The code is written into the `Link Code` column and the full link into `Invite Link`, and both
appear on the party's card in the portal with a **Copy** button. Once made, a code never changes
— so a link you sent in January still works in June. Deleting a party retires its code with it.

The link is the whole key. Anyone holding it can open that party and reply for everyone in it,
which is deliberate: one person can RSVP for the household, and nobody needs a password. It also
means the links are worth sending directly to each party rather than posting anywhere public.

Guests who lose the link are not stuck — the RSVP section lets them search their last name, and
they pick their household from the results. That search only ever finds parties whose RSVP is
**open**.

---

## Adding people quickly, straight in the sheet

For a long list, typing into the **Guests** tab is much faster than the portal. Fill in as little
as the party number, first name and last name; everything else can wait.

Then open the portal and press **Reload from sheet**. Anything missing is filled in for you:
guest IDs, party rows for any party number that didn't have one, party names, link codes and
invite links. A party number typed as `7` becomes `0007`. New parties arrive **closed**, exactly
like ones made in the portal, so nothing you type can accidentally go live before you send it.

---

## Invite waves, and why they stay private

A guest who has not been invited yet searches their last name and gets the same
"we couldn't find that" as someone who mistyped. They cannot tell that other people
were invited first, because closed parties are invisible to the search — not hidden,
genuinely absent from the answer.

So a later wave is just: add the parties, leave them closed, and send their invitations
when you're ready.

---

## Making changes to the pages

Ask Claude. For example:

- "Change the cocktail hour to 5:30 and push it live"
- "Add a question to the FAQ about parking"
- "Make the invitation say something different"

Claude edits the files, commits them and pushes. GitHub Pages rebuilds in about a minute.

Small things — schedule times, FAQ wording, meal options — are faster to do yourself in
the admin portal's Site content tab.

---

## If something needs doing by hand

**Redeploying the backend after a code change.** In the Apps Script editor:
Deploy → Manage deployments → pencil icon → Version: **New version** → Deploy.
That keeps the same URL. Choosing "New deployment" instead mints a new URL and breaks
the site until the new one is pasted into `index.html` and `admin/index.html`.

**The sheet.** "Open the sheet" in the admin portal takes you there. Editing it directly
works — the website reads it live — but the admin portal is safer because it logs
what changed.

---

## Files in the repo

```
index.html          the main page: invitation, schedule, RSVP, FAQ
registry/index.html the Honeymoon Fund page
admin/index.html    the admin portal
styles.css          shared styling
rsvp-backend.gs      the Apps Script source, kept here so changes are tracked
CNAME               the custom domain
```

`rsvp-backend.gs` in the repo is the source of truth for the backend. When it changes it
has to be pasted into the Apps Script editor and redeployed — the repo copy is not live
by itself.
