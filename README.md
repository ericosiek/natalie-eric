# natalie-eric.website

The wedding website for Natalie & Eric, July 11 2027, Riverway Clubhouse, Burnaby BC.
Hosted free on GitHub Pages at https://natalie-eric.website

## Files

| File | What it is |
|---|---|
| `index.html` | The home page: invitation, welcome, schedule, RSVP, FAQ. |
| `registry/index.html` | The registry page at /registry, with the gift tiles and illustrations. |
| `styles.css` | The shared stylesheet for both pages. |
| `CNAME` | Tells GitHub Pages to serve the site at natalie-eric.website. Do not delete. |
| `rsvp-backend.gs` | The Google Apps Script that lives in the RSVP spreadsheet. Not used by the website directly; it is kept here as the source of truth. |

## How the RSVP system works

1. Every guest is one row in the **Guests** tab of the RSVP Google Sheet.
2. The **Party** column groups them. Everyone with Party `0037` is one party.
3. The **Parties** tab has one row per party with a unique **Link Code** and the
   **Invite Link** built from it, e.g. `https://natalie-eric.website/?i=k7x2m9pq`.
4. A guest opens their link, the site loads their whole party, and any of them can
   accept or decline and pick a meal for everyone in it.
5. Replies write straight back into the Guests tab.

### Controlling who can still change their RSVP

- **One party:** tick or untick **RSVP Open** on that party's row in the Parties tab.
  Ticked means they can still edit. Unticked means the site shows them their answers
  read-only, and the server refuses any submission even from a page that was already open.
- **Everyone:** set `RSVP Open` to `FALSE` on the **Config** tab.

Changes take effect immediately. Nothing needs redeploying.

### Adding a second batch of guests later

Add their rows to Guests, then run **Wedding > Refresh parties & links** from the
spreadsheet menu. New parties get a link and are open by default. Existing parties
and their links are never touched. Then run **Wedding > Email invites to new parties**,
which only mails parties that have not been emailed yet.

## Changing the site

Everything a person normally edits is in the `CONFIG` block at the top of the
`<script>` in `index.html`: names, date, venue, email, the Apps Script URL, and the
registry items with their Stripe links. The schedule is the `SCHEDULE` array just
below it and the FAQ is the `FAQ` array below that.

Commits to `main` go live in about a minute.
