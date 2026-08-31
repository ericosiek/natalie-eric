# natalie-eric.website

Natalie & Eric's wedding website. July 11, 2027 · Riverway Clubhouse, Burnaby BC.

Static pages on GitHub Pages, with a Google Apps Script web app and a Google Sheet
behind the RSVP flow and the admin portal.

- **Guests** see `/` — invitation, schedule, RSVP, FAQ — and `/registry/`.
- **Natalie & Eric** use `/admin/` to manage guests, parties, invite waves, emails and
  site content.

**[SETUP.md](SETUP.md) explains how to run it.** Start there.

## Layout

```
index.html            invitation overlay, schedule, RSVP, FAQ
registry/index.html   Honeymoon Fund
admin/index.html      admin portal (sign-in required)
styles.css            shared styling
rsvp-backend.gs       Apps Script source (tracked here; deployed by hand)
CNAME                 natalie-eric.website
```

## Changing things

Ask Claude to edit and push. Content that lives in the sheet — schedule, meals, FAQ —
is quicker to change in the admin portal's Site content tab.
