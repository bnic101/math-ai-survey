# Mathematicians and AI: quick survey

A single-page, five-minute survey hosted on GitHub Pages. Responses are posted to a
Google Apps Script web app that appends them as rows in a private Google Sheet.

- `index.html` — the survey (no build step, no dependencies)
- `apps-script/Code.gs` — the backend to paste into the Sheet's Apps Script editor

## Wiring up responses

1. Create a new Google Sheet.
2. Extensions → Apps Script. Replace the default code with `apps-script/Code.gs`. Save.
3. Deploy → New deployment → Web app. Execute as **Me**; who has access **Anyone**. Deploy and authorise.
4. Copy the Web app URL (ends in `/exec`) into `ENDPOINT` near the bottom of `index.html`. Commit and push.

Until `ENDPOINT` is set, the page shows a notice and logs submissions to the browser console.

## Privacy

Email is optional and used only for follow-up. The sheet is private to the owner of the Google account.
