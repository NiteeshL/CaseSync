# ServiceNow ICS to Outlook Chrome Extension

This extension watches ServiceNow case pages for newly added `.ics` links (for example in Work Notes), fetches and parses the calendar invite, and opens a pre-filled Outlook Web event compose page.

## What It Automates

- Detects new `.ics` links dynamically using a `MutationObserver`
- Fetches the `.ics` file with current ServiceNow session cookies
- Parses:
  - `SUMMARY` (subject)
  - `DTSTART`
  - `DTEND`
  - `DESCRIPTION`
- Extracts:
  - Case number pattern such as `INC12345`
  - Meeting URL from description (Zoom/Teams/etc.)
- Converts ICS datetime values like `20260416T100000Z` to `YYYY-MM-DDTHH:mm:ss`
- Opens Outlook deeplink:
  - `https://outlook.office.com/calendar/0/deeplink/compose`
  - with `subject`, `startdt`, `enddt`, `body`, `location=Online`, `online=1`

## Install (Developer Mode)

1. Open Chrome and go to `chrome://extensions`.
2. Enable **Developer mode**.
3. Click **Load unpacked**.
4. Select this folder.

## Notes

- The extension currently runs on `*.service-now.com` pages. If your ServiceNow portal is on a custom domain, update `matches` and `host_permissions` in `manifest.json`.
- Existing `.ics` links already present during initial page load are ignored intentionally to avoid opening old invites.
- Newly added `.ics` links are processed once per URL.
