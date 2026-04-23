# ServiceNow ICS to Outlook Chrome Extension

This extension watches ServiceNow network activity while you schedule a meeting and opens a pre-filled Outlook Web event compose page.

## What It Automates

- Detects when a meeting-create `POST` request is sent from ServiceNow
- Captures the request timestamp and start/end times from that schedule call
- Waits for `list_history.do` JSON responses and extracts journal entries from Work Notes / Additional Comments
- Opens Outlook only when the journal entry is close to the captured schedule timestamp
- Opens Outlook deeplink:
  - `https://outlook.office.com/calendar/0/deeplink/compose`
  - with `subject`, `startdt`, `enddt`, `body`, `location=Online`, `online=1`

## Install (Developer Mode)

1. Open Chrome and go to `chrome://extensions`.
2. Enable **Developer mode**.
3. Click **Load unpacked**.
4. Select this folder.
## How It Avoids Old Invites

The extension uses two guards so page refresh/history reload does not create duplicate events:

1. **Timestamp check**
  - Each journal item from `list_history.do` has `sys_created_on_adjusted`.
  - The extension only accepts items created within a short window (12 seconds) of the captured schedule `POST` timestamp.

2. **Journal `sys_id` dedup**
  - Processed journal `sys_id` values are stored in `chrome.storage.local`.
  - If a `sys_id` is seen again (including after refresh), it is skipped.

## Test Flow

1. Open a ServiceNow case page.
2. Schedule a meeting from the UI.
3. Ensure the network shows the meeting-create `POST` and `list_history.do` response.
4. The extension should open exactly one Outlook compose tab for the matching fresh journal entry.

## Notes

- This approach is intentionally event-driven from network calls and no longer depends on downloading/parsing `.ics` files.
- If your instance uses non-standard endpoint names for meeting scheduling, adjust keyword matching in `content.js` (`isMeetingSchedulePost`).
