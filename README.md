# ZIET Mumbai Workshop Resource Repository

Central resource catalogue for three connected tiers:
1. Workshop Timetable App
2. Individual Workshop Apps
3. Central Repository App

## Central Google Sheet
Spreadsheet ID: `1zo-CnHZR5Pp88nCtwfuw5VSRfNi-vZcCJ2vlB6BE6hY`

The Apps Script backend is in `Code.gs`. It creates/maintains a `Resources` tab with a common resource schema. Updates use a stable `resourceId`; deletion is a soft-delete tombstone so clients can synchronise removals.

## Deploy the API (one-time setup)
1. Open the central Google Sheet.
2. Select **Extensions → Apps Script**.
3. Replace the editor contents with the complete code from this repository's `Code.gs`.
4. Save the project.
5. Select **Deploy → New deployment → Web app**.
6. Set **Execute as: Me** and **Who has access: Anyone** (or your organisation's approved access setting).
7. Deploy, authorise the requested spreadsheet permissions, and copy the Web App URL ending in `/exec`.
8. Open the deployed repository app and paste that URL under **Central API Settings**, then choose **Save API URL** and **Refresh records**.

The deployment owner must have edit access to the central spreadsheet. Do not publish OAuth tokens, passwords, or private keys in this repository.

## API contract
- `GET ?action=health` — service health.
- `GET ?action=list` — returns `{success:true, resources:[...]}`, including deleted records.
- `POST` JSON `{action:"upsert", resource:{...}, sourceApp:"..." }` — insert/update by `resourceId`.
- `POST` JSON `{action:"delete", resourceId:"...", sourceApp:"..." }` — marks the record `status:"deleted"`.

Required resource fields: `resourceId`, `workshopId`, `title`, `url`. Recommended shared fields include workshop name, day, date, session, type, description, speaker, source app, timestamps and status.

All three app tiers must use this same deployed endpoint and the same field names. A link to another app is navigation only; API calls are what synchronise records.

## Important implementation note
The workshop selector is seeded in `index.html`. Verify and update the master list against the official ZIET Mumbai Weekly Calendar before treating it as a complete authoritative catalogue. GitHub Pages hosts the interface; Google Apps Script hosts the API.
