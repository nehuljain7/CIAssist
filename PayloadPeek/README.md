# PayloadPeek (Chrome MV3)

A **dockable in-page panel** to search attachment payloads of SAP Cloud Integration message
processing logs for a keyword across **all filtered messages on all pages** — **no stored
credentials**, using the session you are already logged in with.

## Why a panel (not a popup)
The UI is injected into the CPI page by the content script, so it:
- **docks left / right / bottom** (toolbar-style buttons in the panel header),
- **stays on screen** when you switch tabs or open a new tab (e.g. clicking a result),
- **keeps its data** for the life of the tab (and restores the last run after a reload).

A toolbar popup couldn't do any of this — it closes the instant focus leaves it.

## Auth (no credentials)
All requests are made **same-origin** to the tenant in the active tab, so the browser sends your
existing session cookie (`JSESSIONID` etc.) automatically. It acts as whoever is logged in. Nothing stored.

## Endpoints (your extracted internal APIs)
- List: `GET <origin>/odata/api/v1/MessageProcessingLogs?$filter=...` (Atom-paged via `next` link)
- Attachments: `GET <origin>/odata/api/v1/MessageProcessingLogs('<MPLID>')/Attachments`
  - name from `entry/m:properties/d:Name`, URL from `entry/id`
- Content: `GET <entry/id>/$value` → keyword search

## Install
1. `chrome://extensions` → **Developer mode** → **Load unpacked** → select this folder.
2. Open your CPI tab (`*.hana.ondemand.com`). If it was open before install, reload it once.

## Use
1. Click the **toolbar icon** to toggle the panel (it opens docked; switch dock with the
   ◧ / ▭ / ◨ header buttons; drag the panel edge to resize).
2. Enter **Keyword** and **Payload (attachment) name** (blank = all).
3. Pick the timeline — **To defaults to the current date/time** — plus optional Status / IFlow.
4. **Search**. Results stream into the table; the **attachment** cell opens the raw `$value`
   payload in a new tab (panel stays put). **Export CSV** / **Copy message IDs** as needed.

## Notes
- Times are local, converted to UTC for the OData filter.
- Payload-name match = case-insensitive *contains*; keyword has case-sensitive and regex toggles.
- Attachments must have been logged by the iFlow and still be within the retention window.
- Custom domain (not `*.hana.ondemand.com`)? add it to `matches`/`host_permissions` in `manifest.json`.
