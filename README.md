# Claude New Tab (Edge / Chromium extension)

Replaces the browser's new tab page with: a Claude prompt box, quick prompts,
today's + upcoming calendar events, and quick links.

## How it's loaded
`~/.config/microsoft-edge-stable-flags.conf` passes this folder in `--load-extension=...`
(same mechanism Omarchy uses for its own extensions). Page files (html/css/js) are read from disk on every load.
The service worker (`sw.js`) is different: Edge keeps the script it registered on first load, even across version bumps,
for extensions loaded with `--load-extension`. After editing it, either click **Reload** on edge://extensions, or rename the
file (and update `background.service_worker` in manifest.json) so the registration URL changes, then restart Edge.
To use in Chromium/Chrome instead, add the same path to `~/.config/chromium-flags.conf`.

Extension ID is pinned by the `key` in manifest.json (private key: `.extension-key.pem`, not needed at runtime):
`gfdflaodlclbohbbambnoacdnieifmao`

## Calendar sources (Settings page: click the extension icon, or the "Settings" link at the bottom)
- **Microsoft 365 (live)** – needs a (free) Entra app registration in your tenant: name it anything, platform
  **Single-page application**, redirect URI `https://<extension-id>.chromiumapp.org/` (the settings page shows the exact
  value), delegated permissions `Calendars.Read` + `User.Read`, then grant admin consent. Paste the Application (client)
  ID and Directory (tenant) ID into Settings and click *Sign in with Microsoft*. The sign-in renews itself silently.
- **iCal feeds** – any `.ics` URL (Outlook "Publish a calendar", Google "Secret address in iCal format", iCloud).
  Recurring events (RRULE/EXDATE/RECURRENCE-ID) and Outlook's Windows time zone names are handled in `ics.js`.

Events refresh every 10 minutes in the background and are cached so the tab paints instantly.

## Claude
- Type in the box and press Enter → opens `https://claude.ai/new?q=<your prompt>` (Ctrl+Enter opens a new tab).
- "Plan my day" builds a prompt from today's events. Prompts ending in `:` or `?` act as prefixes.
- Prompts and links are editable in settings (`Label | text` per line).

## Background
Settings → Background: default (system light/dark), solid color, gradient preset, or an image (URL, or a picture
uploaded from the machine — resized and stored in local extension storage, so it does not sync between devices).
Images get an adjustable darkening overlay and optional frosted cards; text tone is picked automatically or set manually.
Two "picture of the day" modes fetch a fresh image daily and cache it for 6 hours: **NASA APOD** (api.nasa.gov; the shared
DEMO_KEY is rate-limited, a free key can be entered in settings; HD by default) and **Wikimedia Commons POTD** (Commons API,
2560px thumbnail). An "Alternate daily" mode switches between them (NASA on even days, Commons on odd). All show an ⓘ icon bottom-right: hover for the explanation/caption and credit, click to open the source page.

## Developer guide
See [HOW_IT_WORKS.md](HOW_IT_WORKS.md) for architecture, data model, the calendar/OAuth pipelines, and dev workflow.

## Files
manifest.json · sw.js (service worker: fetch/cache/OAuth) · ics.js (parser) ·
newtab.html/css/js (page) · options.html/js (settings)

## Publishing to the Edge Add-ons store
1. Bump `version` in manifest.json, commit.
2. `./build.sh` → `dist/claude-newtab-<version>.zip` (the `key` field is stripped automatically; stores reject it).
3. Partner Center → Microsoft Edge Add-ons → New extension → upload the zip. Set visibility to **Hidden** so only people
   with the link can install. Privacy policy URL: link to PRIVACY.md in this repo.
   Permission justifications: `identity` = Microsoft sign-in; `storage`/`alarms` = cache + 10-minute refresh;
   `graph.microsoft.com` + `login.microsoftonline.com` = calendar API; optional `https://*/*` = user-entered iCal feed URLs.
4. After it's published, the store assigns a **new extension ID**. Open the extension's settings page on an installed copy,
   copy the redirect URI it shows, and add it to the Entra app registration under Authentication → SPA redirect URIs
   (keep the existing `gfdfla…chromiumapp.org` one for unpacked installs).
5. Install from the store link on other machines; Edge sync will install it on every signed-in Edge profile.

## Installing unpacked on another machine
Clone this repo, then Edge → edge://extensions → Developer mode → Load unpacked → pick the folder (on Linux you can
instead add the folder to `--load-extension=` in the Edge flags file). The extension ID comes from the public `key`
in manifest.json, which is committed, so every unpacked install from this repo gets the same ID
(`gfdflaodlclbohbbambnoacdnieifmao`) and needs no Entra change. `.extension-key.pem` is only the matching private key,
kept out of git; it is not needed to load the extension.
