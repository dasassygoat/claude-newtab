# Claude New Tab (Edge / Chromium extension)

Replaces the browser's new tab page with: a Claude prompt box, quick prompts,
today's + upcoming calendar events, and quick links.

## How it's loaded
`~/.config/microsoft-edge-stable-flags.conf` passes this folder in `--load-extension=...`
(same mechanism Omarchy uses for its own extensions). Page files (html/css/js) are read from disk on every load.
After editing background.js or manifest.json, bump `version` in manifest.json and restart Edge: Chromium caches the
service worker script until the extension version changes.
To use in Chromium/Chrome instead, add the same path to `~/.config/chromium-flags.conf`.

Extension ID is pinned by the `key` in manifest.json (private key: `.extension-key.pem`, not needed at runtime):
`gfdflaodlclbohbbambnoacdnieifmao`

## Calendar sources (Settings page: click the extension icon, or the "Settings" link at the bottom)
- **Microsoft 365 (live)** – Entra app registration "Claude New Tab" in the Lan Services Fbg LLC tenant
  (client ID a90740f9-bfe0-4b8f-920b-9312a909e0a6, tenant fb20f7ca-1014-49ad-ac5e-3c15f5f60194, created 2026-09-11),
  SPA redirect URI `https://gfdflaodlclbohbbambnoacdnieifmao.chromiumapp.org/`, delegated `Calendars.Read` + `User.Read`,
  admin consent granted. These IDs are the defaults in background.js, so just click *Sign in with Microsoft*
  (on the new tab's agenda card or in settings).
- **iCal feeds** – any `.ics` URL (Outlook "Publish a calendar", Google "Secret address in iCal format", iCloud).
  Recurring events (RRULE/EXDATE/RECURRENCE-ID) and Outlook's Windows time zone names are handled in `ics.js`.

Events refresh every 10 minutes in the background and are cached so the tab paints instantly.

## Claude
- Type in the box and press Enter → opens `https://claude.ai/new?q=<your prompt>` (Ctrl+Enter opens a new tab).
- "Plan my day" builds a prompt from today's events. Prompts ending in `:` or `?` act as prefixes.
- Prompts and links are editable in settings (`Label | text` per line).

## Files
manifest.json · background.js (service worker: fetch/cache/OAuth) · ics.js (parser) ·
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
