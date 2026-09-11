# Privacy Policy – Claude New Tab

Claude New Tab is a browser extension that replaces the new tab page with a Claude prompt box, a calendar agenda, and quick links.

## Data the extension handles
- **Calendar events.** If you sign in with Microsoft 365, the extension requests the delegated `Calendars.Read` permission and reads upcoming events from Microsoft Graph. If you add iCal (.ics) feed URLs, it fetches those feeds. Events are cached only in the browser's local extension storage on your device and are shown on the new tab page.
- **Sign-in tokens.** Microsoft OAuth tokens are stored in local extension storage on your device and are used only to call Microsoft Graph on your behalf.
- **Settings.** Your name, quick prompts, quick links, feed URLs, and app IDs are stored in browser extension storage (which your browser may sync between your own devices).
- **Prompts.** Text you type in the prompt box is sent to claude.ai only when you press Enter, by opening claude.ai in your browser.

## What the extension does not do
- No data is sent to the extension author or to any third-party server. The only network calls are to Microsoft (login and Graph), the calendar feed URLs you configure, claude.ai when you open it, and favicon images for your quick links.
- No analytics, tracking, or advertising.

## Removing data
Sign out in the extension settings to delete stored tokens, or uninstall the extension to remove all stored data.

Contact: open an issue on the GitHub repository.
