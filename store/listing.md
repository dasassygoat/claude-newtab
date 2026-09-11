# Edge Add-ons store listing text

**Name:** Claude New Tab
**Short description (≤132):** New tab page with a Claude prompt box, your Microsoft 365 agenda for the next few days, and quick links.
**Category:** Productivity
**Visibility:** Hidden
**Privacy policy URL:** (public gist or repo link to PRIVACY.md)

**Description:**
Replaces the new tab page with a focused start page:
- Ask Claude anything from the new tab; Enter opens claude.ai with your prompt prefilled.
- Quick prompt chips, including "Plan my day", which sends today's calendar to Claude.
- Agenda for today and the next few days from Microsoft 365 (sign in with Microsoft) or any iCal (.ics) feed. Recurring events, all-day events, and Teams join links are supported.
- Customizable quick links.
No data leaves your browser except calls to Microsoft (sign-in and calendar), the feed URLs you add, and claude.ai when you open it.

**Permission justifications:**
- identity: Microsoft sign-in (OAuth) for reading the user's calendar.
- storage, alarms: cache events locally and refresh every 10 minutes.
- graph.microsoft.com, login.microsoftonline.com: Microsoft Graph calendar API and sign-in.
- optional https://*/*: only requested when the user adds an iCal feed URL of their own, to fetch that feed.
