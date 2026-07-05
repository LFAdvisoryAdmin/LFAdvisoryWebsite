# LF Advisory — daily task email (Cloudflare Worker)

Sends each **preparer** their tasks every **weekday (Mon–Fri) at 3am Brisbane** (cron `0 17 * * 1,2,3,4,7` = 17:00 UTC Sun–Thu; Cloudflare uses 7 for Sunday), with **Liam CC'd**. Reads the register straight from SharePoint and sends via Outlook — all inside your Microsoft 365 tenant, no external email service. If any send fails or the run crashes, Liam gets an alert email (and bounces/NDRs from undeliverable addresses arrive in Liam's inbox automatically, since it sends from his account).

Email layout: **Overdue** + **Due today** at the top (the priority), then **Due this week** below. Skips completed and not-ready jobs; skips people with nothing due.

## 1. Azure — give the app permission to read the file and send mail

On the existing app registration **`981d2ee1-a2a9-4787-8972-b349938ba7ab`** (Azure Portal → App registrations):

1. **API permissions → Add a permission → Microsoft Graph → Application permissions** → add:
   - `Sites.Read.All` (read the register file)
   - `Mail.Send` (send the emails)
2. Click **Grant admin consent** for the tenant.
3. **Certificates & secrets → New client secret** → copy the **Value** (you'll paste it as `CLIENT_SECRET` below). Save it somewhere safe — it's shown once.

> Security note: `Mail.Send` (application) lets the app send as any mailbox. To lock it to just the sender, create an **Application Access Policy** in Exchange Online restricting this app to `liam@lfadvisory.com.au` (or a shared mailbox). Optional but recommended.

## 2. Verify the people → email map

In `wrangler.toml`, edit `PEOPLE` so every name you use in the app maps to the right address, e.g.
`{"Hazel":"hazel@lfadvisory.com.au","Jennifer":"jennifer@lfadvisory.com.au","Liam":"liam@lfadvisory.com.au"}`
Anyone not listed (or unassigned jobs) falls back to `FALLBACK_EMAIL` (Liam).

## 3. Deploy the Worker

From this folder:

```
npm install -g wrangler
wrangler login
wrangler secret put CLIENT_SECRET   # paste the Azure secret value
wrangler secret put TRIGGER_KEY      # any random string (for the test URL)
wrangler deploy
```

(Or in the Cloudflare dashboard: Workers → Create → paste `src/worker.js`, set the Variables from `[vars]`, add the two Secrets, and add a Cron Trigger `0 17 * * *`.)

## 4. Test without waiting for 3am

Open: `https://lf-register-mailer.<your-subdomain>.workers.dev/run?key=YOUR_TRIGGER_KEY`

It runs immediately and returns e.g. `Sent 3 email(s) for 2026-07-06.` Check the inboxes.

## Notes
- Brisbane has no daylight saving, so `0 17 * * *` is always 3am there.
- "This week" = due within the next 7 days (matches the app's *This week* count).
- To change wording/sections, edit `buildHtml()` in `src/worker.js`.
