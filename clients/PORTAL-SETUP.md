# LF Advisory Client Portal — setup runbook

A client-facing portal at **`/clients`** where each client signs in (email PIN via
Cloudflare Access) and sees **only the tools they're entitled to**. Data is served
by the **`lf-client-api`** Worker, which reads/writes only that client's own file in
LF Advisory SharePoint — the client's browser never gets a Microsoft token and can
never reach another client's data.

Phase 1 client: **Flux Property Group** (Samantha O'Connor), tool: **Investor Loan Tracker**.

```
Samantha ──PIN email──▶ Cloudflare Access ──▶ /clients (her cards)
                              │ signed identity (JWT)
                              ▼
                        lf-client-api Worker ──app-only Graph──▶ SharePoint
                        · validates the JWT                        Client Portal Data/
                        · REGISTRY: email → client + apps            flux-loan-tracker.json
                        · serves ONLY flux-loan-tracker.json
```

Do the steps **in this order** (each depends on the one before).

---

## 1. Azure — let the Worker write client data (5 min)

The Worker reuses the existing Azure app **"LF Advisory Payroll Tool"**
(client id `981d2ee1-a2a9-4787-8972-b349938ba7ab`). It has `Sites.Read.All` today;
it needs **write**.

1. [Entra admin center](https://entra.microsoft.com) → **App registrations** → *LF Advisory Payroll Tool*.
2. **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions** → tick **`Sites.ReadWrite.All`** → Add.
3. Click **Grant admin consent for LF Advisory** (needs a Global/Privileged Role admin).
4. You'll need the app's **client secret** value for step 4. Reuse the one the mailer uses, or **Certificates & secrets → New client secret** and copy the value now (it's only shown once).

> Hardening for later: `Sites.Selected` instead of `Sites.ReadWrite.All` locks the app to just the one site. Not needed for launch.

## 2. SharePoint — create the data folder (2 min)

1. Open the LF Advisory root site document library (same library that holds **"LF Advisory Workpapers"**).
2. Create a new folder named exactly **`Client Portal Data`**.
3. (Seed) Once we've recovered Samantha's data, drop a file named exactly **`flux-loan-tracker.json`** in here. If you skip this, the tool loads the June-2026 baseline on first run — but seeding her recovered data means she picks up exactly where she left off.

## 3. Cloudflare Zero Trust — the login (10 min)

1. Cloudflare dashboard → **Zero Trust**. If it's your first time, create the org and **pick a team name** — this becomes `‹team›.cloudflareaccess.com`. **Write the team name down** (needed in step 4). Free plan is fine.
2. **Settings → Authentication → Login methods**: ensure **One-time PIN** is present (it's built in — emails the client a code, no password).
3. **Access → Applications → Add an application → Self-hosted**:
   - **Name:** `LF Advisory Client Portal`
   - **Session duration:** 24 hours (your call).
   - **Public hostnames / paths** — add **both** paths under the **one** application so a single sign-in covers both:
     - `www.lfadvisory.com.au` — path `clients`
     - `www.lfadvisory.com.au` — path `client-api`
     - (add the apex `lfadvisory.com.au` for both too, if the domain serves apex directly)
   - **Identity providers:** One-time PIN.
4. **Add policy:** Name `Allowed clients`, Action **Allow**, Include → **Emails** →
   `samantha@fluxproperty.com.au`, `liam@fluxproperty.com.au`, and **your own email** (for testing).
5. Open the finished application → **Overview** → copy the **Application Audience (AUD) tag** (a long hex string). **Write it down** (needed in step 4).

> Later, to onboard a whole client at once, add an **Email domain** to the policy instead of individual emails.

## 4. Deploy the `lf-client-api` Worker (10 min, dashboard — no wrangler needed)

Single-file Worker, so the dashboard is easiest (avoids the ARM64 wrangler problem).

1. Cloudflare dashboard → **Workers & Pages → Create → Create Worker**. Name it **`lf-client-api`**. Deploy the starter, then **Edit code**.
2. Paste the entire contents of **`client-api/src/worker.js`** (from this repo) over the starter, and **Deploy**.
3. **Settings → Variables and Secrets** — add these (plain text unless noted):

   | Name | Value |
   |---|---|
   | `SP_HOSTNAME` | `lfadvisoryptyltd.sharepoint.com` |
   | `FOLDER` | `Client Portal Data` |
   | `TENANT_ID` | `f869ffc9-81fa-4bbf-94ec-a9bd5ca6b3a3` |
   | `CLIENT_ID` | `981d2ee1-a2a9-4787-8972-b349938ba7ab` |
   | `ACCESS_TEAM` | your team name from step 3.1 |
   | `ACCESS_AUD` | the AUD tag from step 3.5 |
   | `REGISTRY` | *(the JSON below)* |
   | `CLIENT_SECRET` | **Encrypt** — the Azure secret from step 1.4 |

   `REGISTRY` value (one line):
   ```json
   {"samantha@fluxproperty.com.au":{"client":"Flux Property Group","clientId":"flux","apps":["loan-tracker"]},"liam@fluxproperty.com.au":{"client":"Flux Property Group","clientId":"flux","apps":["loan-tracker"]}}
   ```
4. **Settings → Domains & Routes → Add → Route**:
   - `www.lfadvisory.com.au/client-api/*`  (zone `lfadvisory.com.au`)
   - `lfadvisory.com.au/client-api/*`
   (Worker routes take precedence over the Pages site for these paths.)

## 5. Publish the pages

`git push origin main` — Cloudflare Pages publishes `clients/index.html` (→ `/clients`)
and `clients/loan-tracker.html` (→ `/clients/loan-tracker`). No data lives in the pages;
it only comes through the Access-protected Worker.

## 6. Test, then hand to Samantha

1. Open **https://www.lfadvisory.com.au/clients** in a private window → enter your email → get the PIN → you should see the **Investor Loan Tracker** card → open it → data loads and edits save (watch the status go "saving… / saved").
2. Recover Samantha's data (see the separate recovery note), save as `flux-loan-tracker.json`, upload to `Client Portal Data` (step 2.3), reload the tracker to confirm it shows her latest.
3. Send Samantha the link. She signs in with `samantha@fluxproperty.com.au`.

---

## Onboarding another client/tool later
- **New client on an existing tool:** add one line to `REGISTRY` (their email → client, clientId, apps) and add their email to the Access policy. No code change.
- **New tool:** add it to the `APPS` map in `worker.js`, build the tool page under `clients/`, point its storage at `/client-api/data/‹app›`, add the app id to the relevant clients' `apps`.

## What I still need from you to finish my side
- `ACCESS_TEAM` (team name) and `ACCESS_AUD` (AUD tag) — I'll bake them into `client-api/wrangler.toml` for the record. You still enter them as Worker variables in step 4 either way.
