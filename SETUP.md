# BLIX Parts app – setup (accounts, progress per account, automatic e-mail)

## What's in this folder
- `public/` – the phone app (434 parts with preview pictures, works offline)
- `netlify/functions/api.mjs` – the server: logins, saving progress per account, photos, e-mailing reports
- `netlify.toml`, `package.json` – tell Netlify where everything is and what to install

## How it works
- **You create an account** (name, username, password) for each mechanic on the admin page of the app and give him his login.
- The mechanic **signs in** on his phone. Everything he saves (status, what to change, photos, responsible, date) is stored **on his account** – so he can stop, continue days later, or sign in on another phone and carry on.
- Offline in the workshop? Work is kept on the phone and synced as soon as there is a connection.
- **Send to Engineering** e-mails an Excel file with all his drafts and photos to you (via Microsoft 365). After that the parts show **"Sent · name · date"** for the whole team.
- Passwords are stored as secure hashes (never readable). You can give someone a new password or block an account at any time; he is then signed out on every phone.
- Data is stored in **Netlify Blobs** (Netlify's built-in storage, free, no setup). Drafts and their photos are removed from Netlify once they are e-mailed; only the report text stays so the team sees what is done.

---

## Step 1 (quick option, no IT) – e-mail via Resend (≈5 min)
Use this until IT has done the Microsoft 365 part (or instead of it).
1. Go to https://resend.com → **Sign up** with **t.wiendels@blixautomotive.com** (the address the reports must go to). Confirm the e-mail.
2. In Resend: **API Keys → Create API Key** → name `BLIX Parts App`, permission **Sending access** → **Add**. Copy the key (starts with `re_`) – it is shown only once.
3. In Netlify → **Site configuration → Environment variables**, add:
   - `RESEND_API_KEY` = the key from step 2
   - `MAIL_TO` = `t.wiendels@blixautomotive.com`
4. **Deploys → Trigger deploy → Deploy site**.
The e-mails come from `onboarding@resend.dev`. Without your own domain, Resend only delivers to the address you signed up with – that is fine, because the reports only go to you.
Later, to send from a Blix address or to more people: in Resend **Domains → Add domain** (IT adds 3 DNS records), then set `RESEND_FROM` = `BLIX Parts App <partsreview@blixautomotive.com>`.
When `RESEND_API_KEY` is set, Resend is used; remove it to switch to Microsoft 365.

## Step 1 (alternative) – IT: send via Microsoft 365 (≈15 min, once)
In https://entra.microsoft.com (Microsoft 365 administrator):
1. **App registrations → New registration** – name `BLIX Parts App Mailer`, *this organizational directory only*. Copy **Application (client) ID** and **Directory (tenant) ID**.
2. **API permissions → Add → Microsoft Graph → Application permissions → Mail.Send → Add**, then **Grant admin consent**.
3. **Certificates & secrets → New client secret** (24 months). Copy the **Value** immediately.
4. Choose the sender mailbox, preferably a free shared mailbox like `partsreview@blixautomotive.com`.
5. Recommended – limit the app to that mailbox (Exchange Online PowerShell):
   ```powershell
   Connect-ExchangeOnline
   New-ApplicationAccessPolicy -AppId <client ID> -PolicyScopeGroupId partsreview@blixautomotive.com -AccessRight RestrictAccess -Description "BLIX Parts app"
   ```
Until this is done, everything works except the actual e-mail (the app says "The mail service is not set up yet" and keeps the drafts).

## Step 2 – GitHub (free, 5 min)
1. https://github.com → sign up with your work e-mail → **+ → New repository** → name `blix-parts-app`, **Private** → Create.
2. Click **uploading an existing file**, drag in the **contents** of this folder (`public`, `netlify`, `netlify.toml`, `package.json`, `SETUP.md`) → **Commit changes**.

## Step 3 – Netlify (free, 5 min)
1. https://app.netlify.com → **Sign up with GitHub**.
2. **Add new site → Import an existing project → GitHub →** `blix-parts-app` → **Deploy** (leave build fields empty).
3. **Site configuration → Environment variables**, add:

| Key | Value |
|---|---|
| `ADMIN_PASSWORD` | a strong password you choose – for the built-in **admin** login |
| `SESSION_SECRET` | `XpK8hBLg5ARaU0VWWmvCMlzPqAjyDNcL_X_SRMfFm8hfRtc0DQ6Gnw` |
| `MS_TENANT_ID` | from IT (step 1) |
| `MS_CLIENT_ID` | from IT |
| `MS_CLIENT_SECRET` | from IT |
| `MAIL_FROM` | `partsreview@blixautomotive.com` |
| `MAIL_TO` | `t.wiendels@blixautomotive.com` (more: separate with commas) |

4. **Deploys → Trigger deploy → Deploy site.**
5. **Site configuration → Change site name** → e.g. `blix-parts` → app at `https://blix-parts.netlify.app`.

## Step 4 – Create the accounts
1. Open the app, sign in with username **admin** and your `ADMIN_PASSWORD`.
2. **First create an account for yourself** with role **Admin** (e.g. username `twiendels`). Sign out and use that account from now on – it can do everything the admin login can, and you can also report parts yourself.
3. Create an account for each mechanic (full name → username is filled in, a password is generated). Tap **Copy login details** and send them via WhatsApp or e-mail.
4. Fill in the **Names for "Responsible"** and tap **Save names** – everyone gets this list.

## Step 5 – Mechanics
They open the link, sign in, and install it: Android Chrome ⋮ → **Install app**; iPhone Safari → Share → **Add to Home Screen**.

### Good to know
- **Forgot password / new phone:** Accounts → **New password**. The old password stops working immediately.
- **Someone leaves:** Accounts → **Block**. He is signed out everywhere.
- **Keep `ADMIN_PASSWORD` and `SESSION_SECRET` private.** Changing `SESSION_SECRET` signs everyone out.
- **Updating the app:** upload the new files to GitHub (Add file → Upload files); Netlify publishes automatically.

## Revisions and locking
- When a report is sent, every part in it is **locked** (🔒 Sent) – it can't be changed or sent again.
- Directly below it a **new revision** opens with the same number plus the next letter: `CPF-I-LA200` → `CPF-I-LA200-A` → `-B` … (a part that already ends in `-B` continues with `-C`). New changes go there, so the list matches the engineering folder once the part is updated in CAD.
- The e-mail and Excel show the new revision number next to each part.
- If two mechanics work on the same part and one sends first, the other's draft is moved to the new revision automatically.

## Clearing changes
- Any account can open a part and tap **Clear changes…** (on a locked part or a new revision). A full-screen warning shows exactly what will be removed; choose *only the last sent revision* or *everything*, write a reason and tap **Clear and notify Engineering**.
- Admin accounts also see **Clear section** next to each category heading in a vehicle list, to clear all sent changes in that section at once.
- Clearing removes the sent reports (the part is unlocked and its revisions disappear) and deletes unsent drafts of all users on that part – for everyone.
- Engineering first receives an e-mail (who, when, reason, what was removed). If the e-mail can't be sent, nothing is cleared.
- The part then shows a line "⟲ Changes cleared on … by … – reason" so everyone sees it before making the new change.

## Settings (gear icon, Admin accounts only)
- **Maintenance – reset list:** choose 1 · vehicle (or all vehicles) → 2 · category (or all) → 3 · check the path and counts → **Reset…** → type **CONFIRM** → **Reset list**.
- Removes all sent reports, revisions, drafts of every user and "cleared" notes in that path. **No e-mail** – every reset is logged under **Maintenance history**, and users see a notice on the home screen for a few days.
- Also here: Accounts & names, Refresh data now, app version.

## Engineering account and "Released in CAD"
- Create an account with role **Engineering** (Accounts page; the role of any account can now be changed there with the drop-down).
- Engineering sees a **Waiting for Engineering** list on the home screen. On a sent part (or its new revision) tap **Mark …-A as released in CAD** (optional note). **Undo release** is possible.
- Everyone sees the result: "✓ …-A released" / "⏳ Engineering" tags, a green "Released in CAD" line on the revision, filters *Waiting for Eng.* and *Released*, and a progress line on the home screen.
- Engineering accounts can't create drafts or send reports; release buttons are only visible to them.

## Download overview (gear icon, every account)
- Choose vehicle and category (or all), optionally with pictures, and tap **Download Excel**. On a phone the share sheet opens (save to Files, e-mail, WhatsApp …).

## Weekly backup
- `netlify/functions/backup.mjs` runs every **Monday 05:00 UTC** (07:00 summer / 06:00 winter, Amsterdam) and e-mails an Excel overview plus a full .json data file to MAIL_TO.
- Admins can also tap **Send a backup now** under Settings → Backup. Netlify → **Logs → Functions → backup** shows each run.

## Draw on photos
- After taking a photo the drawing screen opens (tap **Done** to skip). Tap any photo later to draw on it again.
- Tools: **Draw** (finger), **Arrow**, **Circle**, **Text**, **Number** (①②③ – refer to them in "What to change"), 6 colours, **Undo**. The marked-up photo is what goes in the e-mail/Excel.

## Messages (envelope icon)
- Every account (not the built-in `admin` login) has an envelope icon with a red counter for unread messages and new releases.
- **New message:** choose one or more people, optionally link a part, write, add photos (with drawing) and send. From a part: **💬 Discuss this part** (mechanics → Engineering is preselected; Engineering → the mechanic who sent it).
- Conversations update every few seconds while open; only the people in a conversation can read it. "Updates from Engineering" shows releases of parts you sent.

## v12 – full engineering & workshop tool
- **Credit saving:** the app syncs every 3 minutes (and when opened), chats refresh every 20 s. Settings shows server requests per day (Admin).
- **Change process:** Engineering approves or rejects a sent change (cost, old stock, effective from, deadline); after release the mechanic records **Fitted on vehicle** (chassis no.). The **Change board** shows every change in its column: New → Approved → Released → Fitted (or Rejected).
- **Part tools** on every part: PDF **drawing** (Engineering uploads, everyone opens), **assembly instructions** with photos and torque values (anyone), **stock** with minimum and location, **supplier** with "send drawing to supplier" (Engineering), printable **QR label**, full **history**.
- **Parts** (Engineering/Admin, Settings or home tile): add a part with picture, edit, hide, or paste rows from Excel to import.
- **Problems:** problem reports with category, severity, chassis, part, photos, assignee, cause/action and timeline.
- **Checklists & inspections:** Engineering makes checklist templates; mechanics run them per chassis (OK / Not OK, values, photos) and sign off; Excel report per inspection.
- **Vehicles:** build record per chassis number with fitted changes, inspections and problem reports.
- **Dashboard**, **Stock & labels** (order list + label sheets), filters **My parts** / **Overdue**, weekly **overdue e-mail** (Monday, also "Send now" in Settings), **Nederlands/English** switch.
- Note: e-mails to suppliers only work once your own domain is verified in Resend (in test mode Resend only delivers to your own address).
