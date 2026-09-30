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

## Step 1 – IT: allow the server to send e-mail (≈15 min, once)
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
