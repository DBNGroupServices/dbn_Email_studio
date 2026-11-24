# DBN Email Studio for Outlook Classic — install & run (v2.0)

The Outlook add-in is a pane that opens from the **Email Studio** button on the Home ribbon. Its pages are served by the DBN Email Studio engine running on your own computer at `https://localhost:8788`. **If the engine isn't running, Outlook can only say "Sorry, we can't load the add-in"** — that was the cause of the 1.x failures. Version 2 fixes this by starting the engine when you sign in to Windows and renewing its certificate automatically.

## Install or upgrade (about 3 minutes, no admin rights)

1. Unzip the release (or download the `DBN-Email-Studio-…` artifact from the GitHub Actions run) so you have:
   ```
   DBN Email Studio\
     DBN Email Studio.exe
     Outlook add-in\Install Email Studio.cmd   <- double-click this
   ```
2. Double-click **Install Email Studio.cmd**. It:
   - copies the app to `%LOCALAPPDATA%\Programs\DBN Email Studio` and adds a Start menu shortcut
   - turns on **start at Windows sign-in** (background, no window)
   - starts the engine and creates a local certificate authority (localhost-only, name-constrained)
   - asks Windows to trust that certificate for **your user only** — a Windows security prompt appears; check the thumbprint matches what the script printed and choose **Yes**
   - confirms the pane page loads over trusted HTTPS
   - opens Outlook's custom add-in page
3. **First install, or upgrading from 1.x:** in Outlook choose **Get Add-ins → My add-ins → Custom add-ins → Add a custom add-in → Add from file**, and select `manifest.xml` in `%LOCALAPPDATA%\Programs\DBN Email Studio`. (Upgrading replaces the 1.x add-in; v2 adds pinning.) Restart Outlook.
4. Open any email → **Email Studio**. The first time, open **Settings (gear) → AI connection → Continue with ChatGPT** if the pill says *Not connected*.

If your organisation blocks custom add-ins, give `manifest.xml` to IT for central deployment; don't work around the policy.

## If the pane says "can't load the add-in"

Double-click **Fix Email Studio.cmd** (in the install folder or the release's *Outlook add-in* folder). It checks — and with repair, fixes — each cause in turn:

| Check | Typical fix |
|---|---|
| Engine running on 8767 | Starts it |
| Add-in HTTPS port 8788 listening | Re-run the installer |
| Local certificate authority trusted / in date | Re-run the installer, choose Yes |
| Pane page loads with Windows trust | — |
| Start at sign-in | Turns it on |
| Outlook add-in cache | Clears `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef` (Outlook must be closed) |

Then click **Retry** in Outlook (or close and reopen the pane).

## Scripts

| File | Purpose |
|---|---|
| `Install Email Studio.cmd` / `Install-EmailStudio.ps1` | Install, upgrade, trust, verify. Options: `-NoAutoStart`, `-NoTrust`, `-ClearOfficeCache`, `-ExePath`. |
| `Fix Email Studio.cmd` / `Diagnose-EmailStudio.ps1` | Diagnose (`.ps1` alone) or diagnose and repair (`-Repair`). |
| `Prepare-Outlook-Addin.ps1` | 1.x name, now calls the installer. |
| `manifest.xml` | Outlook add-in manifest (same add-in ID as 1.x, version 2.0.0.0). |

## Permissions and data

- **ReadWriteItem** — the add-in reads only the email you have open and can open a reply / insert into your draft. It cannot read the rest of your mailbox and never sends mail.
- Email content goes to OpenAI only when you press Generate. Everything saved locally (sign-in, settings, voice profiles, templates, history, certificate keys) is encrypted with Windows DPAPI for your account.
- The certificate authority can only vouch for `localhost` / `127.0.0.1`, so trusting it cannot be used to impersonate other sites. To remove it: `certmgr.msc` → Trusted Root Certification Authorities → *DBN Email Studio Local CA*.

## Scaling to a team later

The current design is one engine per PC (each person signs in with their own ChatGPT plan). For a team rollout without per-PC installs, host the pane on an approved HTTPS site and move AI calls to a small server with an organisation API key, then deploy the manifest centrally from the Microsoft 365 admin centre. See `FEATURE-ROADMAP.md`.
