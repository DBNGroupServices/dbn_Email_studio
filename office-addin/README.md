# DBN Email Studio for Outlook Classic — personal preview

This is a Microsoft 365 Outlook web add-in, supplied as an XML add-in manifest. It is for this computer initially. It adds a ribbon button in message read/compose mode and opens a task pane. It is not yet an app published in Microsoft Marketplace.

Implemented: read the current email/selected compose text; reply, rewrite and compose; summaries and action extraction; proofreading and translation; follow-up drafting; structuring supplied meeting notes; manual files and supported current-email attachments; saved templates; copy, insert at cursor or open an editable reply. Nothing is sent automatically. The existing local ChatGPT connection is reused.

## Installation

1. Keep DBN Email Studio running. Run `Prepare-Outlook-Addin.ps1` to prepare its loopback HTTPS endpoint. This generates a certificate and encrypts its private key with your Windows account. It does **not** install certificate trust by default.
2. Review the displayed certificate fingerprint. After approving trust for this app's localhost certificate, run `Prepare-Outlook-Addin.ps1 -TrustLocalCertificate`. This adds only that certificate to **Current User → Trusted Root Certification Authorities**. It does not change machine-wide trust or disable certificate checks.
3. Install `manifest.xml` through Outlook's custom add-in manager: open [Microsoft's sideloading dialog](https://aka.ms/olksideload), choose **My add-ins → Add a custom add-in → Add from file**, select the XML and review the permissions. The requested permission is **ReadWriteItem** for the current mail item. Your Microsoft 365 organisation must allow custom add-ins. If installation is blocked, provide the manifest to IT; do not change organisation policy to work around it.
4. Open an email or compose window in Outlook Classic and choose **DBN Email Studio → Email Studio**. Sign-in remains in your local desktop engine. The panel needs the engine running after a Windows restart; no hidden startup task is installed.

The initial preview uses `https://localhost:8788` and the official Microsoft Office.js CDN. Only loopback traffic is accepted. Outlook and Microsoft 365 hosting requirements need HTTPS; do not bypass a browser or Office certificate warning.

### Troubleshooting the preview

Keep the updated desktop engine running in the same normal Windows account as Outlook Classic. If the panel says it could not be started, close only the DBN Email Studio task pane and reopen its ribbon button after the engine update. Version 1.5.1 accepts the host parameters Outlook adds to its page URL. No manifest reinstallation is needed for that fix. Do not run the engine as a different Windows user or disable certificate checks. If a browser still reports an untrusted certificate, stop and check the certificate rather than bypassing the warning.

Version 1.5.2 offers **Create draft in Outlook Classic** when there is no linked Outlook item, including browser use and typed notes. This opens a new email with the result's subject/body and no recipients; it does not claim to reply to an existing thread. When opened from Outlook's ribbon, **Open reply in Outlook** replies to the linked message and **Insert at cursor** inserts into a compose window. A body-reading error no longer removes the independently available draft target. If your Outlook selection changes, the app preserves the result and refuses to insert it into the wrong email. Copy an existing result before reopening an older panel to load an update.

## What is not included yet

Whole-mailbox AI search/indexing, automatic categorisation and filing, rules, scheduled send, snooze/reminders, unattended reply drafting, calendar booking, meeting attendance/recording/transcription, CRM/Slack/Teams integrations, shared inbox collaboration, tracking pixels and organisation administration remain planned. They require separate implementation, permissions and testing. See `FEATURE-ROADMAP.md` for the reviewed scope across all seven reference products. Reading supplied meeting notes is not a meeting recorder.

For a team or hosted deployment, replace localhost hosting with an approved HTTPS deployment, review Microsoft/Graph permissions and review ChatGPT plan eligibility for that deployment. A working personal sign-in does not establish organisation-wide eligibility.
