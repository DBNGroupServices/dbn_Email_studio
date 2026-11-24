# Validation — 27 October 2026

Passed 17 backend tests on Windows using Python 3.12.14. Covered saved Outlook MSG and EML extraction, embedded attachments, DOCX paragraphs, HTML-to-text conversion, multimodal PDF/image request construction, invalid uploads, user edits and refinement context, structured AI responses, incomplete/refused response handling, safe API error messages, encrypted credential/history storage, history limits/settings, local request protection, static assets and concurrent-generation protection.

The built one-file Windows executable passed a separate smoke test: start without Python invocation, load embedded interface assets, read real compound-format Outlook and EML fixtures, store and remove an encrypted fictional API key, and shut down cleanly. Tests used an isolated workspace data directory.

Browser checks verified the first-run connection dialog, loading an example, rendering a simulated generated draft, preserving a manual edit through a simulated revision, copying subject/body, showing recent history, uploading EML/MSG files, inspecting Outlook text, and pasting a JPEG screenshot. Browser console inspection found no warning/error entries during those checks.

AI generation was simulated for validation. No real user API key was available, so live OpenAI drafting, model access and billing were not verified. The packaged app requires the user's own API key to perform live drafting.

## Version 1.1 interface redesign

The light, blue-accented redesign retains the original element bindings and writing workflow. All 17 backend tests and JavaScript syntax checks passed after the changes. The build also includes an updated icon, clearer task selection cards, larger editable draft text, consistent SVG controls and reduced-motion support. The existing production app was checked for an empty workspace before replacement.

Browser regression checks passed for example loading, simulated draft creation, preserving a manual edit during refinement, copying the edited email, recent-draft navigation and settings. The settings modal was visually checked with its save action visible and one internal scrolling region. No browser console warning/error entries were reported during these checks.

The version 1.1 executable passed the packaged smoke test, including its new version identifier and redesigned embedded interface assets, saved Outlook/EML readers and isolated encrypted storage.

## Version 1.2 Outlook and connection fixes

Reproduced the reported upload failure using the supplied saved Outlook email. An embedded attachment named image.png contains JPEG data. Content-based image detection now reads the whole email and all eight embedded images without warnings. The supplied message has not been copied into the examples or source package.

All 34 automated tests passed. New coverage includes misnamed images, safe handling of unreadable embedded images, Outlook import and opaque reply links, expired Outlook source protection, duplicate-action protection, escaped draft transport, temporary-export cleanup, PowerShell script parsing, and distinct DNS/certificate/timeout/Windows-network errors. Network diagnostics never include an API key or email content.

The Outlook Classic integration is tested with fixtures and mocks. No live mailbox import or reply was performed. The native bridge requires an already-running Outlook Classic session and permitted local Office automation. A ribbon or task-pane add-in is not installed.

Unauthenticated HTTPS to OpenAI was reachable from the computer during diagnosis. The existing running app reported no stored API key. The updated Save & test connection action stores a supplied key before testing, and a failed test does not remove it. Account/model authentication still requires the user's key to be saved in the app.

The executable-only connection failure was reproduced and traced to a missing codex-rejected-cas.pem support file loaded by the build runtime's SSLContext. The custom SSL packaging hook includes that adjacent policy file in the frozen app. The runtime's certificate checks and distrust policy remain intact; no TLS verification is disabled or bypassed. Future builds use the same hook through build.ps1.

After rebuilding with the hook, the executable passed a live unauthenticated OpenAI HTTPS probe (expected HTTP 401 without credentials), the supplied-message regression (body plus all eight images, no warnings), the saved Outlook/EML readers, and encrypted key storage/removal in the isolated test profile. This confirms transport from the actual executable; authenticated drafting awaits the user's saved key.

The user subsequently saved their key and reported a successful connection test. An independent authenticated check from the installed app confirmed access to gpt-6.1-sol. A short fictional generation test reached OpenAI but was rejected with HTTP 429, code credit_balance_exhausted, type insufficient_quota: no API credits remain. No real email or attachments were submitted for that test. Version 1.2.1 correctly distinguishes that billing condition from a temporary rate limit. The account needs API credits before live drafting can be verified. All 36 final automated tests passed, including the specific billing error and retention of an existing local session during an update.

## Version 1.2.2 — 28 October 2026

The saved key was checked with a short fictional request. OpenAI again returned credit_balance_exhausted for the organisation and project linked to that key. This does not independently establish the balance of another organisation visible in the user's billing dashboard. No actual email, attachment or saved signature was submitted.

The connection test now checks actual drafting using a fictional prompt and default preferences. Error handling distinguishes credit balance, organisation spend limit, project spend limit and approved usage limit. Only validated identifiers are returned in support details; API response messages and submitted content are excluded. All 40 automated tests and the JavaScript syntax check passed. Successful live drafting remains unverified pending resolution of the key's billing response.

## Version 1.3.0 — ChatGPT plan connection, 28 October 2026

Implemented the official local-client Sign in with ChatGPT contract, with dynamic registration, stable host identity, PKCE/state/nonce, verified RS256 ID tokens, account-specific model discovery, encrypted sessions, rotating refresh tokens, sign-out revocation, and Responses streaming. Requests omit unsupported subscription-route fields and use the public Responses endpoint. Plan errors never fall back to separate API billing.

All 54 automated tests and the JavaScript syntax check passed. Added coverage includes callback replay, mismatched state/client/identity/issuer/audience/nonce/expiry, declined plan permission, session rotation, encrypted credential redaction, protected connection endpoints, interrupted/failed streams, subscription usage limits and unsupported-field removal. UI drafting and edited-draft refinement were checked with a fictional simulated connection. A successful real ChatGPT sign-in and plan-backed draft require the user's browser sign-in and consent; simulation does not confirm account eligibility.

The blue envelope/sparkle icon is rendered at high resolution with all Windows icon sizes. It is embedded in the executable and supplied separately for a fresh shortcut icon reference.

The final 1.3.0 standalone executable passed startup, embedded interface, verified HTTPS, fictional MSG/EML parsing, encrypted key storage/removal, and clean shutdown checks. The Desktop installation and shortcut were updated, and the prior API credential file was verified unchanged. The installed executable successfully opened the official browser authorization flow. Real sign-in, account eligibility and plan-backed inference await the user's authentication and consent.

## Version 1.3.1 — live ChatGPT plan verification

The user completed sign-in and enabled plan permission. Live model discovery succeeded. A real completed response showed that this route can omit output from the terminal response envelope, carrying the draft in earlier item/text stream events. The reader now assembles those events and returns text only after response.completed; failures, refusals and interruptions still stop drafting. Connection polling now reads account state after the authorization transaction finishes.

All 58 automated tests passed. The corrected reader successfully returned a complete fictional email using the user's ChatGPT OAuth connection. A further live request with the fictional Outlook MSG fixture and blue icon image also returned a complete draft. No real user email or attachments were submitted. Separate API-key billing was not used for these tests.

The final 1.3.1 executable passed the isolated packaged smoke test. Both installed executable copies were replaced, and the encrypted ChatGPT sign-in was verified unchanged across the restart. The restarted app reported the ChatGPT plan connection enabled.

Final installed-app browser verification also passed: composing from fictional dot points produced an editable subject/body and saved the generated result to local history. The interface displayed Using ChatGPT plan. The demonstration and its source are clearly labelled fictional; no email was sent from Outlook.

## Version 1.4.0 — Outlook email picker

Import now opens a local email/folder picker. Metadata is read on demand through Outlook tables in pages of 50, with subject/sender search and mailbox/folder navigation. The import endpoint requires a valid, short-lived picker selection and retrieves that exact message by its Outlook IDs. It cannot silently import whichever item happens to be highlighted in Outlook. Raw Outlook IDs stay in server memory.

All 65 automated tests and JavaScript syntax checks passed. New checks cover required selection, opaque folder/message references, expired/forged/previous-page choices, safe query transport, read-only header listing, matching reply context and native PowerShell syntax. Live Outlook header listing, an apostrophe-containing search and an older-page request succeeded without opening full messages or contacting AI.

Browser tests used a fictional 75-message mailbox to verify paging, search, oversize-message handling, safe text rendering of an HTML-like subject, explicit selection, cancellation, selected-email import and preservation of dot points. The screenshot preview uses fictional test emails. No real user email was imported or sent to AI for this change.

The final standalone executable passed its isolated startup/assets, verified HTTPS, saved MSG/EML readers and encrypted-storage smoke test. Its live Outlook picker returned local email headers and mailbox choices with opaque references. A positive sender search also found the expected header. No full message was opened during these live mailbox checks.

Installed version 1.4.0 was checked in the browser: the picker listed real Outlook headers and mailbox choices, and Import remained disabled until a choice. Cancellation returned to the workspace. The saved encrypted ChatGPT connection was verified unchanged across the upgrade.

## Version 1.5.0 - personal Outlook add-in preview, 28 October 2026

All 68 Python regression tests passed with normal localhost and Windows temporary-file access. The Office JavaScript adapter checks passed for selected compose text, HTML-escaped insertion, reply forms and protection against applying a result to a different email. The final standalone executable passed startup/assets, verified OpenAI HTTPS connectivity, fictional MSG/EML reading and encrypted-storage smoke checks. Microsoft's manifest validator accepted the add-in manifest.

The installed desktop engine reports version 1.5.0 with the existing ChatGPT plan connection enabled. After the user's explicit approval, the localhost leaf certificate was installed into Current User Trusted Root Certification Authorities. Its SHA256 was checked before import. A fresh Windows process fetched the HTTPS task pane with normal certificate validation and received HTTP 200. The in-app browser still reported ERR_CERT_AUTHORITY_INVALID; no browser warning was bypassed and certificate checks were not disabled.

Outlook sideloading and live Office-host read/compose/insertion remain pending installation. The native running-window icon update is implemented but has not yet been visually verified. Whole-mailbox tools and meeting integrations remain roadmap items, not completed features.

## Version 1.5.1 - Outlook startup and reconnection repair, 28 October 2026

Microsoft's add-in manager subsequently showed DBN Email Studio as Added. The user then reported an Outlook startup error and a desktop import error. Diagnostics reproduced HTTP 404 for the task-pane and command URLs when query parameters were present. The static router now ignores query parameters only for known page/asset paths, while keeping callback handling and API authorization intact. The Office page policy explicitly permits outlook.cloud.microsoft.

All 68 Python regression tests passed, including real TLS requests with Office host parameters and an unauthorized API request with a query string. JavaScript checks verified automatic reconnection after the exact pre-execution stale-session denial, preserved POST content, bounded retry and no retry of origin denials, generation failures or rate limits. The rebuilt 1.5.1 executable passed its isolated smoke test and listed six Outlook headers from three mailbox choices without importing any message body or contacting AI.

The standalone import connection worked after moving the service back to the normal Windows session. Live task-pane loading and insertion in Outlook still need verification after the installed update; a successful HTTPS fetch alone does not establish Office-host operation.

A further isolated regression reproduced a TLS handshake timeout when one TCP client preconnected without sending a handshake. TLS negotiation now runs in the per-request worker with a timeout, allowing Outlook's next connection to proceed. The same test and all 68 backend tests passed after the repair. HTTPS status now correctly reports the running TLS listener when queried from the add-in itself.

The current OneDrive installation was replaced and verified as 1.5.1. Its secure endpoint reports running, accepts host-parameter URLs with HTTP 200, and lists Outlook headers. The encrypted ChatGPT sign-in file was unchanged and the plan connection remained enabled. Browser verification confirmed the desktop picker lists real Outlook headers and the HTTPS task-pane preview loads with Using ChatGPT plan and no certificate warning. The original input window was retained. Final native Outlook testing was unavailable because computer-control approval timed out; the user was asked to reopen the installed task pane.

## Version 1.5.2 - enabled Outlook draft action, 28 October 2026

The user confirmed the panel loads in Outlook and shows Using ChatGPT plan, then reported a generated result with its Outlook action disabled. The previous panel required a successful email-body read to enable any Outlook output. It also captured the next result's target before generation succeeded, which could incorrectly rebind an earlier result after a failed request.

The panel now captures the Outlook target independently of reading its body and offers Create draft in Outlook Classic when no linked item exists. That action uses the existing local bridge to open a new editable email with no recipients. Native replies/insertion remain tied to the result's original item. Failed generation, reading a different email and the fictional example cannot redirect an existing result into another message. The result's current manual edits are used, blank results stay disabled, and duplicate opening clicks are ignored while an action is pending. A failed selected-text read falls back to the full draft with an explicit review message.

All 68 Python tests passed. JavaScript controller checks covered notes-only/browser drafting, unreadable compose context, manually edited output, blank results, duplicate clicks, failed regeneration and changed-item protection. Office adapter and local-session reconnection checks also passed. Tests used fictional content; live verification is recorded separately when completed.

The 1.5.2 executable passed the isolated packaged smoke test and was installed in the user's current OneDrive app folder. The saved encrypted ChatGPT sign-in was unchanged. A fresh HTTPS browser panel generated a fictional reply using the existing plan and showed an enabled Create draft in Outlook Classic button. The text supplied in the user's screenshot was separately opened through the local draft endpoint without further AI processing; Outlook's object model confirmed the draft was open, contained the supplied text, had zero recipients and was not sent. The previous input/result windows were left intact. Insertion through the native Office host is covered by adapter/controller tests; it was not exercised against a real message during this repair.
