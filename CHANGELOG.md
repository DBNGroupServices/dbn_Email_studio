# Changelog

## 2.0.0 — 24 November 2026

### Fixed: "Sorry, we can't load the add-in"
Root cause: the Outlook pane is served by the Email Studio engine on this PC (`https://localhost:8788`). After a Windows restart nothing started the engine, so Outlook could not reach the pane. A 1.x certificate expiring would cause the same error.
- **Start at Windows sign-in** (per-user Run key, background, no window). Set by the installer; toggle in pane Settings → Reliability.
- **Self-renewing certificates.** A 10-year local certificate authority, name-constrained to `localhost`/`127.0.0.1`, signs a server certificate that renews itself 45 days before expiry. It is checked daily, so no re-trust is needed.
- **`Install Email Studio.cmd`**: one-click install or upgrade into `%LOCALAPPDATA%\Programs`, trust, verification and Start menu shortcut.
- **`Fix Email Studio.cmd`**: diagnoses the engine, ports, certificate trust and expiry, start-up and the Outlook add-in cache, and repairs what it can.
- **In-pane reconnect.** If the engine stops while the pane is open, a banner appears, your work is kept, and the pane reconnects automatically.

### New in the Outlook pane
- **Reads the open email automatically**, including sender, recipients and date, and shows a sender card.
- **Pinnable pane** (Mailbox 1.5+) that follows you as you change emails. If you move away from an unsent draft, the pane keeps it and offers to load the new email.
- **10 one-click quick replies**: Acknowledge, Confirm, Ask for info, Decline, Schedule, Chase up, Escalate, Passing on, Incident reply and Thanks.
- **Voice profiles**: formality, warmth and directness sliders, greeting, sign-off, things to avoid, signature handling, and up to three example emails for style matching. Five starter profiles are included, plus a default-voice setting and "use the open email as an example".
- **More tones**: Formal, Friendly, Diplomatic, Apologetic, Assertive, Reassuring and "Use profile".
- **Compare 2–3 different drafts** side by side.
- **Refine chips**: Shorter, Longer, Warmer, Firmer, More formal, Simpler, Bullets and Polish, plus free-text changes. **Version history** (‹ ›) lets you step back.
- **Reply all** (shown when there are multiple recipients). Replies are written with everyone copied in mind.
- **Proper Outlook formatting**: paragraphs and real bulleted and numbered lists instead of `<br>` text.
- **Before-you-send checks**: unfilled `[confirm …]` placeholders, a missing attachment that is mentioned, commitments (refund, guarantee, liability) and deadlines. Stop-level items need a second click.
- **Security disclosure guard** (on by default). It blocks access, alarm and door codes, PINs and passwords, and flags card numbers and coverage or roster detail, both in the AI instructions and in the pane's checks.
- Template delete; settings drawer; Office dark theme; Ctrl+Enter to generate; remembered preferences.

### Engineering
- New modules: `profiles.py`, `autostart.py` and `web/office/draft-tools.js`.
- New API: `/api/profiles` (+`/delete`), `/api/templates/delete`, `/api/autostart`, `/api/meta`. `/health` reports Office status.
- Tests: `tests/test_v2_features.py`, `tests/test_draft_tools.js`, and a real-browser end-to-end test `tests/test_taskpane_browser.js` (Playwright, simulated AI and Outlook).
- GitHub Actions builds and tests the Windows exe on every push and publishes the release zip as an artifact.
- Test-only scrambled store (`DBN_EMAIL_TEST_PLAINTEXT_STORE=1`, non-Windows only) so the suite runs on Linux CI.

## 1.5.2 and earlier
See `TEST-RESULTS.md`.
