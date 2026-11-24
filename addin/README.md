# DBN Email Studio 3.0 — hosted Outlook add-in (no desktop app)

This folder is the whole add-in. It is a static website published to GitHub Pages. Outlook Classic loads it straight from there, and the pane calls Claude directly with your own Anthropic API key. Nothing has to be installed or kept running on your PC.

```
Outlook pane (https://dbngroupservices.github.io/dbn_Email_studio/taskpane.html)
   ├─ reads the open email (Office.js, ReadWriteItem: current item only)
   ├─ settings, key, voice profiles, templates → Outlook roaming settings (your mailbox, 32 KB)
   └─ HTTPS → api.anthropic.com (official Anthropic SDK, bundled in vendor/)
```

## One-time setup

1. **Publish the site.** The repo needs GitHub Pages enabled. Pages for a private repository needs a paid GitHub plan; otherwise make the repository public. Only the add-in's code is public; your key and emails never touch GitHub.
   - GitHub → repository → **Settings → Pages → Build and deployment → Source: GitHub Actions**.
   - Merge this branch into `main`. The **Publish Outlook add-in** workflow runs the tests and deploys `addin/`. You can also run it manually from the **Actions** tab.
   - Check that https://dbngroupservices.github.io/dbn_Email_studio/ opens.
2. **Get a Claude API key.** At console.anthropic.com, add credit under **Billing**, then create a key under **API keys**.
3. **Install in Outlook.** Download `manifest.xml` from the site, then in Outlook choose **Get Add-ins → My add-ins → Custom add-ins → Add from file**. It uses the same add-in ID as earlier versions, so it replaces the old localhost one. Restart Outlook.
4. In the pane, open **Settings (gear) → AI connection**, paste the key and press **Save & test**.

## Changing the hosting address

Run `node addin-build/make-manifest.js https://your.host/path/`, then commit the new `addin/manifest.xml`, publish the folder at that address, and reinstall the manifest in Outlook.

## Updating the bundled Anthropic SDK

```
cd addin-build && npm install && npm run build
```
The SDK version is pinned in `addin-build/package.json`. Run `node tests/test_addin.js` and `node tests/test_addin_browser.js` afterwards.

## Models and cost

| Model | Use | Approx. cost per reply* |
|---|---|---|
| Claude Opus 5.5 (default) | Best writing | 2–5 cents |
| Claude Sonnet 5.5 | Faster, about half the cost | 1–3 cents |
| Claude Haiku 5.5 | Quickest, lowest cost | under 1 cent |

*Estimated for a typical email without attachments. Screenshots and PDFs add to the cost. Compare-3-options costs roughly 2–3 times as much.

Opus and Sonnet requests turn on Anthropic's server-side fallback (`fallbacks: "default"`). If a safety classifier declines a request, the API retries it on a fallback model within the same call.

## Security notes

- The key is stored in your mailbox's add-in roaming settings. Only this add-in can read them. The key is sent only to `api.anthropic.com`, which a Content-Security-Policy on the page enforces.
- This design suits personal use. For a team, move the key to a small server-side proxy so that it never reaches Outlook.
- The security disclosure guard is on by default, both in Claude's instructions and in the pane's checks before you send.
