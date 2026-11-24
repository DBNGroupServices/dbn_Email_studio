# DBN Email Studio

A standalone Windows email writing assistant. Open the app, drop in an email or screenshot, add your dot points and create an editable response. It runs independently of ChatGPT and Codex.

Version 1.1 refreshes the interface with a light workspace, blue accents, consistent line icons, clearer task cards and more readable draft text. The design takes inspiration from [Devoq's modern SaaS UI article](https://www.devoq.io/modern-saas-website-design-best-ui-ideas-for-software-product-websites/).

Version 1.5 adds a personal Outlook web add-in preview for writing, summaries, action items, proofreading, translation, follow-up drafts and saved templates. The standalone app retains its Outlook email picker, existing ChatGPT plan connection and blue Windows icon. See [Outlook add-in setup](office-addin/README.md) and the [feature roadmap](office-addin/FEATURE-ROADMAP.md). The add-in requires a separate installation in Outlook; preparing its local HTTPS service does not install it.

Version 1.5.1 fixes add-in URLs containing Outlook host parameters, permits embedding from Outlook's current cloud address, and reconnects existing windows to a restarted local service without clearing draft inputs. These server changes do not require reinstalling the manifest.

Version 1.5.2 fixes the disabled Outlook draft action when no email body can be read. A panel opened in a browser can create a new editable Outlook Classic email with no recipients. Inside Outlook, replies and insertion remain tied to the email used for the result. Changing the selected email or a failed generation cannot redirect an earlier result to another message.

## Start here

1. Double-click **DBN Email Studio.exe** in the `dist` folder, or use **Open DBN Email Studio.cmd** in the project folder.
2. Open **Settings & writing style**. Leave **My existing ChatGPT plan** selected and choose **Continue with ChatGPT**.
3. Sign in on OpenAI's website and enable **Use your ChatGPT plan**. Return to Email Studio and acknowledge the plan-use explanation. Models are loaded from that account. **Test drafting with fictional text** checks generation using a small amount of your plan allowance. Choose **Save settings** for writing preferences.
4. Add your name, signature and writing preferences once. These apply to every email.

Your ChatGPT connection uses eligible subscription allowance or available ChatGPT credits, subject to your plan and app limits. **Manage usage** opens ChatGPT usage settings. No additional Email Studio subscription is required. This personal app runs locally and requires internet for drafting. Sign-in does not provide access to your ChatGPT conversations or memories.

Separate OpenAI API billing remains an explicit option under **Use for drafting**. It requires your own API key and funded API account. The app never changes billing routes automatically when a limit or connection error occurs. Existing saved API keys are retained but are not used while ChatGPT plan is selected.

Sign-in uses the documented dynamic local-client registration, state, nonce and PKCE flow. ID-token signatures, issuer, audience and expiry are verified. Credentials are encrypted with your Windows account; refresh is serialized, and sign-out attempts remote revocation before removing local tokens. Identity-only sign-in without the plan permission cannot draft.

Official integration details: https://developers.openai.com/siwc/token-sharing-open-source
Eligibility is determined by OpenAI and the selected account/workspace. This build is for personal/local use. Paid or remotely hosted distribution requires separate eligibility review.

Version 1.2.2 distinguishes an empty API balance from organisation and project spend limits. Error details show the organisation, project, error code and request ID returned by OpenAI. If your billing page shows credits but OpenAI rejects the key, compare the organisation and project with the key's details. A model-access check alone does not verify drafting or billing access.

## Everyday use

- **Reply to an email:** paste an incoming email, attach a screenshot or saved email, and add the points you want your reply to cover.
- **Rewrite a draft:** paste your draft and describe any changes. Choose tone, length and paragraph or bullet layout.
- **From dot points:** enter the points for a new email. Reference files are optional.
- Paste screenshots with **Ctrl V**, choose files or drag files onto the upload area.
- Click an attachment to inspect the text extracted locally or preview the image.
- Edit the resulting subject and email directly. Choose **Copy email**, **Copy with subject**, or download a text file.
- Use **Shorter**, **Warmer**, **Firmer** or **Simpler**, or type a specific revision. The app uses your latest edited draft and the source material still in the workspace.
- Use **New email** to clear the workspace. Your standing writing preferences stay saved.

Supported uploads: PNG, JPG/JPEG, WEBP, EML, Outlook MSG, PDF, DOCX, TXT and HTML. Up to 10 files, 10 MB per file and 25 MB total, including expanded email attachments. Encrypted/password-protected files must be unlocked first. Save older Word `.doc` files as PDF or DOCX. Direct dragging of Outlook emails depends on your Outlook version; save the message as `.msg` or `.eml` first if dragging does not produce a file.

Images are detected by their actual contents. A JPEG named `image.png`, which can occur in Outlook threads, is read as JPEG. An unreadable embedded image is listed for review without discarding the email body; an unreadable image uploaded by itself is rejected clearly.

## Outlook Classic shortcuts

Open Outlook Classic on this Windows computer. In Email Studio choose **Import from Outlook Classic** to open the email picker. It starts in Outlook's current mail folder, or the default Inbox if no mail folder is open. Browse mailboxes and folders, search by subject or sender, and use Previous/Next for older messages. Select one email and press **Import selected email**. Import stays disabled until a message is selected.

The picker reads up to 50 email headers per page, newest first, without displaying messages or marking them as read. Search applies to the current folder. Selecting a different folder clears the search. Messages over the 10 MB limit cannot be selected. Cancel keeps your workspace; importing replaces the current source/draft only after confirmation when needed, keeps your dot points and switches to Reply. No mailbox content is sent to AI by browsing or importing.

After generating and editing your response, choose **Open reply in Outlook Classic**. For an email imported from Outlook during the current session, this opens a reply to that original sender with your wording above the quoted thread. For other sources, **Open in Outlook Classic** opens a new email for you to address. Review the recipients, wording and any duplicate signature before sending. The app never calls Outlook's Send method.

This is an on-demand local connection to the already-running Outlook Classic application, rather than a ribbon add-in. It does not continuously sync or index the mailbox. It needs Windows PowerShell and permitted local Office automation; it does not change Outlook security settings. If Outlook or your organisation blocks the connection, use saved MSG files and Copy email instead. New Outlook is not supported by this local COM connection.

The import briefly exports the selected email to a private temporary folder, reads it and removes that copy. Only source identifiers needed for a reply are held in server memory until the app closes; these are not added to saved history. Reimport a moved/deleted message or after restarting the app. Files dropped or uploaded normally do not carry this live Outlook reply link.

EML and MSG text and supported embedded attachments are read locally. Unsupported embedded files are listed visibly. Embedded Outlook messages should be saved and attached separately. DOCX text is extracted locally; Word images should be attached separately or the document exported to PDF. Images and PDF pages are read by the selected AI. Unclear details are shown separately for review when identified by the AI. Always review names, amounts, dates and commitments before using a draft.

## Models

- **GPT-6.1 Sol** is the default for everyday quality and cost balance.
- **GPT-6 Astra** is available for the most demanding responses, at a higher API cost.
- **GPT-6.1 Terra** is available for quick, lower-cost writing.

For ChatGPT plan use, the model picker is populated from the selected account's live catalog. The choices above are the separate API connection's defaults. The app does not silently switch models or billing routes, and does not retry generation automatically. The implementation uses the public Responses API with image and file inputs when supported by the selected model.

## Local data and privacy

The app listens only on `127.0.0.1`. It opens in a Microsoft Edge app window when Edge is installed, or your default browser otherwise. You do not need Python, Node.js, a terminal, ChatGPT or Codex to run the packaged executable.

Your OAuth tokens, any optional API key, preferences and draft history are encrypted with Windows DPAPI for your current Windows account and stored in `%LOCALAPPDATA%\DBN Email Studio`. Tokens and API keys are used only by the local server over verified HTTPS and are never returned to the interface. Do not place credentials in the shared project folder. Sign-in and renewal use `auth.openai.com`; model discovery and drafting use `api.openai.com`.

Source text, instructions, current draft, saved writing identity and attached content are sent to OpenAI when you create or refine a draft. The app sets `store: false`; OpenAI's applicable service policies still apply. The optional drafting test sends a short fictional prompt using the selected connection and model, without your emails or saved identity/preferences. The app does not log emails or credentials.

Uploaded files and source text are kept in memory, not saved by the app. Windows or your browser may manage caches or temporary data independently. Optional local history saves only the latest 30 generated draft results. Turn it off in Settings to stop new saves. Clear existing history under **Recent drafts → Clear history**. A restored draft does not restore its original source or attachments; reattach those if a revision needs them. Only generated versions are saved automatically; copy manual edits before closing the window.

The app compiles drafts for you to use in Outlook or another email program. Outlook folder names and email headers are read locally when you open or navigate the picker. The full message and attachments are exported only for the email you select to import. No messages are sent automatically. To completely stop the local server, choose **Settings → Quit app**. Closing the browser window leaves the local app available for a quick reopening.

## Examples and troubleshooting

The `examples` folder contains fictional Outlook and EML emails and dot points. **Try an example** fills a sample request; creating its reply uses the real connection and allowance selected in Settings.

- **Key rejected:** check the API key in Settings, rather than a ChatGPT password.
- **Connection error:** choose **Settings → Check network**. This checks the app's verified HTTPS connection to OpenAI without using an API key, email content or paid generation. Errors distinguish DNS, certificates, Windows network restrictions and timeouts. Certificate verification is always enabled.
- **ChatGPT usage limit:** choose Manage usage to review your plan or this app's limit.
- **ChatGPT connection expired:** choose Continue with ChatGPT to reconnect.
- **Separate API billing error:** check your API billing balance and project limits if you deliberately selected API billing.
- **Model unavailable:** refresh models or select another listed model for the connected account.
- **Blank/unreadable source:** use a clearer screenshot or unlocked PDF.
- **App already running:** open the executable again; it reuses the running local app.
- **Port in use:** launch with `"DBN Email Studio.exe" --port 8768`.

## Source and rebuilding

`app.py` is the backend. `web` contains the interface. `tests` contains local integration tests with mocked AI calls. `build.ps1` rebuilds the standalone executable using Python 3.12+ and PyInstaller. Build dependencies are listed in `requirements-build.txt`; the app uses extract-msg for saved Outlook files. Run `python -m unittest discover -s tests -v` for backend tests and `node --check web/app.js` for JavaScript syntax validation. `tests/preview_server.py` is a separate developer-only simulated-AI test server; it is never included in the executable.

Connection security and drafting workflows are covered by isolated tests and simulated responses. Live ChatGPT-plan generation requires the user to complete OpenAI sign-in and consent. See TEST-RESULTS.md for the exact validation completed for each release.
