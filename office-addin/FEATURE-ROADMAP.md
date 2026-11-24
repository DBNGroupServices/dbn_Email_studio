# Product feature scope — 28 October 2026

This is an independent DBN implementation of requested capabilities, not an integration with or copy of another vendor's proprietary product. Public product descriptions are scope references; they do not establish feature parity. No claim is made to reproduce every present or future feature across seven separate services.

## Reference review

| Product | Capabilities reviewed | Primary source |
|---|---|---|
| Microsoft Copilot | Drafting, thread summaries, writing feedback, email triage, rules and calendar assistance. Client availability differs. | [Microsoft Outlook FAQ](https://support.microsoft.com/en-us/outlook/frequently-asked-questions-about-copilot-in-outlook) |
| MailMaestro / Maestro Labs | Draft/reply/improve, tone, thread/attachment summaries, labels, inbox questions and reminders; TeamsMaestro and scheduling are separate products in the suite. | [MailMaestro email management](https://www.maestrolabs.com/ai-email-management), [Maestro Labs products](https://www.maestrolabs.com/) |
| Lindy | Email/calendar assistance, workflows, integrations, memory, scheduled routines, meeting notes and follow-ups. | [Lindy product](https://www.lindy.ai/product) |
| Shortwave | AI drafting, summaries, mailbox questions, search, organisation, calendar, custom instructions and integrations. | [Shortwave assistant guide](https://www.shortwave.com/docs/guides/ai-assistant/) |
| Superhuman Mail | AI writing, instant replies, voice matching, snippets, split inbox, labels, auto drafts, reminders, snooze, send later, collaboration and integrations. | [Superhuman feature table](https://superhuman.com/plans/mail) |
| Spark | Smart/unified inbox, priority senders, grouping, gatekeeper, search, summaries, writing/translation and calendar/meeting context. | [Spark features](https://sparkmailapp.com/features), [AI assistant](https://sparkmailapp.com/features/ai-assistant) |
| Fyxer | Categorisation, personal writing style, reply drafts, meeting notes, scheduling, follow-ups, team/admin controls and connected tools. | [Fyxer help](https://support.fyxer.com/article/meet-fyxer-your-ai-email-and-meeting-assistant) |

## Delivery stages and current status

| Stage | Features | Status / dependency |
|---|---|---|
| 1 — Current item | Outlook read/compose ribbon, task pane, local plan connection, reply/rewrite/compose, subject/body review, tone/length, selected text, screenshot/files, current attachments, summaries, actions, proofreading, translation, follow-up drafting, supplied meeting-note structuring, templates, insertion/reply | Implemented as a personal add-in preview. Live Outlook installation requires certificate trust and custom add-in consent. Office API adapter is tested with fixtures; actual Outlook round-trip must be verified after installation. |
| 2 — Inbox control | Priority categories, VIPs, bulk triage, filing/archiving, flagging, rules, sender grouping, unified accounts, unsubscribe/block, inbox digest, searchable email knowledge | Planned. Decide local COM versus Microsoft Graph access. Add reversible action previews and exclusions before enabling mailbox mutations. |
| 3 — Time and follow-up | Waiting-for-reply tracking, reminders, snooze, send later, scheduling, availability and bookings, daily briefing, user-triggered automation | Planned. Needs a durable local worker or approved service; calendar/mail permissions and delivery behaviour require user choices. |
| 4 — Meetings | Calendar context, imported transcripts, summaries, decisions, assigned tasks, post-meeting draft; later live recording/notetaker | Imported text structuring implemented. Live attendance, capture and transcription are not implemented. They require a supported capture/transcription path and meeting consent; the current ChatGPT plan HTTP route does not provide audio/video input. |
| 5 — Team and external tools | Shared templates/style, collaboration, assignments, CRM, Slack/Teams, task tools, cross-app workflows, central policies/audits and administration | Planned after the personal version. Needs supported external connections, hosting and separate permissions. |

## Decisions still needed for later stages

- Which inbox mutations may be automated versus reviewed first.
- Whether tools should work only while this computer runs or via an approved hosted worker.
- Which calendar, task, meeting and CRM providers to connect.
- Whether read tracking, automatic sending, meeting recording and cross-account indexing are actually wanted. They are separate opt-in capabilities, not enabled by this build.

The personal preview reuses the user's existing ChatGPT plan. It does not add a paid provider automatically. Later recording, hosting or organisation deployment must be assessed before promising it can run entirely within that allowance.

Architecture references: [Outlook add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-add-ins-overview), [current-item permissions](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/understanding-outlook-add-in-permissions), [ChatGPT plan local-client integration](https://developers.openai.com/siwc/token-sharing-open-source), [plan-route input limits](https://developers.openai.com/siwc/token-sharing-open-source/preview-limitations).
