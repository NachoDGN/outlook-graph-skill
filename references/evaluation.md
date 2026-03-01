# Evaluation Suite

Use this suite to validate trigger quality, behavior, and reliability.

## Trigger tests

### Should trigger

- "Summarize my last unread Outlook emails"
- "Download the PDF attachment from this email"
- "Draft a reply to this sender in Outlook"
- "Mark these unread Outlook messages as read"

### Should not trigger

- "What is the weather in New York?"
- "Help me debug this Python stack trace"
- "Create a spreadsheet budget template"

## Functional tests

- `auth onboard` returns deterministic `questions_for_user`, `required_user_actions`, `login_command`, and `status_command`.
- Browser auth works and session persists for selected profile.
- Device code auth works in terminal-only environment.
- `mail list --unread-only --top 10` returns expected fields and ordering.
- `folders tree --root inbox` returns both nested tree and flat index rows with deterministic paths.
- `mail list --folder-path "Inbox/Some/Subfolder"` resolves target folder and returns `resolved_folder_id`.
- `mail list --folder-id FOLDER_ID` fetches from explicit folder id without path resolution.
- `mail mark` updates `isRead` to requested value.
- `mail draft` creates a draft and returns message ID.
- `mail send-draft` fails without `--confirm-send` and succeeds with it.
- Attachment download writes non-empty files and stable metadata output.
- `attachments download-recent --top 10` scans recent messages and downloads attachments with per-message summary.
- First run of `attachments download-new` backfills only last 15 days.
- Second run of `attachments download-new` skips already downloaded attachments using state ledger.
- Failed attachment entries are retained in `pending_failures` and retried on future `download-new` runs.
- `attachments state status` reports stream metadata and counters.
- `attachments state reset --confirm-reset` clears state for the selected folder stream.

## Robustness tests

- Missing `OUTLOOK_CLIENT_ID` returns actionable config error.
- Missing dependencies (`msal`, `requests`) return install guidance.
- Invalid message ID returns Graph API error payload.
- Attachment name sanitization prevents invalid path characters.
- Folder traversal max-node guard raises an actionable error when exceeded.
- Folder path resolution fails deterministically on missing/ambiguous segments.
- `download-new` behavior is independent of `isRead` state.

## Coverage notes

- Test with 10-20 paraphrased prompts across trigger and non-trigger sets.
- Include at least one mixed-intent prompt to ensure the skill activates only for Outlook tasks.
- Track under-triggering and over-triggering after real usage and refine description keywords.
