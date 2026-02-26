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

- Browser auth works and session persists for selected profile.
- Device code auth works in terminal-only environment.
- `mail list --unread-only --top 10` returns expected fields and ordering.
- `mail mark` updates `isRead` to requested value.
- `mail draft` creates a draft and returns message ID.
- `mail send-draft` fails without `--confirm-send` and succeeds with it.
- Attachment download writes non-empty files and stable metadata output.

## Robustness tests

- Missing `OUTLOOK_CLIENT_ID` returns actionable config error.
- Missing dependencies (`msal`, `requests`) return install guidance.
- Invalid message ID returns Graph API error payload.
- Attachment name sanitization prevents invalid path characters.

## Coverage notes

- Test with 10-20 paraphrased prompts across trigger and non-trigger sets.
- Include at least one mixed-intent prompt to ensure the skill activates only for Outlook tasks.
- Track under-triggering and over-triggering after real usage and refine description keywords.
