---
name: outlook-graph
description: Authenticate and automate Outlook mailbox workflows through Microsoft Graph. Use when users ask to connect an Outlook account, list or summarize unread inbox emails, mark messages read or unread, draft or send emails with confirmation, or download email attachments such as PDFs and other binaries.
---

# Outlook Graph

## Overview

Use this skill to access Outlook mailboxes with delegated Microsoft Graph auth for a user-selected account. It supports browser login and device code login, batched mailbox retrieval, draft-first email writing, explicit-confirmation send, and binary attachment downloads.

## Quick start

1. Create and activate a virtual environment:

```bash
python3 -m venv venv
source venv/bin/activate
```

2. If the agent is running outside the skill folder, resolve an absolute CLI path first. In Hermes runtimes prefer `HERMES_HOME`; otherwise fall back to `CODEX_HOME`:

```bash
export HERMES_HOME="${HERMES_HOME:-$HOME/.hermes}"
export CODEX_HOME="${CODEX_HOME:-$HOME/.codex}"
if [ -f "$HERMES_HOME/skills/outlook-graph/scripts/outlook_cli.py" ]; then
  export OUTLOOK_CLI="$HERMES_HOME/skills/outlook-graph/scripts/outlook_cli.py"
else
  export OUTLOOK_CLI="$CODEX_HOME/skills/outlook-graph/scripts/outlook_cli.py"
fi
```

3. Install script dependencies (venv-safe default):

```bash
python3 -m pip install -r scripts/requirements.txt
```

If you are intentionally using system Python (no active virtualenv), user-site install is also valid:

```bash
python3 -m pip install --user -r scripts/requirements.txt
```

4. Mandatory after activating the venv: pin the interpreter path used by this skill:

```bash
python3 "$OUTLOOK_CLI" auth pin-interpreter --profile default
```

This keeps macOS Keychain trust bound to one Python executable path and prevents repeated password prompts.

5. Configure required environment variables:

```bash
export OUTLOOK_CLIENT_ID="your-app-client-id"
export OUTLOOK_TENANT_ID="common"
export OUTLOOK_REDIRECT_URI="http://localhost:8765"
export OUTLOOK_SCOPES="User.Read Mail.ReadWrite Mail.Send"
export OUTLOOK_PROFILE="default"
export OUTLOOK_OUTPUT_DIR="./outlook_downloads"
```

Do not include `openid`, `profile`, or `offline_access` in `OUTLOOK_SCOPES` when using MSAL Python. Those are reserved OIDC scopes.

6. Create or verify the Microsoft app registration by following [app registration setup](references/app_registration.md).

7. Authenticate:

```bash
python3 "$OUTLOOK_CLI" auth login --method browser --profile default
```

For headless sessions, use device code:

```bash
python3 "$OUTLOOK_CLI" auth login --method device --profile default
```

## Mandatory agent onboarding protocol

When a user asks to read/write emails and auth is not yet established, always execute this sequence before any mail command:

1. Run onboarding planner:

```bash
python3 "$OUTLOOK_CLI" auth onboard --profile default
```

2. Ask the user only for missing fields listed under `questions_for_user`.
3. Explain and confirm `required_user_actions` in plain language.
4. Run the exact `login_command` returned by onboarding planner.
5. Wait for user to complete browser/device sign-in and consent.
6. Run `status_command` returned by onboarding planner.
7. Continue with mailbox operations only if `authenticated` is `true`, and prefer the returned `first_mail_command` so the path is always correct.

This flow is designed so a non-technical user only performs one-time app setup and first login.

## Core commands

### Mail read and summarization workflows

Discover Inbox subtree (tree + flat index):

```bash
python3 "$OUTLOOK_CLI" folders tree --root inbox
```

List messages in a nested folder by resolved path:

```bash
python3 "$OUTLOOK_CLI" mail list --folder-path "Inbox/Finance/Invoices" --top 20
```

List messages by explicit folder ID:

```bash
python3 "$OUTLOOK_CLI" mail list --folder-id FOLDER_ID --top 20
```

List recent unread inbox messages in one call:

```bash
python3 "$OUTLOOK_CLI" mail list --folder inbox --unread-only --top 10
```

Fetch one full message:

```bash
python3 "$OUTLOOK_CLI" mail get --message-id MESSAGE_ID
```

Mark read or unread:

```bash
python3 "$OUTLOOK_CLI" mail mark --message-id MESSAGE_ID --read true
python3 "$OUTLOOK_CLI" mail mark --message-id MESSAGE_ID --read false
```

### Draft and send with guardrails

Create a draft:

```bash
python3 "$OUTLOOK_CLI" mail draft \
  --to recipient@example.com \
  --subject "Subject" \
  --body-file ./body.txt
```

Send an existing draft only after explicit confirmation:

```bash
python3 "$OUTLOOK_CLI" mail send-draft --message-id MESSAGE_ID --confirm-send
```

### Attachment workflows

List message attachments:

```bash
python3 "$OUTLOOK_CLI" attachments list --message-id MESSAGE_ID
```

Download a specific attachment:

```bash
python3 "$OUTLOOK_CLI" attachments download \
  --message-id MESSAGE_ID \
  --attachment-id ATTACHMENT_ID \
  --output-dir ./outlook_downloads
```

Download all attachments from a message:

```bash
python3 "$OUTLOOK_CLI" attachments download-all --message-id MESSAGE_ID
```

Batch download attachments from latest messages (default top 10):

```bash
python3 "$OUTLOOK_CLI" attachments download-recent --folder inbox --top 10
```

Incremental download of only new attachments (stateful ledger):

```bash
python3 "$OUTLOOK_CLI" attachments download-new --folder inbox
```

Inspect or reset per-folder download state:

```bash
python3 "$OUTLOOK_CLI" attachments state status --folder inbox
python3 "$OUTLOOK_CLI" attachments state reset --folder inbox --confirm-reset
```

## Output format

- Default output is deterministic JSON for agent parsing.
- Use `--format text` for a human-readable format.

## Critical behavior rules

- Require explicit `--confirm-send` before sending drafts.
- Never hardcode mailbox accounts; select account via delegated user auth and profile.
- Never print credentials, refresh tokens, or access tokens.
- On auth errors, report actionable steps (login again, verify scopes, verify app config).

## References

- [App registration and consent setup](references/app_registration.md)
- [Operations and command patterns](references/operations.md)
- [Security review checklist](references/security-review.md)
- [Evaluation test suite](references/evaluation.md)
