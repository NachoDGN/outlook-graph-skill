---
name: outlook-graph
description: Authenticate and automate Outlook mailbox workflows through Microsoft Graph. Use when users ask to connect an Outlook account, list or summarize unread inbox emails, mark messages read or unread, draft or send emails with confirmation, or download email attachments such as PDFs and other binaries.
---

# Outlook Graph

## Overview

Use this skill to access Outlook mailboxes with delegated Microsoft Graph auth for a user-selected account. It supports browser login and device code login, batched mailbox retrieval, draft-first email writing, explicit-confirmation send, and binary attachment downloads.

## Quick start

1. Install script dependencies:

```bash
python3 -m pip install --user -r scripts/requirements.txt
```

2. Configure required environment variables:

```bash
export OUTLOOK_CLIENT_ID="your-app-client-id"
export OUTLOOK_TENANT_ID="common"
export OUTLOOK_REDIRECT_URI="http://localhost:8765"
export OUTLOOK_SCOPES="User.Read Mail.ReadWrite Mail.Send"
export OUTLOOK_PROFILE="default"
export OUTLOOK_OUTPUT_DIR="./outlook_downloads"
```

Do not include `openid`, `profile`, or `offline_access` in `OUTLOOK_SCOPES` when using MSAL Python. Those are reserved OIDC scopes.

3. Create or verify the Microsoft app registration by following [app registration setup](references/app_registration.md).

4. Authenticate:

```bash
python3 scripts/outlook_cli.py auth login --method browser --profile default
```

For headless sessions, use device code:

```bash
python3 scripts/outlook_cli.py auth login --method device --profile default
```

## Core commands

### Mail read and summarization workflows

List recent unread inbox messages in one call:

```bash
python3 scripts/outlook_cli.py mail list --folder inbox --unread-only --top 10
```

Fetch one full message:

```bash
python3 scripts/outlook_cli.py mail get --message-id MESSAGE_ID
```

Mark read or unread:

```bash
python3 scripts/outlook_cli.py mail mark --message-id MESSAGE_ID --read true
python3 scripts/outlook_cli.py mail mark --message-id MESSAGE_ID --read false
```

### Draft and send with guardrails

Create a draft:

```bash
python3 scripts/outlook_cli.py mail draft \
  --to recipient@example.com \
  --subject "Subject" \
  --body-file ./body.txt
```

Send an existing draft only after explicit confirmation:

```bash
python3 scripts/outlook_cli.py mail send-draft --message-id MESSAGE_ID --confirm-send
```

### Attachment workflows

List message attachments:

```bash
python3 scripts/outlook_cli.py attachments list --message-id MESSAGE_ID
```

Download a specific attachment:

```bash
python3 scripts/outlook_cli.py attachments download \
  --message-id MESSAGE_ID \
  --attachment-id ATTACHMENT_ID \
  --output-dir ./outlook_downloads
```

Download all attachments from a message:

```bash
python3 scripts/outlook_cli.py attachments download-all --message-id MESSAGE_ID
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
