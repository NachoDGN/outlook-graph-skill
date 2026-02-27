# Operations Guide

From any working directory, resolve CLI path first:

```bash
export HERMES_HOME="${HERMES_HOME:-$HOME/.hermes}"
export CODEX_HOME="${CODEX_HOME:-$HOME/.codex}"
if [ -f "$HERMES_HOME/skills/outlook-graph/scripts/outlook_cli.py" ]; then
  export OUTLOOK_CLI="$HERMES_HOME/skills/outlook-graph/scripts/outlook_cli.py"
else
  export OUTLOOK_CLI="$CODEX_HOME/skills/outlook-graph/scripts/outlook_cli.py"
fi
```

## Deterministic onboarding flow (agent-first)

1. Generate setup plan:

```bash
python3 "$OUTLOOK_CLI" auth onboard --profile default
```

2. If `questions_for_user` is non-empty, ask those exact questions.
3. Run `login_command` returned by onboarding output.
4. Ask user to complete sign-in/consent in browser.
5. Run `status_command` and confirm `authenticated=true`.
6. Only then run mailbox commands.

## Auth flows

Browser flow:

```bash
python3 "$OUTLOOK_CLI" auth login --method browser --profile default
```

Device flow:

```bash
python3 "$OUTLOOK_CLI" auth login --method device --profile default
```

Check session status:

```bash
python3 "$OUTLOOK_CLI" auth status --profile default
```

## Unread inbox summary pipeline

1. Fetch unread messages:

```bash
python3 "$OUTLOOK_CLI" mail list --folder inbox --unread-only --top 20
```

2. Summarize using message fields (`subject`, `from`, `bodyPreview`, `receivedDateTime`).

3. Optionally mark processed messages as read:

```bash
python3 "$OUTLOOK_CLI" mail mark --message-id MESSAGE_ID --read true
```

## Draft and send workflow

1. Draft:

```bash
python3 "$OUTLOOK_CLI" mail draft \
  --to manager@example.com teammate@example.com \
  --subject "Weekly update" \
  --body-file ./weekly-update.txt
```

2. Review message metadata and draft ID.
3. Send only after explicit confirmation:

```bash
python3 "$OUTLOOK_CLI" mail send-draft --message-id MESSAGE_ID --confirm-send
```

## Attachment download workflow

List attachments:

```bash
python3 "$OUTLOOK_CLI" attachments list --message-id MESSAGE_ID
```

Download one:

```bash
python3 "$OUTLOOK_CLI" attachments download \
  --message-id MESSAGE_ID \
  --attachment-id ATTACHMENT_ID
```

Download all:

```bash
python3 "$OUTLOOK_CLI" attachments download-all --message-id MESSAGE_ID
```

## Multi-profile usage

Use different profiles per mailbox/account:

```bash
python3 "$OUTLOOK_CLI" auth login --method browser --profile work
python3 "$OUTLOOK_CLI" auth login --method browser --profile personal
python3 "$OUTLOOK_CLI" mail list --unread-only --profile work
```

## Output handling

- Default JSON is sorted and stable for downstream agent parsing.
- `--format text` is intended for manual terminal use.
