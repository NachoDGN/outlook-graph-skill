# Operations Guide

## Auth flows

Browser flow:

```bash
python3 scripts/outlook_cli.py auth login --method browser --profile default
```

Device flow:

```bash
python3 scripts/outlook_cli.py auth login --method device --profile default
```

Check session status:

```bash
python3 scripts/outlook_cli.py auth status --profile default
```

## Unread inbox summary pipeline

1. Fetch unread messages:

```bash
python3 scripts/outlook_cli.py mail list --folder inbox --unread-only --top 20
```

2. Summarize using message fields (`subject`, `from`, `bodyPreview`, `receivedDateTime`).

3. Optionally mark processed messages as read:

```bash
python3 scripts/outlook_cli.py mail mark --message-id MESSAGE_ID --read true
```

## Draft and send workflow

1. Draft:

```bash
python3 scripts/outlook_cli.py mail draft \
  --to manager@example.com teammate@example.com \
  --subject "Weekly update" \
  --body-file ./weekly-update.txt
```

2. Review message metadata and draft ID.
3. Send only after explicit confirmation:

```bash
python3 scripts/outlook_cli.py mail send-draft --message-id MESSAGE_ID --confirm-send
```

## Attachment download workflow

List attachments:

```bash
python3 scripts/outlook_cli.py attachments list --message-id MESSAGE_ID
```

Download one:

```bash
python3 scripts/outlook_cli.py attachments download \
  --message-id MESSAGE_ID \
  --attachment-id ATTACHMENT_ID
```

Download all:

```bash
python3 scripts/outlook_cli.py attachments download-all --message-id MESSAGE_ID
```

## Multi-profile usage

Use different profiles per mailbox/account:

```bash
python3 scripts/outlook_cli.py auth login --method browser --profile work
python3 scripts/outlook_cli.py auth login --method browser --profile personal
python3 scripts/outlook_cli.py mail list --unread-only --profile work
```

## Output handling

- Default JSON is sorted and stable for downstream agent parsing.
- `--format text` is intended for manual terminal use.
