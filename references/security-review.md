# Security Review Checklist

## Baseline controls

- Use delegated OAuth only (no ROPC, no stored user password).
- Require explicit send confirmation (`--confirm-send`).
- Avoid logging access tokens, refresh tokens, or secrets.
- Use keyring-backed token cache when available.
- Fall back to file cache with strict permissions (`700` dir, `600` files).

## Permission and scope review

- Ensure least-privilege delegated scopes:
  - `User.Read`
  - `Mail.ReadWrite`
  - `Mail.Send`
- Review whether `Mail.Read` or `Mail.ReadWrite` is sufficient for your use case.
- Confirm tenant admin consent policy and user consent expectations.
- Do not pass reserved OIDC scopes (`openid`, `profile`, `offline_access`) in `OUTLOOK_SCOPES` for MSAL Python.

## Risk tiers

- Low risk: read-only mailbox listing and summarization.
- Medium risk: message state changes (mark read/unread).
- High risk: send operations on behalf of user.

## Operational safeguards

- Require draft-first workflow in user-facing runbooks.
- Use profile separation for personal vs work accounts.
- Validate output directory permissions for attachment downloads.
- Sanitize attachment filenames before writing to disk.

## Incident response cues

- Revoke app consent for compromised account sessions.
- Rotate app registration and force re-auth when suspicious behavior appears.
- Remove local token caches (`auth logout`) on shared machines.
