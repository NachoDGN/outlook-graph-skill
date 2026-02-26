# App Registration Setup (Manual)

Use this guide to configure Microsoft Entra app registration for delegated Outlook access.

## 1. Create app registration

1. Open Azure Portal: https://portal.azure.com
2. Go to Microsoft Entra ID -> App registrations -> New registration.
3. Name: `codex-outlook-graph` (or your preferred name).
4. Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**.
5. Redirect URI (Public client/native): `http://localhost:8765`.
6. Create app and copy the **Application (client) ID**.

## 2. Configure authentication

1. Open app -> Authentication.
2. Ensure the redirect URI above is present.
3. Enable public client/native flow if prompted by tenant policy.
4. Save changes.

## 3. Configure delegated permissions

In API permissions, add delegated Microsoft Graph permissions:

- `User.Read`
- `Mail.ReadWrite`
- `Mail.Send`

Grant admin consent where organization policy requires it.

## 4. Set local environment variables

```bash
export OUTLOOK_CLIENT_ID="your-client-id"
export OUTLOOK_TENANT_ID="common"
export OUTLOOK_REDIRECT_URI="http://localhost:8765"
export OUTLOOK_SCOPES="User.Read Mail.ReadWrite Mail.Send"
```

MSAL Python reserved scopes (`openid`, `profile`, `offline_access`) must not be passed in `OUTLOOK_SCOPES`.

## 5. Validate login

```bash
python3 scripts/outlook_cli.py auth login --method browser --profile default
python3 scripts/outlook_cli.py auth status --profile default
```

## Common setup failures

- `invalid_client`: wrong client ID.
- `AADSTS50011 redirect_uri mismatch`: redirect URI in env does not match app registration.
- `insufficient privileges`: missing delegated permissions or consent.
- `interaction_required`: user must re-run interactive login.
