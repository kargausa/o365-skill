# Microsoft 365 skill — Azure CLI setup playbook

End-to-end registration of the Azure AD app for this skill using `az` only —
no Azure portal clicks. Complements `README.md` (portal walkthrough) and
`SKILL.md` (LLM usage contract).

## Prerequisites

- `az` CLI logged in (`az login`) to the **tenant that owns the user's
  mailbox** — not the subscription tenant if those differ.
- Permission to register apps in that tenant (regular user is enough unless
  the tenant has disabled user-registered apps).
- The Hermes Python venv (or any Python 3.11+) with `msal` and `requests`
  installed.

## Decisions to make first

1. **Target tenant** — the home tenant of the user whose Outlook/Calendar
   /Teams/OneDrive you'll access. Find it with:
   ```bash
   az account list --query "[].{sub: name, tenantId: tenantId, user: user.name}" -o table
   ```
2. **App name** — pick something neutral (e.g. `${USER}-m365-cli`); do not
   embed agent or product names.
3. **Sign-in audience** — `AzureADMyOrg` (single-tenant) is the secure
   default. Only use `AzureADMultipleOrgs` if you genuinely need to sign in
   users from other tenants.
4. **User to authorize as** — a **regular, non-privileged** account whose
   mailbox you actually want to use. Never authorize with a Global Admin or
   any other privileged account; the resulting refresh token in
   `~/.hermes/o365_token_cache.bin` is long-lived and would expose those
   privileges to anyone who reads the file.

## Variables

```bash
TENANT_ID="00000000-0000-0000-0000-000000000000"   # your Entra tenant ID (mailbox home tenant)
APP_NAME="m365-cli"                                 # Entra app registration display name (any name)
SKILL_DIR="${HOME}/.hermes/skills/productivity/o365"   # or wherever the skill lives
PYTHON="${HOME}/.hermes/hermes-agent/venv/bin/python3"   # Hermes venv interpreter
```

Switch `az` context to the right tenant:

```bash
az account set --subscription "<a sub in that tenant>"
# Optional sanity check:
az account show --query "{user: user.name, tenantId: tenantId}" -o json
```

## Step 1 — Resolve Microsoft Graph delegated permission IDs

The skill needs these delegated scopes. Pull their GUIDs from the Graph
service principal so you don't hard-code them:

```bash
az ad sp show --id 00000003-0000-0000-c000-000000000000 \
  --query "oauth2PermissionScopes[?value=='User.Read' || value=='Mail.Read' \
            || value=='Mail.Send' || value=='Mail.ReadWrite' \
            || value=='Calendars.ReadWrite' || value=='Files.ReadWrite' \
            || value=='Chat.Read' || value=='ChatMessage.Send' \
            || value=='ChannelMessage.Send' \
            || value=='Team.ReadBasic.All' || value=='Channel.ReadBasic.All' \
            || value=='ChannelMessage.Read.All'] \
            .{name:value, id:id}" \
  -o json
```

> `offline_access` is **not** included here — MSAL adds it automatically,
> and passing it explicitly raises `ValueError: ... reserved`.
>
> `Chat.ReadWrite` is also omitted — `ChatMessage.Send` is the narrower
> permission and it's enough to post replies to existing chats. Use
> `Chat.ReadWrite` only if you also need to edit/delete previously-sent
> chat messages, which the skill code does not currently do.
>
> Admin-consent-required scopes in this set: `ChatMessage.Send`,
> `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`.
> If you cannot run Step 5, the corresponding subcommands will return
> HTTP 403; `outlook *`, `calendar *`, `onedrive *`, and `teams chats /
> chat-messages` (`Chat.Read` is auto-consented in most tenants) will
> still work.

## Step 2 — Write the permission manifest

Save the resolved IDs to a manifest file. The order and exact values used
when this playbook was authored are below; re-verify with Step 1 if you
suspect Microsoft has rotated any IDs.

```bash
cat > /tmp/o365-app-permissions.json <<'EOF'
[
  {
    "resourceAppId": "00000003-0000-0000-c000-000000000000",
    "resourceAccess": [
      {"id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d", "type": "Scope"},
      {"id": "570282fd-fa5c-430d-a7fd-fc8dc98a9dca", "type": "Scope"},
      {"id": "e383f46e-2787-4529-855e-0e479a3ffac0", "type": "Scope"},
      {"id": "024d486e-b451-40bb-833d-3e66d98c5c73", "type": "Scope"},
      {"id": "1ec239c2-d7c9-4623-a91a-a9775856bb36", "type": "Scope"},
      {"id": "5c28f0bf-8a70-41f1-8ab2-9032436ddb65", "type": "Scope"},
      {"id": "f501c180-9344-439a-bca0-6cbf209fd270", "type": "Scope"},
      {"id": "116b7235-7cc6-461e-b163-8e55691d839e", "type": "Scope"},
      {"id": "ebf0f66e-9fb1-49e4-a278-222f76911cf4", "type": "Scope"},
      {"id": "485be79e-c497-4b35-9400-0e3fa7f2a5d4", "type": "Scope"},
      {"id": "9d8982ae-4365-4f57-95e9-d6032a4c0b87", "type": "Scope"},
      {"id": "767156cb-16ae-4d10-8f8b-41b657c8c8c8", "type": "Scope"},
      {"id": "7427e0e9-2fba-42fe-b0c0-848c9e6a8182", "type": "Scope"}
    ]
  }
]
EOF
```

Permission IDs in order: `User.Read`, `Mail.Read`, `Mail.Send`,
`Mail.ReadWrite`, `Calendars.ReadWrite`, `Files.ReadWrite`, `Chat.Read`,
`ChatMessage.Send`, `ChannelMessage.Send`, `Team.ReadBasic.All`,
`Channel.ReadBasic.All`, `ChannelMessage.Read.All`, `offline_access`.

(Yes — `offline_access` is listed *here* in the API permissions manifest.
It is allowed at registration time. The thing MSAL rejects is requesting
it as a runtime scope.)

## Step 3 — Create the app

```bash
APP_RESP=$(az ad app create \
  --display-name "$APP_NAME" \
  --sign-in-audience AzureADMyOrg \
  --required-resource-accesses @/tmp/o365-app-permissions.json \
  --query "{appId: appId, objectId: id}" -o json)
echo "$APP_RESP"

APP_ID=$(echo "$APP_RESP" | python3 -c "import json,sys; print(json.load(sys.stdin)['appId'])")
OBJ_ID=$(echo "$APP_RESP" | python3 -c "import json,sys; print(json.load(sys.stdin)['objectId'])")
```

## Step 4 — Enable public client flows + create service principal

Device-code flow requires the app to be a *public client*. The CLI's
`--is-fallback-public-client` flag has been flaky historically, so set it
explicitly with `--set` instead:

```bash
az ad app update --id "$OBJ_ID" --set isFallbackPublicClient=true
az ad sp create --id "$APP_ID" --query "{spObjectId: id, appId: appId}" -o json
```

## Step 5 — Grant tenant-wide admin consent for all delegated scopes

Without this, every user who signs in will be prompted to consent
individually — and if your tenant restricts user consent, sign-in will
fail outright. Grant once, for everyone:

```bash
GRAPH_SP_ID=$(az ad sp show --id 00000003-0000-0000-c000-000000000000 --query id -o tsv)
SP_ID=$(az ad sp show --id "$APP_ID" --query id -o tsv)

az rest --method POST \
  --uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants" \
  --headers "Content-Type=application/json" \
  --body "{
    \"clientId\": \"$SP_ID\",
    \"consentType\": \"AllPrincipals\",
    \"resourceId\": \"$GRAPH_SP_ID\",
    \"scope\": \"User.Read Mail.Read Mail.Send Mail.ReadWrite Calendars.ReadWrite Files.ReadWrite Chat.Read ChatMessage.Send ChannelMessage.Send Team.ReadBasic.All Channel.ReadBasic.All ChannelMessage.Read.All offline_access\"
  }"
```

Granting admin consent requires `Application.ReadWrite.All` or a directory
role like *Cloud Application Administrator*. If you don't have it, ask a
tenant admin to run just this step; everything before it is normal-user
territory.

## Step 6 — Persist the tenant ID for Hermes

The skill's MSAL authority defaults to `/common` when `O365_TENANT_ID` is
unset, which fails for single-tenant apps with
`AADSTS50059: No tenant-identifying information found`. Pin the tenant for
all future Hermes invocations:

```bash
printf '\n# Microsoft 365 (o365 skill)\nO365_TENANT_ID=%s\n' "$TENANT_ID" \
  >> ~/.hermes/.env
```

## Step 7 — Device-code authentication

Run as the regular user whose mailbox you want to use. **Do not** sign in
with a privileged or admin account.

```bash
O365_TENANT_ID="$TENANT_ID" PYTHONUNBUFFERED=1 \
  "$PYTHON" -u "$SKILL_DIR/scripts/setup.py" --auth "$APP_ID"
```

The script prints a URL and a code. Open the URL in a private browser
session you control, paste the code, and sign in. The script polls
silently until you finish or the code expires (~15 min). On success:

```
OK: Authenticated as <Display Name>
Token cache saved to ~/.hermes/o365_token_cache.bin
```

> **`PYTHONUNBUFFERED=1` matters** when running the script in the
> background or piped to a log — without it, the URL/code stays buffered
> in the Python process and you'll see nothing until the process exits.

## Step 8 — Verify

```bash
O365_TENANT_ID="$TENANT_ID" "$PYTHON" "$SKILL_DIR/scripts/setup.py" --check-live
# → LIVE_CHECK_OK: Authenticated as <Display Name> (<email>)

O365_TENANT_ID="$TENANT_ID" "$PYTHON" "$SKILL_DIR/scripts/o365_api.py" outlook folders | head
```

Anything other than `LIVE_CHECK_OK` plus a real email address means the
wrong account signed in — revoke and retry (Step 9).

## Step 9 — Revoke / start over

Local credentials only:

```bash
"$PYTHON" "$SKILL_DIR/scripts/setup.py" --revoke
```

Also revoke the app's grants in the tenant (so an old refresh token stops
working remotely):

- Personal accounts: <https://account.live.com/consent/Manage>
- Work/school accounts: <https://myapps.microsoft.com> → find the app → remove

Delete the app entirely:

```bash
az ad app delete --id "$APP_ID"
```

## Troubleshooting

| Symptom (from setup.py / o365_api.py) | Cause | Fix |
|---|---|---|
| `ValueError: You cannot use any scope value that is reserved` | `SCOPES` includes `offline_access`, `openid`, or `profile` | Remove the reserved scope from `SCOPES`; MSAL injects them automatically |
| `AADSTS50059: No tenant-identifying information found` | Single-tenant app + `/common` authority | Set `O365_TENANT_ID=<tenant>` in `~/.hermes/.env` |
| `AADSTS65001: The user or administrator has not consented` | Admin consent skipped (Step 5) and tenant restricts user consent | Run Step 5, or ask a tenant admin to grant it |
| `AADSTS700016: was not found in the directory` | Signed in as a user from a different tenant than where the app lives, with a single-tenant app | Either change `signInAudience` to `AzureADMultipleOrgs`, or re-register in the user's home tenant |
| `LIVE_CHECK_OK: ... (None)` | Authenticated as an account without a mailbox (e.g. an admin shadow account) | Revoke (Step 9), then re-auth as the regular user with the mailbox |
| Background `setup.py --auth` produces empty log file | Python stdout buffering when not a TTY | Use `PYTHONUNBUFFERED=1` and `python3 -u` |
| `HTTP 403` from a Graph call after auth succeeded | Missing one of the required delegated permissions, or scope dropped from manifest | Re-add the permission in Step 2's manifest, re-create the app or update via `az ad app update --required-resource-accesses`, re-run Step 5 |

## Security notes

- Treat `~/.hermes/o365_token_cache.bin` as a password. It is `chmod 0600`
  and holds a long-lived refresh token. Anyone who copies that file off
  the host can act as the signed-in user until the token is revoked.
- The Application (client) ID is **not** a secret. Logs and code commits
  can include it. The tenant ID is similarly not a secret.
- Pin to a **non-privileged** signed-in user. Do not authorize with Global
  Admin / Privileged Role Admin / any account in a PIM-eligible role.
- Re-run `setup.py --revoke` whenever you change hosts, suspect leakage,
  or rotate the user's password.
- For work tenants, monitor sign-in logs at:
  *Entra ID → Enterprise applications → \<app name\> → Sign-in logs*.
