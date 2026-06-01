# Microsoft 365 Skill — Setup Guide

A Hermes skill for Outlook, Calendar, Teams, and OneDrive via the Microsoft
Graph API. Authentication uses the OAuth 2.0 **device code flow** through a
public Azure AD application — **no client secret**, no redirect URI, no web
server.

This README is the human-facing setup guide. For the LLM-facing usage
contract (commands, output formats, rules), see [`SKILL.md`](SKILL.md).
For a fully scripted `az`-CLI walkthrough, see
[`AZURE-CLI-SETUP.md`](AZURE-CLI-SETUP.md).

---

## Quickstart (≈5 minutes)

The fastest path from a fresh checkout to a working skill on a local
Hermes install.

```bash
# 0. Get the skill
git clone https://github.com/kargausa/o365-skill.git
cd o365-skill

# 0b. Pick the Entra ID tenant your mailbox lives in
TENANT_ID="<your-tenant-guid>"          # az account list to find it
APP_NAME="${USER}-m365-cli"             # any neutral name; avoid product names

# 1. Drop the skill into Hermes (symlink keeps repo edits live)
ln -s "$PWD" ~/.hermes/skills/productivity/o365

# 2. Install Python deps in BOTH places Hermes might use:
#    a) the Hermes venv (gateway itself)
#    b) the host's python3 (subprocesses spawned for shell tools)
VIRTUAL_ENV=~/.hermes/hermes-agent/venv uv pip install msal requests
python3 -m pip install --user msal requests

# 3. Register the Azure AD app via az CLI (see AZURE-CLI-SETUP.md for
#    the full breakdown; this is the condensed form)
az ad sp show --id 00000003-0000-0000-c000-000000000000 \
  --query "oauth2PermissionScopes[?value=='User.Read' || value=='Mail.Read' \
            || value=='Mail.Send' || value=='Mail.ReadWrite' \
            || value=='Calendars.ReadWrite' || value=='Files.ReadWrite' \
            || value=='Chat.Read' || value=='ChatMessage.Send' \
            || value=='ChannelMessage.Send' \
            || value=='Team.ReadBasic.All' || value=='Channel.ReadBasic.All' \
            || value=='ChannelMessage.Read.All'] \
            .{name:value, id:id}" -o json   # confirm permission IDs

# (See AZURE-CLI-SETUP.md Steps 2–5 — they emit APP_ID + grant consent.)
APP_ID=<paste-from-step-3>

# 4. Persist tenant for Hermes (loaded into the gateway env on startup)
printf '\n# Microsoft 365 (o365 skill)\nO365_TENANT_ID=%s\n' "$TENANT_ID" \
  >> ~/.hermes/.env

# 5. Authenticate as a NON-PRIVILEGED user whose mailbox you want to use
O365_TENANT_ID="$TENANT_ID" PYTHONUNBUFFERED=1 \
  python3 -u ~/.hermes/skills/productivity/o365/scripts/setup.py --auth "$APP_ID"
# → open the printed URL, enter the code, sign in

# 6. Verify
O365_TENANT_ID="$TENANT_ID" python3 \
  ~/.hermes/skills/productivity/o365/scripts/setup.py --check-live
# → LIVE_CHECK_OK: Authenticated as <Name> (<email>)

# 7. Restart the gateway so the skill enters the system prompt
rm -f ~/.hermes/.skills_prompt_snapshot.json   # drop cached skill list
hermes gateway restart
```

That's it. Open Hermes and try *"search my Outlook for unread messages
from this week"*.

The longer sections below are reference material: portal-based
registration if you can't use `az`, individual step explanations, and
troubleshooting.

---

## What you need

1. A **Microsoft account** — any of:
   - Personal (`@outlook.com`, `@hotmail.com`, `@live.com`)
   - School or work (Microsoft 365 / Entra ID)
2. About **5 minutes** to register an Azure AD application (free, one-time).
3. The host running Hermes must reach `login.microsoftonline.com` and
   `graph.microsoft.com` outbound on port 443.

You do **not** need:

- A paid Azure subscription. App registration is free.
- A client secret. Device code flow uses a public client.
- Admin consent (for personal accounts). Work/school accounts may need an
  admin to grant the listed permissions — talk to your IT admin if Azure
  shows "admin consent required".

---

## Step 1 — Register an Azure AD application

This produces the **Application (client) ID** that the skill needs.

1. Sign in to the Azure Portal:
   <https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApplications/ApplicationsListBlade>
2. Click **New registration**.
3. Fill in:
   - **Name**: `Hermes Agent` (or anything you like — only you see this)
   - **Supported account types**: *Accounts in any organizational directory
     and personal Microsoft accounts* (most permissive — pick a narrower one
     if your IT requires it)
   - **Redirect URI**: leave blank
4. Click **Register**.
5. On the new app's **Overview** page, copy the **Application (client) ID**
   (a UUID like `12345678-90ab-cdef-1234-567890abcdef`). You will pass this
   to `setup.py`.

### Step 1a — Enable public client flows

Device code flow requires the app to be marked as a public client.

1. In the app, go to **Authentication** (left sidebar).
2. Scroll down to **Advanced settings** → **Allow public client flows**.
3. Toggle to **Yes** → **Save**.

### Step 1b — Grant Graph API permissions

The skill needs these **delegated** Microsoft Graph permissions:

| Permission | Why |
|------------|-----|
| `User.Read` | Identify the signed-in user (for `--check-live`) |
| `Mail.Read` | Read Outlook messages |
| `Mail.Send` | Send email |
| `Mail.ReadWrite` | Reply to messages, mark read/unread |
| `Calendars.ReadWrite` | List, create, delete calendar events |
| `Files.ReadWrite` | OneDrive search, upload, download, share |
| `Chat.Read` | List Teams personal/group chats and read messages |
| `ChatMessage.Send` | Reply to Teams chats (admin consent required) |
| `ChannelMessage.Send` | Post to Teams channels |
| `Team.ReadBasic.All` | List joined Teams (admin consent required) |
| `Channel.ReadBasic.All` | List channels within a team (admin consent required) |
| `ChannelMessage.Read.All` | Read messages from Teams channels (admin consent required) |
| `offline_access` | Refresh tokens (so you don't re-auth every hour) |

To add them:

1. In the app, go to **API permissions** → **Add a permission**.
2. **Microsoft Graph** → **Delegated permissions**.
3. Search and check each permission in the table above → **Add permissions**.
4. (Work/school accounts only) Click **Grant admin consent for &lt;tenant&gt;**.
   Personal accounts grant consent at first sign-in instead.

> **Drop permissions you don't need.** If you only want Outlook search, just
> grant `User.Read`, `Mail.Read`, and `offline_access`. The skill will still
> load — you'll just get HTTP 403 from operations whose permission is
> missing, and the troubleshooting table in `SKILL.md` covers that.

---

## Step 2 — Install dependencies

The skill needs `msal` and `requests`. Run:

```bash
python3 ${HERMES_HOME:-$HOME/.hermes}/skills/productivity/o365/scripts/setup.py --install-deps
```

This pip-installs into whatever Python `setup.py` is invoked with (i.e.
into the Hermes runtime environment). If you manage Python with `uv` or
a venv, make sure `setup.py` runs under the same interpreter the agent
will use.

> **For Hermes specifically**, install into the gateway venv AND into
> the host's `python3` user site, because Hermes runs the gateway from
> its venv but spawns shell-tool subprocesses that may resolve `python3`
> via PATH. Both must have `msal`:
>
> ```bash
> VIRTUAL_ENV=~/.hermes/hermes-agent/venv uv pip install msal requests
> python3 -m pip install --user msal requests
> ```

---

## Step 3 — Authenticate (device code flow)

Run, passing the client ID from Step 1:

```bash
# For single-tenant apps you MUST set O365_TENANT_ID. Multi-tenant /
# personal-account apps can omit it (authority defaults to /common).
O365_TENANT_ID=<your-tenant-guid> \
  python3 ${HERMES_HOME:-$HOME/.hermes}/skills/productivity/o365/scripts/setup.py \
    --auth 12345678-90ab-cdef-1234-567890abcdef
```

**Authorize as a non-privileged user** — the user whose mailbox you
actually want to read/send from. The refresh token in the resulting
cache is long-lived; pinning it to a Global Admin or any PIM-eligible
role makes a stolen token cache equal to a tenant compromise.

The script prints:

```
============================================================
DEVICE CODE AUTHENTICATION
============================================================
URL:  https://microsoft.com/devicelogin
Code: ABC123DEF
...
Waiting for authorization...
```

On any device (your phone is fine):

1. Open the URL.
2. Enter the code shown.
3. Sign in with the Microsoft account you want the skill to act as.
4. Approve the requested permissions.

Back in the terminal, the script polls every few seconds and prints
`OK: Authenticated as <name>` once you finish. The device code is valid for
about **15 minutes** — if it expires, just rerun `--auth`.

If the client ID was already saved during an earlier run, you can omit it:

```bash
python3 ${HERMES_HOME:-$HOME/.hermes}/skills/productivity/o365/scripts/setup.py --auth
```

---

## Step 4 — Verify

```bash
O365_TENANT_ID=<your-tenant-guid> \
  python3 ${HERMES_HOME:-$HOME/.hermes}/skills/productivity/o365/scripts/setup.py --check
```

Prints `AUTHENTICATED: Token valid for <email>` and exits 0 if the cache is
healthy. For a stronger check that actually calls Graph:

```bash
O365_TENANT_ID=<your-tenant-guid> \
  python3 ${HERMES_HOME:-$HOME/.hermes}/skills/productivity/o365/scripts/setup.py --check-live
```

A healthy result looks like `LIVE_CHECK_OK: Authenticated as Jane Doe
(jane@example.com)`. If the email shown is `(None)`, the signed-in user
has no mailbox (typical of admin / service accounts) — revoke and re-auth
as a real user (see Step 9 below).

Tokens refresh automatically from now on — you should never need to
re-auth unless you revoke or the refresh token is invalidated by
Microsoft (e.g. password change, conditional access policy).

---

## Step 5 — Wire it into Hermes

The skill must live where Hermes scans for skills, and the gateway must
be told to reload its skill index.

```bash
# Symlink (recommended while iterating) or copy into the productivity
# category — alongside google-workspace, notion, linear, etc.
ln -s "$PWD/skills/o365" ~/.hermes/skills/productivity/o365

# Pin the tenant so the gateway env carries it to every shell-tool
# subprocess that runs the o365 scripts.
printf '\n# Microsoft 365 (o365 skill)\nO365_TENANT_ID=<your-tenant-guid>\n' \
  >> ~/.hermes/.env

# Drop the cached skill-list snapshot so the prompt rebuilds with o365
# in it; then restart so the in-process cache is also dropped.
rm -f ~/.hermes/.skills_prompt_snapshot.json
hermes gateway restart
```

`hermes skills list | grep o365` should now show the skill as `enabled`.
Open Hermes and the model will mention the skill on relevant requests.

If `python` (no `3`) is what the model emits in shell commands and your
host lacks it (macOS only ships `python3`), add a tiny shim somewhere on
`$PATH` — e.g. `~/.local/bin/python`:

```bash
cat > ~/.local/bin/python <<'EOF'
#!/usr/bin/env bash
exec /usr/bin/python3 "$@"
EOF
chmod +x ~/.local/bin/python
```

---

## Where files land

Everything lives under `$HERMES_HOME` (default `~/.hermes`):

| Path | What it is | Permissions |
|------|------------|-------------|
| `~/.hermes/o365_app_config.json` | Your Azure AD client ID | `0600` |
| `~/.hermes/o365_token_cache.bin` | MSAL token cache (refresh + access tokens) | `0600` |

Both files are written atomically (`.tmp` + `os.replace`) and chmodded to
`0600`. **Treat `o365_token_cache.bin` like a password** — anyone with read
access to it can act as you against Microsoft 365 until you revoke.

---

## Revoking access

To remove local credentials only:

```bash
python3 ${HERMES_HOME:-$HOME/.hermes}/skills/productivity/o365/scripts/setup.py --revoke
```

This deletes the token cache and app config from disk. To also revoke the
app's access from your Microsoft account so the refresh token stops working:

- Personal accounts: <https://account.live.com/consent/Manage>
- Work/school accounts: <https://myapps.microsoft.com> → find "Hermes Agent"
  → remove

If you suspect the token cache leaked, do both — local delete + remote
revoke.

---

## Troubleshooting

| Symptom | Cause / Fix |
|---------|-------------|
| `NOT_AUTHENTICATED: No app config` | First-time setup — pass the client ID: `setup.py --auth CLIENT_ID` |
| `NOT_AUTHENTICATED: No token cache` | Run `setup.py --auth` to complete device flow |
| `TOKEN_INVALID: ...` | Refresh token revoked or expired. `setup.py --revoke` then re-auth |
| `ERROR: Failed to initiate device flow` | "Allow public client flows" not enabled in Azure (Step 1a) |
| `AADSTS50059: No tenant-identifying information found` | Single-tenant app + default `/common` authority. Set `O365_TENANT_ID=<tenant-guid>` in `~/.hermes/.env` (and the current shell for any direct CLI runs) |
| `AADSTS65001: ... has not consented` | Tenant restricts user consent and admin consent wasn't granted. Have a tenant admin run [`AZURE-CLI-SETUP.md`](AZURE-CLI-SETUP.md) Step 5 (`oauth2PermissionGrants` POST) |
| `AADSTS70011: Invalid scope` | A requested Graph permission is missing in Azure (Step 1b) |
| `AADSTS700016: was not found in the directory` | Signed in with a user from a different tenant than where the app is registered (single-tenant app). Re-register as `AzureADMultipleOrgs`, or move the app to the user's home tenant |
| `ValueError: You cannot use any scope value that is reserved` | `SCOPES` includes one of `offline_access`, `openid`, `profile`. Remove it — MSAL adds these implicitly when needed |
| `LIVE_CHECK_OK: ... (None)` | Authenticated as an account with no mailbox (e.g. admin / shadow account). Revoke and re-auth as the real user |
| `HTTP 403` from any command | The matching delegated permission isn't granted — add it in Azure → API permissions and re-run admin consent |
| `HTTP 401` after working previously | Token expired and silent refresh failed. Re-run `setup.py --auth` |
| `ModuleNotFoundError: msal` | Install into the right Python. For Hermes: `VIRTUAL_ENV=~/.hermes/hermes-agent/venv uv pip install msal requests` **and** `python3 -m pip install --user msal requests` |
| `python: command not found` (macOS) when Hermes runs the command | Add a `python` → `python3` shim at `~/.local/bin/python` (see Step 5) |
| Background `setup.py --auth` produces empty log file | Python stdout buffering. Run with `PYTHONUNBUFFERED=1` and `python3 -u` |
| Setup hangs on "Waiting for authorization..." | You haven't entered the code yet. Open the printed URL, paste the code, sign in. Times out after ~15 minutes |
| Work account: "admin approval required" | An IT admin must grant consent for the listed Graph permissions for your tenant |
| Hermes model says "I don't have access to the o365 skill" | Skill not in the gateway's prompt cache. `rm ~/.hermes/.skills_prompt_snapshot.json && hermes gateway restart` |

---

## Security notes

- **No client secret means no shared secret to rotate** — your only secret
  is the refresh token in `o365_token_cache.bin`. Keep that file at `0600`
  on a single trusted host.
- **The client ID is not secret.** It's safe to commit to a repo or share
  in logs. (You still shouldn't, since it makes phishing slightly easier,
  but a leak is not a credential leak.)
- **Tokens are bound to your Microsoft account, not to the host.** Anyone
  who copies `o365_token_cache.bin` off the host can use it from anywhere
  until you revoke.
- If you run multiple Hermes instances against the same `$HERMES_HOME`,
  they share one token cache. Cache writes are atomic, but there is no
  cross-process lock — concurrent `--auth` runs could race. Don't.
