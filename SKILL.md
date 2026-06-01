---
name: o365
description: "Outlook, Calendar, Teams, OneDrive via Microsoft Graph API."
version: 1.0.0
author: kargausa
license: MIT
platforms: [linux, macos, windows]
required_credential_files:
  - path: o365_token_cache.bin
    description: MSAL token cache (created by setup script)
  - path: o365_app_config.json
    description: Azure AD app client ID (created during setup)
metadata:
  hermes:
    tags: [Microsoft, Outlook, Calendar, Teams, OneDrive, Email, OAuth, O365]
    related_skills: [google-workspace]
---

# Microsoft 365

Outlook, Calendar, Teams, and OneDrive — through device code OAuth and the Microsoft Graph API. No client secret required (public client / device code flow), ideal for headless environments.

## References

- `references/outlook-search-syntax.md` — Outlook search operators ($search KQL syntax, $filter OData syntax)

## Scripts

- `scripts/setup.py` — OAuth2 device code setup (run once to authorize)
- `scripts/o365_api.py` — Graph API CLI for all M365 services

## First-Time Setup

The setup is fully non-interactive — you drive it step by step so it works
on CLI, Telegram, Discord, or any platform.

Define a shorthand first:

```bash
O365SETUP="python3 ${HERMES_HOME:-$HOME/.hermes}/skills/productivity/o365/scripts/setup.py"
```

### Step 0: Check if already set up

```bash
$O365SETUP --check
```

If it prints `AUTHENTICATED`, skip to Usage — setup is already done.

### Step 1: Register an Azure AD application (one-time, ~5 minutes)

Tell the user:

> You need an Azure AD app registration. This is a one-time setup:
>
> 1. Go to Azure Portal → App registrations:
>    https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApplications/ApplicationsListBlade
> 2. Click **New registration**
> 3. Name: "Hermes Agent" (or any name)
> 4. Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**
> 5. Redirect URI: Leave blank (not needed for device code flow)
> 6. Click **Register**
> 7. Copy the **Application (client) ID** from the Overview page
> 8. Go to **Authentication** → Under **Advanced settings**, set **Allow public client flows** to **Yes** → Save
> 9. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
>    Add: `User.Read`, `Mail.Read`, `Mail.Send`, `Mail.ReadWrite`, `Calendars.ReadWrite`,
>    `Files.ReadWrite`, `Chat.Read`, `ChatMessage.Send`, `ChannelMessage.Send`,
>    `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Read.All`,
>    `offline_access`.
>    `ChatMessage.Send` and the three Teams-channel reads (`Team.ReadBasic.All`,
>    `Channel.ReadBasic.All`, `ChannelMessage.Read.All`) require **admin consent**
>    — click **Grant admin consent for &lt;tenant&gt;** after adding them, or have
>    a tenant admin do it.
> 10. Tell me the Application (client) ID

### Step 2: Authenticate with device code

```bash
$O365SETUP --auth CLIENT_ID_HERE
```

Or if the client ID was already saved:

```bash
$O365SETUP --auth
```

The script will print a URL and a code. Tell the user:

> Open this URL and enter the code shown:
> URL: https://microsoft.com/devicelogin
> Code: XXXXXXXXX

The script polls automatically and completes when the user authorizes. The
device code is valid for ~15 minutes — the script will block (no progress
output) until the user enters the code or the code expires.

### Step 3: Verify

```bash
$O365SETUP --check
```

Should print `AUTHENTICATED`. Setup is complete — tokens refresh automatically.

### Notes

- Token cache is stored at `~/.hermes/o365_token_cache.bin` and auto-refreshes via MSAL.
- App config (just the client ID) is at `~/.hermes/o365_app_config.json`.
- No client secret is needed — device code flow uses public clients.
- To revoke: `$O365SETUP --revoke`

## Usage

All commands go through the API script. Set `O365` as a shorthand:

```bash
O365="python3 ${HERMES_HOME:-$HOME/.hermes}/skills/productivity/o365/scripts/o365_api.py"
```

### Outlook

```bash
# Search (returns JSON array with id, subject, from, date, snippet)
$O365 outlook search "from:boss@company.com subject:urgent" --max 10
$O365 outlook search "hasAttachments:true received>=2026-05-01"

# Search with $filter instead of $search (for structured queries)
$O365 outlook search "isRead eq false and importance eq 'high'" --filter

# Read full message (returns JSON with body text — body is truncated to
# 50,000 chars by default; pass --max-body 0 for unlimited)
$O365 outlook get MESSAGE_ID
$O365 outlook get MESSAGE_ID --max-body 5000   # short preview
$O365 outlook get MESSAGE_ID --max-body 0      # no truncation

# Send
$O365 outlook send --to user@example.com --subject "Hello" --body "Message text"
$O365 outlook send --to "a@example.com,b@example.com" --cc c@example.com --subject "Report" --body "<h1>Q4</h1>" --html

# Reply (Reply All by default, so existing recipients stay included)
$O365 outlook reply MESSAGE_ID --body "Thanks, that works for me."

# List mail folders
$O365 outlook folders
```

### Calendar

```bash
# List events (defaults to next 7 days)
$O365 calendar list
$O365 calendar list --start 2026-06-01T00:00:00 --end 2026-06-07T23:59:59

# Create event (ISO 8601, specify timezone)
$O365 calendar create --summary "Team Standup" --start 2026-06-01T10:00:00 --end 2026-06-01T10:30:00 --timezone "America/Chicago"
$O365 calendar create --summary "Lunch" --start 2026-06-01T12:00:00 --end 2026-06-01T13:00:00 --location "Cafe" --attendees "alice@co.com,bob@co.com"

# Delete event
$O365 calendar delete EVENT_ID
```

### Teams

```bash
# List joined teams
$O365 teams list

# List channels in a team
$O365 teams channels TEAM_ID

# Get channel messages
$O365 teams messages TEAM_ID CHANNEL_ID --max 20

# Send a channel message
$O365 teams send TEAM_ID CHANNEL_ID --body "Hello team!"
$O365 teams send TEAM_ID CHANNEL_ID --body "<b>Important</b> update" --html

# List chats
$O365 teams chats --max 20

# Get chat messages
$O365 teams chat-messages CHAT_ID --max 20

# Send a chat message (1:1 or group)
$O365 teams chat-send CHAT_ID --body "Thanks, will follow up shortly."
$O365 teams chat-send CHAT_ID --body "<b>Heads up</b>: meeting moved." --html
```

### OneDrive

```bash
# Search files
$O365 onedrive search "quarterly report" --max 10

# Get file metadata
$O365 onedrive get ITEM_ID

# Upload a local file (<4MB simple upload)
$O365 onedrive upload /path/to/report.pdf
$O365 onedrive upload /path/to/image.png --name "Logo.png" --parent-id FOLDER_ID

# Download a file
$O365 onedrive download ITEM_ID
$O365 onedrive download ITEM_ID --output ~/Downloads/report.pdf

# Create a folder
$O365 onedrive create-folder "Reports"
$O365 onedrive create-folder "Q4" --parent-id FOLDER_ID

# Share (creates an anonymous link)
$O365 onedrive share ITEM_ID --type view
$O365 onedrive share ITEM_ID --type edit
```

## Output Format

All commands return JSON. Parse with `jq` or read directly. Key fields:

- **Outlook search**: `[{id, subject, from, to, date, isRead, snippet, hasAttachments}]`
- **Outlook get**: `{id, conversationId, subject, from, to, cc, date, isRead, hasAttachments, bodyType, bodyLength, bodyTruncated, body}`
- **Outlook send**: `{status: "sent", to, subject}`
- **Outlook reply**: `{status: "replied", messageId}`
- **Outlook folders**: `[{id, name, unreadCount, totalCount}]`
- **Calendar list**: `[{id, summary, start, startTz, end, endTz, location, organizer, isAllDay, snippet, webLink}]`
- **Calendar create**: `{status: "created", id, summary, webLink}`
- **Calendar delete**: `{status: "deleted", eventId}`
- **Teams list**: `[{id, name, description}]`
- **Teams channels**: `[{id, name, description}]`
- **Teams messages**: `[{id, from, date, body, bodyType}]`
- **Teams send**: `{status: "sent", id, teamId, channelId}`
- **Teams chats**: `[{id, topic, type, lastUpdated}]`
- **Teams chat messages**: `[{id, from, date, body, bodyType}]`
- **OneDrive search**: `[{id, name, size, mimeType, lastModified, webUrl}]`
- **OneDrive get**: `{id, name, size, mimeType, lastModified, webUrl, createdBy, parentPath}`
- **OneDrive upload**: `{status: "uploaded", id, name, size, webUrl}`
- **OneDrive download**: `{status: "downloaded", id, name, path, size}`
- **OneDrive create-folder**: `{status: "created", id, name, webUrl}`
- **OneDrive share**: `{status: "shared", itemId, type, webUrl}`

## Rules

1. **Never send email, create/delete calendar events, send Teams messages, upload/delete/share OneDrive files without confirming with the user first.** Show what will be done (recipients, content, file IDs) and ask for approval.
2. **Check auth before first use** — run `setup.py --check`. If it fails, guide the user through setup.
3. **Use the Outlook search syntax reference** for complex queries — load it with `skill_view("o365", file_path="references/outlook-search-syntax.md")`.
4. **Calendar times**: Use ISO 8601 and specify `--timezone` (e.g., `America/New_York`, `UTC`).
5. **Respect rate limits** — avoid rapid-fire sequential API calls. Batch reads when possible.
6. **OneDrive upload limit**: Simple upload supports files up to 4MB. For larger files, instruct the user to use the OneDrive web UI.

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `NOT_AUTHENTICATED` | Run setup Steps 1-3 above |
| `TOKEN_INVALID` | Token expired or revoked — `$O365SETUP --revoke` then redo Steps 2-3 |
| `HTTP 403` | Missing API permission — add the permission in Azure Portal → API permissions |
| `HTTP 401` | Token expired — usually auto-refreshes; if persistent, `$O365SETUP --revoke` and re-auth |
| `ModuleNotFoundError` | Run `$O365SETUP --install-deps` |
| Device code flow fails | Ensure "Allow public client flows" is set to Yes in Azure Portal → Authentication |
| `AADSTS70011: Invalid scope` | The requested scope is not configured — add it in Azure Portal → API permissions |

## Revoking Access

```bash
$O365SETUP --revoke
```

This deletes the local token cache and app config. To also revoke the app's access from the Microsoft account, visit: https://account.live.com/consent/Manage
