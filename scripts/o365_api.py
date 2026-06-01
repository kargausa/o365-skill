#!/usr/bin/env python3
"""Microsoft 365 Graph API CLI for Hermes Agent.

All operations use the Microsoft Graph REST API via ``requests`` with an
MSAL-managed access token.  Output is always JSON to stdout.

Usage:
  python o365_api.py outlook search "from:alice subject:report" [--max 10]
  python o365_api.py outlook get MESSAGE_ID
  python o365_api.py outlook send --to user@example.com --subject "Hi" --body "Hello"
  python o365_api.py outlook reply MESSAGE_ID --body "Thanks"
  python o365_api.py outlook folders
  python o365_api.py calendar list [--start DATE] [--end DATE]
  python o365_api.py calendar create --summary "Meeting" --start DATETIME --end DATETIME
  python o365_api.py calendar delete EVENT_ID
  python o365_api.py teams list
  python o365_api.py teams channels TEAM_ID
  python o365_api.py teams messages TEAM_ID CHANNEL_ID [--max 20]
  python o365_api.py teams send TEAM_ID CHANNEL_ID --body "Hello team"
  python o365_api.py teams chats [--max 20]
  python o365_api.py teams chat-messages CHAT_ID [--max 20]
  python o365_api.py onedrive search "quarterly report" [--max 10]
  python o365_api.py onedrive get ITEM_ID
  python o365_api.py onedrive upload LOCAL_PATH [--name NAME] [--parent-id PARENT_ID]
  python o365_api.py onedrive download ITEM_ID [--output PATH]
  python o365_api.py onedrive create-folder NAME [--parent-id PARENT_ID]
  python o365_api.py onedrive share ITEM_ID [--type view|edit]
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import urllib.parse
from datetime import datetime, timedelta, timezone
from pathlib import Path

# Ensure sibling modules (_hermes_home) are importable when run standalone.
_SCRIPTS_DIR = str(Path(__file__).resolve().parent)
if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

from _hermes_home import get_hermes_home

HERMES_HOME = get_hermes_home()
TOKEN_CACHE_PATH = HERMES_HOME / "o365_token_cache.bin"
APP_CONFIG_PATH = HERMES_HOME / "o365_app_config.json"

AUTHORITY = "https://login.microsoftonline.com/" + os.environ.get("O365_TENANT_ID", "common")
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

SCOPES = [
    "User.Read",
    "Mail.Read",
    "Mail.Send",
    "Mail.ReadWrite",
    "Calendars.ReadWrite",
    "Files.ReadWrite",
    "Chat.Read",
    "ChatMessage.Send",
    "ChannelMessage.Send",
    "Team.ReadBasic.All",
    "Channel.ReadBasic.All",
    "ChannelMessage.Read.All",
    # MSAL Python treats offline_access as a reserved scope in newer releases;
    # refresh tokens are still issued for public-client/device-code flows without
    # explicitly requesting it here.
]


# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

def _load_app_config() -> dict:
    """Load app config or exit."""
    if not APP_CONFIG_PATH.exists():
        print(json.dumps({"error": "Not authenticated. Run setup.py --auth first."}))
        sys.exit(1)
    try:
        data = json.loads(APP_CONFIG_PATH.read_text())
        if data.get("client_id"):
            return data
    except Exception:
        pass
    print(json.dumps({"error": f"Invalid app config at {APP_CONFIG_PATH}"}))
    sys.exit(1)


def _get_token() -> str:
    """Acquire an access token from the MSAL cache (silent / refresh)."""
    import msal

    config = _load_app_config()
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        cache.deserialize(TOKEN_CACHE_PATH.read_text())

    app = msal.PublicClientApplication(
        config["client_id"],
        authority=AUTHORITY,
        token_cache=cache,
    )

    accounts = app.get_accounts()
    if not accounts:
        print(json.dumps({"error": "No accounts in token cache. Run setup.py --auth."}))
        sys.exit(1)

    result = app.acquire_token_silent(SCOPES, account=accounts[0])

    # Persist refreshed cache (atomic write + 0600 perms, so a crash mid-write
    # can't corrupt the cache and concurrent processes don't trample each other).
    if cache.has_state_changed:
        TOKEN_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
        tmp = TOKEN_CACHE_PATH.with_suffix(TOKEN_CACHE_PATH.suffix + ".tmp")
        tmp.write_text(cache.serialize())
        os.chmod(tmp, 0o600)
        os.replace(tmp, TOKEN_CACHE_PATH)

    if result and "access_token" in result:
        return result["access_token"]

    error = result.get("error_description", "token acquisition failed") if result else "token acquisition failed"
    print(json.dumps({"error": f"Token error: {error}"}))
    sys.exit(1)


# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------

def _headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


def _graph_get(path: str, token: str, *, params: dict | None = None) -> dict:
    """GET from Graph API. Returns parsed JSON."""
    import requests
    url = f"{GRAPH_BASE}{path}"
    resp = requests.get(url, headers=_headers(token), params=params, timeout=30)
    if resp.status_code >= 400:
        return {"error": f"HTTP {resp.status_code}", "detail": resp.text[:500]}
    if not resp.content:
        return {}
    return resp.json()


def _graph_get_binary(path: str, token: str) -> bytes:
    """GET binary content from Graph API."""
    import requests
    url = f"{GRAPH_BASE}{path}"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    if resp.status_code >= 400:
        print(json.dumps({"error": f"HTTP {resp.status_code}", "detail": resp.text[:500]}))
        sys.exit(1)
    return resp.content


def _graph_post(path: str, token: str, body: dict | None = None) -> dict:
    """POST to Graph API. Returns parsed JSON."""
    import requests
    url = f"{GRAPH_BASE}{path}"
    resp = requests.post(url, headers=_headers(token), json=body, timeout=30)
    if resp.status_code >= 400:
        return {"error": f"HTTP {resp.status_code}", "detail": resp.text[:500]}
    if not resp.content:
        return {"status": "ok"}
    return resp.json()


def _graph_put(path: str, token: str, data: bytes, content_type: str = "application/octet-stream") -> dict:
    """PUT binary content to Graph API."""
    import requests
    url = f"{GRAPH_BASE}{path}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": content_type,
    }
    resp = requests.put(url, headers=headers, data=data, timeout=60)
    if resp.status_code >= 400:
        return {"error": f"HTTP {resp.status_code}", "detail": resp.text[:500]}
    return resp.json()


def _graph_patch(path: str, token: str, body: dict) -> dict:
    """PATCH to Graph API."""
    import requests
    url = f"{GRAPH_BASE}{path}"
    resp = requests.patch(url, headers=_headers(token), json=body, timeout=30)
    if resp.status_code >= 400:
        return {"error": f"HTTP {resp.status_code}", "detail": resp.text[:500]}
    if not resp.content:
        return {"status": "ok"}
    return resp.json()


def _graph_delete(path: str, token: str) -> dict:
    """DELETE via Graph API."""
    import requests
    url = f"{GRAPH_BASE}{path}"
    resp = requests.delete(url, headers=_headers(token), timeout=30)
    if resp.status_code == 204:
        return {"status": "deleted"}
    if resp.status_code >= 400:
        return {"error": f"HTTP {resp.status_code}", "detail": resp.text[:500]}
    return {"status": "deleted"}


def _graph_list(path: str, token: str, *, params: dict | None = None, max_items: int = 0) -> dict:
    """GET a list endpoint and follow @odata.nextLink until max_items is reached.

    Returns a dict with the same shape as a single Graph response (``{"value": [...]}``)
    or ``{"error": ..., "detail": ...}`` if the initial request fails. Errors
    encountered while following nextLink stop pagination but keep whatever was
    collected so far — Graph's per-page cap (often 50) means a single $top is
    not enough to satisfy --max for the user.
    """
    import requests
    initial = _graph_get(path, token, params=params)
    if "error" in initial:
        return initial
    items = list(initial.get("value", []))
    next_link = initial.get("@odata.nextLink")
    while next_link and (max_items <= 0 or len(items) < max_items):
        resp = requests.get(next_link, headers=_headers(token), timeout=30)
        if resp.status_code >= 400:
            break
        data = resp.json()
        items.extend(data.get("value", []))
        next_link = data.get("@odata.nextLink")
    if max_items > 0:
        items = items[:max_items]
    return {"value": items}


def _out(data):
    """Print JSON to stdout."""
    print(json.dumps(data, indent=2, default=str))


# ---------------------------------------------------------------------------
# Outlook
# ---------------------------------------------------------------------------

def _outlook_search(args, token: str):
    # Graph caps message page size at 25, so use the paginating helper to
    # actually deliver --max items rather than silently truncating.
    params = {"$top": "25"}
    query = args.query

    if args.filter:
        params["$filter"] = query
        # $orderby is valid alongside $filter.
        params["$orderby"] = "receivedDateTime desc"
    else:
        params["$search"] = f'"{query}"'
        # Graph rejects $search combined with $orderby — results come back
        # ranked by relevance from the search index.

    params["$select"] = "id,subject,from,toRecipients,receivedDateTime,isRead,bodyPreview,hasAttachments"

    base = "/me/messages"
    if getattr(args, "folder_id", None):
        base = f"/me/mailFolders/{args.folder_id}/messages"

    data = _graph_list(base, token, params=params, max_items=args.max)
    if "error" in data:
        _out(data)
        return

    messages = data.get("value", [])
    results = []
    for m in messages:
        from_addr = m.get("from", {}).get("emailAddress", {})
        to_addrs = [r.get("emailAddress", {}).get("address", "") for r in m.get("toRecipients", [])]
        results.append({
            "id": m.get("id"),
            "subject": m.get("subject"),
            "from": f"{from_addr.get('name', '')} <{from_addr.get('address', '')}>",
            "to": to_addrs,
            "date": m.get("receivedDateTime"),
            "isRead": m.get("isRead"),
            "snippet": m.get("bodyPreview", "")[:200],
            "hasAttachments": m.get("hasAttachments"),
        })
    _out(results)


def _outlook_get(args, token: str):
    params = {"$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,body,hasAttachments,conversationId"}
    data = _graph_get(f"/me/messages/{args.message_id}", token, params=params)
    if "error" in data:
        _out(data)
        return

    from_addr = data.get("from", {}).get("emailAddress", {})
    to_addrs = [r.get("emailAddress", {}).get("address", "") for r in data.get("toRecipients", [])]
    cc_addrs = [r.get("emailAddress", {}).get("address", "") for r in data.get("ccRecipients", [])]

    body_obj = data.get("body", {})
    body_text = body_obj.get("content", "")
    body_full_length = len(body_text)
    body_truncated = False
    if args.max_body and body_full_length > args.max_body:
        body_text = body_text[: args.max_body]
        body_truncated = True

    _out({
        "id": data.get("id"),
        "conversationId": data.get("conversationId"),
        "subject": data.get("subject"),
        "from": f"{from_addr.get('name', '')} <{from_addr.get('address', '')}>",
        "to": to_addrs,
        "cc": cc_addrs,
        "date": data.get("receivedDateTime"),
        "isRead": data.get("isRead"),
        "hasAttachments": data.get("hasAttachments"),
        "bodyType": body_obj.get("contentType"),
        "bodyLength": body_full_length,
        "bodyTruncated": body_truncated,
        "body": body_text,
    })


def _outlook_send(args, token: str):
    import base64, mimetypes, os
    to_recipients = [{"emailAddress": {"address": addr.strip()}} for addr in args.to.split(",")]
    cc_recipients = []
    if args.cc:
        cc_recipients = [{"emailAddress": {"address": addr.strip()}} for addr in args.cc.split(",")]

    content_type = "html" if args.html else "text"
    message = {
        "subject": args.subject,
        "body": {"contentType": content_type, "content": args.body},
        "toRecipients": to_recipients,
    }
    if cc_recipients:
        message["ccRecipients"] = cc_recipients

    attachments = getattr(args, "attachment", None) or []
    if attachments:
        message["attachments"] = []
        for path in attachments:
            mime = mimetypes.guess_type(path)[0] or "application/octet-stream"
            with open(path, "rb") as f:
                data = base64.b64encode(f.read()).decode()
            message["attachments"].append({
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": os.path.basename(path),
                "contentType": mime,
                "contentBytes": data,
            })

    result = _graph_post("/me/sendMail", token, {"message": message})
    if "error" in result:
        _out(result)
    else:
        _out({"status": "sent", "to": args.to, "subject": args.subject,
              "attachments": [os.path.basename(p) for p in attachments]})


def _outlook_reply(args, token: str):
    # Hasan's default workflow is Reply All so existing thread participants stay included.
    body = {"comment": args.body}
    result = _graph_post(f"/me/messages/{args.message_id}/replyAll", token, body)
    if "error" in result:
        _out(result)
    else:
        _out({"status": "replied_all", "messageId": args.message_id})


def _outlook_folders(args, token: str):
    data = _graph_get("/me/mailFolders", token, params={"$top": "50"})
    if "error" in data:
        _out(data)
        return
    folders = [
        {
            "id": f.get("id"),
            "name": f.get("displayName"),
            "unreadCount": f.get("unreadItemCount"),
            "totalCount": f.get("totalItemCount"),
        }
        for f in data.get("value", [])
    ]
    _out(folders)


def _outlook_move(args, token: str):
    """Move a message to a different mail folder."""
    body = {"destinationId": args.destination_id}
    result = _graph_post(f"/me/messages/{args.message_id}/move", token, body)
    if "error" in result:
        _out(result)
        return
    _out({
        "status": "moved",
        "messageId": args.message_id,
        "destinationId": args.destination_id,
        "newMessageId": result.get("id"),
    })


# ---------------------------------------------------------------------------
# Calendar
# ---------------------------------------------------------------------------

def _calendar_list(args, token: str):
    start = args.start or datetime.now(timezone.utc).isoformat()
    if args.end:
        end = args.end
    else:
        # Default: 7 days from start
        try:
            start_dt = datetime.fromisoformat(start.replace("Z", "+00:00"))
        except ValueError:
            start_dt = datetime.now(timezone.utc)
        end = (start_dt + timedelta(days=7)).isoformat()

    params = {
        "startDateTime": start,
        "endDateTime": end,
        "$select": "id,subject,start,end,location,bodyPreview,organizer,attendees,isAllDay,webLink",
        "$orderby": "start/dateTime",
        "$top": "50",
    }

    data = _graph_list("/me/calendarView", token, params=params, max_items=args.max)
    if "error" in data:
        _out(data)
        return

    events = []
    for e in data.get("value", []):
        location = e.get("location", {})
        organizer = e.get("organizer", {}).get("emailAddress", {})
        events.append({
            "id": e.get("id"),
            "summary": e.get("subject"),
            "start": e.get("start", {}).get("dateTime"),
            "startTz": e.get("start", {}).get("timeZone"),
            "end": e.get("end", {}).get("dateTime"),
            "endTz": e.get("end", {}).get("timeZone"),
            "location": location.get("displayName", ""),
            "organizer": f"{organizer.get('name', '')} <{organizer.get('address', '')}>",
            "isAllDay": e.get("isAllDay"),
            "snippet": e.get("bodyPreview", "")[:200],
            "webLink": e.get("webLink"),
        })
    _out(events)


def _calendar_create(args, token: str):
    event = {
        "subject": args.summary,
        "start": {"dateTime": args.start, "timeZone": args.timezone},
        "end": {"dateTime": args.end, "timeZone": args.timezone},
    }
    if args.location:
        event["location"] = {"displayName": args.location}
    if args.body:
        event["body"] = {"contentType": "text", "content": args.body}
    if args.attendees:
        event["attendees"] = [
            {"emailAddress": {"address": addr.strip()}, "type": "required"}
            for addr in args.attendees.split(",")
        ]

    result = _graph_post("/me/events", token, event)
    if "error" in result:
        _out(result)
    else:
        _out({
            "status": "created",
            "id": result.get("id"),
            "summary": result.get("subject"),
            "webLink": result.get("webLink"),
        })


def _calendar_delete(args, token: str):
    result = _graph_delete(f"/me/events/{args.event_id}", token)
    result["eventId"] = args.event_id
    _out(result)


# ---------------------------------------------------------------------------
# Teams
# ---------------------------------------------------------------------------

def _teams_list(args, token: str):
    data = _graph_get("/me/joinedTeams", token)
    if "error" in data:
        _out(data)
        return
    teams = [{"id": t.get("id"), "name": t.get("displayName"), "description": t.get("description")}
             for t in data.get("value", [])]
    _out(teams)


def _teams_channels(args, token: str):
    data = _graph_get(f"/teams/{args.team_id}/channels", token)
    if "error" in data:
        _out(data)
        return
    channels = [{"id": c.get("id"), "name": c.get("displayName"), "description": c.get("description")}
                for c in data.get("value", [])]
    _out(channels)


def _teams_messages(args, token: str):
    params = {"$top": "50"}
    data = _graph_list(
        f"/teams/{args.team_id}/channels/{args.channel_id}/messages",
        token,
        params=params,
        max_items=args.max,
    )
    if "error" in data:
        _out(data)
        return
    messages = []
    for m in data.get("value", []):
        sender = m.get("from") or {}
        # Graph can return {"from": null}, {"from": {"user": null}}, or a system-message
        # shape with {"from": {"application": {...}}} for bot/app posts.
        user = (sender.get("user") or {}) if isinstance(sender, dict) else {}
        app = (sender.get("application") or {}) if isinstance(sender, dict) else {}
        body = m.get("body") or {}
        messages.append({
            "id": m.get("id"),
            "from": (
                (user.get("displayName") if isinstance(user, dict) else None)
                or (app.get("displayName") if isinstance(app, dict) else None)
                or ""
            ),
            "date": m.get("createdDateTime"),
            "body": body.get("content", "")[:500],
            "bodyType": body.get("contentType"),
        })
    _out(messages)


def _teams_send(args, token: str):
    body = {
        "body": {"contentType": "html" if args.html else "text", "content": args.body}
    }
    result = _graph_post(f"/teams/{args.team_id}/channels/{args.channel_id}/messages", token, body)
    if "error" in result:
        _out(result)
    else:
        _out({"status": "sent", "id": result.get("id"), "teamId": args.team_id, "channelId": args.channel_id})


def _teams_chats(args, token: str):
    params = {"$top": "50", "$select": "id,topic,chatType,lastUpdatedDateTime"}
    data = _graph_list("/me/chats", token, params=params, max_items=args.max)
    if "error" in data:
        _out(data)
        return
    chats = [{"id": c.get("id"), "topic": c.get("topic"), "type": c.get("chatType"),
              "lastUpdated": c.get("lastUpdatedDateTime")}
             for c in data.get("value", [])]
    _out(chats)


def _teams_chat_messages(args, token: str):
    params = {"$top": "50"}
    data = _graph_list(f"/me/chats/{args.chat_id}/messages", token, params=params, max_items=args.max)
    if "error" in data:
        _out(data)
        return
    messages = []
    for m in data.get("value", []):
        sender = m.get("from") or {}
        # Graph can return {"from": null}, {"from": {"user": null}}, or a system-message
        # shape with {"from": {"application": {...}}} for bot/app posts.
        user = (sender.get("user") or {}) if isinstance(sender, dict) else {}
        app = (sender.get("application") or {}) if isinstance(sender, dict) else {}
        body = m.get("body") or {}
        messages.append({
            "id": m.get("id"),
            "from": (
                (user.get("displayName") if isinstance(user, dict) else None)
                or (app.get("displayName") if isinstance(app, dict) else None)
                or ""
            ),
            "date": m.get("createdDateTime"),
            "body": body.get("content", "")[:500],
            "bodyType": body.get("contentType"),
        })
    _out(messages)


def _teams_chat_send(args, token: str):
    """Send a message to a Teams chat (group or 1:1).

    Graph API endpoint: POST /chats/{chat-id}/messages
    """
    body = {
        "body": {"contentType": "html" if args.html else "text", "content": args.body}
    }
    result = _graph_post(f"/chats/{args.chat_id}/messages", token, body)
    if "error" in result:
        _out(result)
    else:
        _out({"status": "sent", "id": result.get("id"), "chatId": args.chat_id})


# ---------------------------------------------------------------------------
# OneDrive
# ---------------------------------------------------------------------------

def _onedrive_search(args, token: str):
    # The query lives inside a single-quoted OData string literal — escape
    # embedded quotes per OData (double them up) and URL-encode for safety.
    escaped = args.query.replace("'", "''")
    quoted = urllib.parse.quote(escaped, safe="")
    params = {"$top": "50"}
    data = _graph_list(
        f"/me/drive/root/search(q='{quoted}')",
        token,
        params=params,
        max_items=args.max,
    )
    if "error" in data:
        _out(data)
        return
    items = [
        {"id": i.get("id"), "name": i.get("name"), "size": i.get("size"),
         "mimeType": i.get("file", {}).get("mimeType") if i.get("file") else "folder",
         "lastModified": i.get("lastModifiedDateTime"),
         "webUrl": i.get("webUrl")}
        for i in data.get("value", [])
    ]
    _out(items)


def _onedrive_get(args, token: str):
    data = _graph_get(f"/me/drive/items/{args.item_id}", token)
    if "error" in data:
        _out(data)
        return
    _out({
        "id": data.get("id"),
        "name": data.get("name"),
        "size": data.get("size"),
        "mimeType": data.get("file", {}).get("mimeType") if data.get("file") else "folder",
        "lastModified": data.get("lastModifiedDateTime"),
        "webUrl": data.get("webUrl"),
        "createdBy": data.get("createdBy", {}).get("user", {}).get("displayName"),
        "parentPath": data.get("parentReference", {}).get("path"),
    })


def _onedrive_upload(args, token: str):
    local_path = Path(args.path).expanduser().resolve()
    if not local_path.exists():
        _out({"error": f"File not found: {local_path}"})
        return

    name = args.name or local_path.name
    parent = args.parent_id or "root"

    # Simple upload (< 4MB). For larger files, a resumable session would be needed.
    file_data = local_path.read_bytes()
    if len(file_data) > 4 * 1024 * 1024:
        _out({"error": "File too large for simple upload (>4MB). Use OneDrive web UI for large files."})
        return

    import mimetypes
    mime = mimetypes.guess_type(str(local_path))[0] or "application/octet-stream"

    if parent == "root":
        path = f"/me/drive/root:/{name}:/content"
    else:
        path = f"/me/drive/items/{parent}:/{name}:/content"

    result = _graph_put(path, token, file_data, content_type=mime)
    if "error" in result:
        _out(result)
    else:
        _out({
            "status": "uploaded",
            "id": result.get("id"),
            "name": result.get("name"),
            "size": result.get("size"),
            "webUrl": result.get("webUrl"),
        })


def _onedrive_download(args, token: str):
    # First get metadata to find the filename
    meta = _graph_get(f"/me/drive/items/{args.item_id}", token)
    if "error" in meta:
        _out(meta)
        return

    filename = meta.get("name", args.item_id)
    output_path = Path(args.output).expanduser().resolve() if args.output else Path.cwd() / filename

    content = _graph_get_binary(f"/me/drive/items/{args.item_id}/content", token)
    output_path.write_bytes(content)

    _out({
        "status": "downloaded",
        "id": args.item_id,
        "name": filename,
        "path": str(output_path),
        "size": len(content),
    })


def _onedrive_create_folder(args, token: str):
    parent_id = args.parent_id or "root"
    body = {
        "name": args.name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename",
    }

    if parent_id == "root":
        path = "/me/drive/root/children"
    else:
        path = f"/me/drive/items/{parent_id}/children"

    result = _graph_post(path, token, body)
    if "error" in result:
        _out(result)
    else:
        _out({
            "status": "created",
            "id": result.get("id"),
            "name": result.get("name"),
            "webUrl": result.get("webUrl"),
        })


def _onedrive_share(args, token: str):
    link_type = "view" if args.type == "view" else "edit"
    body = {"type": link_type, "scope": "anonymous"}

    result = _graph_post(f"/me/drive/items/{args.item_id}/createLink", token, body)
    if "error" in result:
        _out(result)
    else:
        link = result.get("link", {})
        _out({
            "status": "shared",
            "itemId": args.item_id,
            "type": link_type,
            "webUrl": link.get("webUrl"),
        })


# ---------------------------------------------------------------------------
# Argparse
# ---------------------------------------------------------------------------

def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="o365_api.py",
        description="Microsoft 365 Graph API CLI for Hermes Agent",
    )
    subs = parser.add_subparsers(dest="service", required=True)

    # --- Outlook ---
    outlook = subs.add_parser("outlook", help="Outlook mail operations")
    outlook_subs = outlook.add_subparsers(dest="command", required=True)

    # search
    os_search = outlook_subs.add_parser("search", help="Search messages")
    os_search.add_argument("query", help="KQL search query or OData filter")
    os_search.add_argument("--max", type=int, default=10, help="Max results (default: 10)")
    os_search.add_argument("--filter", action="store_true", help="Use $filter instead of $search")
    os_search.add_argument("--folder-id", dest="folder_id", help="Restrict search to a specific mailFolder id")

    # get
    os_get = outlook_subs.add_parser("get", help="Get a single message")
    os_get.add_argument("message_id", help="Message ID")
    os_get.add_argument(
        "--max-body",
        type=int,
        default=50000,
        help="Truncate message body to N chars (default: 50000, 0 = unlimited)",
    )

    # send
    os_send = outlook_subs.add_parser("send", help="Send an email")
    os_send.add_argument("--to", required=True, help="Recipient(s), comma-separated")
    os_send.add_argument("--cc", help="CC recipient(s), comma-separated")
    os_send.add_argument("--subject", required=True, help="Subject line")
    os_send.add_argument("--body", required=True, help="Message body")
    os_send.add_argument("--html", action="store_true", help="Send body as HTML")
    os_send.add_argument("--attachment", action="append", metavar="PATH",
                         help="File path to attach (repeatable)")

    # reply
    os_reply = outlook_subs.add_parser("reply", help="Reply to a message")
    os_reply.add_argument("message_id", help="Message ID to reply to")
    os_reply.add_argument("--body", required=True, help="Reply body")

    # folders
    outlook_subs.add_parser("folders", help="List mail folders")

    # move
    os_move = outlook_subs.add_parser("move", help="Move a message to a folder")
    os_move.add_argument("message_id", help="Message ID to move")
    os_move.add_argument("destination_id", help="Destination mailFolder ID")

    # --- Calendar ---
    cal = subs.add_parser("calendar", help="Calendar operations")
    cal_subs = cal.add_subparsers(dest="command", required=True)

    # list
    cal_list = cal_subs.add_parser("list", help="List events")
    cal_list.add_argument("--start", help="Start datetime (ISO 8601)")
    cal_list.add_argument("--end", help="End datetime (ISO 8601)")
    cal_list.add_argument("--max", type=int, default=25, help="Max results (default: 25)")

    # create
    cal_create = cal_subs.add_parser("create", help="Create an event")
    cal_create.add_argument("--summary", required=True, help="Event title")
    cal_create.add_argument("--start", required=True, help="Start datetime (ISO 8601)")
    cal_create.add_argument("--end", required=True, help="End datetime (ISO 8601)")
    cal_create.add_argument("--timezone", default="UTC", help="Timezone (default: UTC)")
    cal_create.add_argument("--location", help="Event location")
    cal_create.add_argument("--body", help="Event description")
    cal_create.add_argument("--attendees", help="Attendee emails, comma-separated")

    # delete
    cal_del = cal_subs.add_parser("delete", help="Delete an event")
    cal_del.add_argument("event_id", help="Event ID")

    # --- Teams ---
    teams = subs.add_parser("teams", help="Microsoft Teams operations")
    teams_subs = teams.add_subparsers(dest="command", required=True)

    # list
    teams_subs.add_parser("list", help="List joined teams")

    # channels
    t_channels = teams_subs.add_parser("channels", help="List channels in a team")
    t_channels.add_argument("team_id", help="Team ID")

    # messages
    t_messages = teams_subs.add_parser("messages", help="Get channel messages")
    t_messages.add_argument("team_id", help="Team ID")
    t_messages.add_argument("channel_id", help="Channel ID")
    t_messages.add_argument("--max", type=int, default=20, help="Max results (default: 20)")

    # send
    t_send = teams_subs.add_parser("send", help="Send a channel message")
    t_send.add_argument("team_id", help="Team ID")
    t_send.add_argument("channel_id", help="Channel ID")
    t_send.add_argument("--body", required=True, help="Message body")
    t_send.add_argument("--html", action="store_true", help="Send body as HTML")

    # chats
    t_chats = teams_subs.add_parser("chats", help="List chats")
    t_chats.add_argument("--max", type=int, default=20, help="Max results (default: 20)")

    # chat-messages
    t_chat_msgs = teams_subs.add_parser("chat-messages", help="Get chat messages")
    t_chat_msgs.add_argument("chat_id", help="Chat ID")
    t_chat_msgs.add_argument("--max", type=int, default=20, help="Max results (default: 20)")

    # chat-send
    t_chat_send = teams_subs.add_parser("chat-send", help="Send a chat message")
    t_chat_send.add_argument("chat_id", help="Chat ID")
    t_chat_send.add_argument("--body", required=True, help="Message body")
    t_chat_send.add_argument("--html", action="store_true", help="Send body as HTML")

    # --- OneDrive ---
    drive = subs.add_parser("onedrive", help="OneDrive file operations")
    drive_subs = drive.add_subparsers(dest="command", required=True)

    # search
    d_search = drive_subs.add_parser("search", help="Search files")
    d_search.add_argument("query", help="Search query")
    d_search.add_argument("--max", type=int, default=10, help="Max results (default: 10)")

    # get
    d_get = drive_subs.add_parser("get", help="Get file metadata")
    d_get.add_argument("item_id", help="Item ID")

    # upload
    d_upload = drive_subs.add_parser("upload", help="Upload a file (<4MB)")
    d_upload.add_argument("path", help="Local file path")
    d_upload.add_argument("--name", help="Override filename")
    d_upload.add_argument("--parent-id", help="Parent folder ID (default: root)")

    # download
    d_download = drive_subs.add_parser("download", help="Download a file")
    d_download.add_argument("item_id", help="Item ID")
    d_download.add_argument("--output", help="Output file path")

    # create-folder
    d_mkdir = drive_subs.add_parser("create-folder", help="Create a folder")
    d_mkdir.add_argument("name", help="Folder name")
    d_mkdir.add_argument("--parent-id", help="Parent folder ID (default: root)")

    # share
    d_share = drive_subs.add_parser("share", help="Create a sharing link")
    d_share.add_argument("item_id", help="Item ID")
    d_share.add_argument("--type", choices=["view", "edit"], default="view", help="Link type (default: view)")

    return parser


# ---------------------------------------------------------------------------
# Dispatch
# ---------------------------------------------------------------------------

_DISPATCH = {
    ("outlook", "search"): _outlook_search,
    ("outlook", "get"): _outlook_get,
    ("outlook", "send"): _outlook_send,
    ("outlook", "reply"): _outlook_reply,
    ("outlook", "folders"): _outlook_folders,
    ("outlook", "move"): _outlook_move,
    ("calendar", "list"): _calendar_list,
    ("calendar", "create"): _calendar_create,
    ("calendar", "delete"): _calendar_delete,
    ("teams", "list"): _teams_list,
    ("teams", "channels"): _teams_channels,
    ("teams", "messages"): _teams_messages,
    ("teams", "send"): _teams_send,
    ("teams", "chats"): _teams_chats,
    ("teams", "chat-messages"): _teams_chat_messages,
    ("teams", "chat-send"): _teams_chat_send,
    ("onedrive", "search"): _onedrive_search,
    ("onedrive", "get"): _onedrive_get,
    ("onedrive", "upload"): _onedrive_upload,
    ("onedrive", "download"): _onedrive_download,
    ("onedrive", "create-folder"): _onedrive_create_folder,
    ("onedrive", "share"): _onedrive_share,
}


def main():
    parser = _build_parser()
    args = parser.parse_args()

    token = _get_token()
    handler = _DISPATCH.get((args.service, args.command))
    if handler:
        handler(args, token)
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
