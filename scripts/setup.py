#!/usr/bin/env python3
"""Microsoft 365 OAuth2 setup for Hermes Agent (device code flow).

Fully non-interactive — designed to be driven by the agent via terminal commands.
Uses MSAL device code flow: no client secret, no redirect URI, no web server.

Commands:
  setup.py --check                # Is auth valid? Exit 0 = yes, 1 = no
  setup.py --check-live           # Verify with a real Graph API call (/me)
  setup.py --auth                 # Start device code flow (interactive polling)
  setup.py --revoke               # Delete stored tokens
  setup.py --install-deps         # Install Python dependencies only

Agent workflow:
  1. Run --check. If exit 0, auth is good — skip setup.
  2. If no app config exists, ask user for Azure AD client ID.
  3. Run --auth. Script prints device login URL + code, then polls for completion.
  4. Run --check to verify. Done.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

# Ensure sibling modules (_hermes_home) are importable when run standalone.
_SCRIPTS_DIR = str(Path(__file__).resolve().parent)
if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

from _hermes_home import display_hermes_home, get_hermes_home

HERMES_HOME = get_hermes_home()
TOKEN_CACHE_PATH = HERMES_HOME / "o365_token_cache.bin"
APP_CONFIG_PATH = HERMES_HOME / "o365_app_config.json"

AUTHORITY = "https://login.microsoftonline.com/" + os.environ.get("O365_TENANT_ID", "common")

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

REQUIRED_PACKAGES = ["msal", "requests"]


def _write_secret(path: Path, data: str) -> None:
    """Write a secret file atomically with 0600 permissions."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(data)
    os.chmod(tmp, 0o600)
    os.replace(tmp, path)


def install_deps() -> bool:
    """Install MSAL and requests if missing. Returns True on success."""
    try:
        import msal  # noqa: F401
        import requests  # noqa: F401
        print("Dependencies already installed.")
        return True
    except ImportError:
        pass

    print("Installing Microsoft 365 dependencies...")
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet"] + REQUIRED_PACKAGES,
            stdout=subprocess.DEVNULL,
        )
        print("Dependencies installed.")
        return True
    except subprocess.CalledProcessError as e:
        print(f"ERROR: Failed to install dependencies: {e}")
        print(f"Manually: {sys.executable} -m pip install {' '.join(REQUIRED_PACKAGES)}")
        return False


def _ensure_deps():
    """Check deps are available, install if not, exit on failure."""
    try:
        import msal  # noqa: F401
        import requests  # noqa: F401
    except ImportError:
        if not install_deps():
            sys.exit(1)


def _load_app_config() -> dict | None:
    """Load the app config (client_id). Returns None if not found."""
    if not APP_CONFIG_PATH.exists():
        return None
    try:
        data = json.loads(APP_CONFIG_PATH.read_text())
        if data.get("client_id"):
            return data
    except Exception:
        pass
    return None


def _save_app_config(client_id: str):
    """Save the Azure AD app client ID."""
    _write_secret(APP_CONFIG_PATH, json.dumps({"client_id": client_id}, indent=2))
    print(f"OK: App config saved to {APP_CONFIG_PATH}")


def _build_msal_app(client_id: str):
    """Create an MSAL PublicClientApplication with persistent token cache."""
    import msal

    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        cache.deserialize(TOKEN_CACHE_PATH.read_text())

    app = msal.PublicClientApplication(
        client_id,
        authority=AUTHORITY,
        token_cache=cache,
    )
    return app, cache


def _save_cache(cache):
    """Persist the MSAL token cache if it changed."""
    if cache.has_state_changed:
        _write_secret(TOKEN_CACHE_PATH, cache.serialize())


def _get_token_silent(app) -> dict | None:
    """Try to acquire a token silently (from cache / refresh)."""
    accounts = app.get_accounts()
    if not accounts:
        return None
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if result and "access_token" in result:
        return result
    return None


def check_auth(quiet: bool = False) -> bool:
    """Check if stored credentials are valid. Prints status."""
    config = _load_app_config()
    if not config:
        print(f"NOT_AUTHENTICATED: No app config at {APP_CONFIG_PATH}")
        return False

    if not TOKEN_CACHE_PATH.exists():
        print(f"NOT_AUTHENTICATED: No token cache at {TOKEN_CACHE_PATH}")
        return False

    _ensure_deps()
    app, cache = _build_msal_app(config["client_id"])

    accounts = app.get_accounts()
    if not accounts:
        print("NOT_AUTHENTICATED: No accounts in token cache.")
        return False

    result = _get_token_silent(app)
    _save_cache(cache)

    if result and "access_token" in result:
        if not quiet:
            username = accounts[0].get("username", "unknown")
            print(f"AUTHENTICATED: Token valid for {username}")
        return True

    error = result.get("error_description", "unknown error") if result else "token acquisition failed"
    print(f"TOKEN_INVALID: {error}")
    return False


def check_auth_live() -> bool:
    """Check auth with a real Graph API call to /me."""
    if not check_auth(quiet=True):
        return False

    _ensure_deps()
    import requests

    config = _load_app_config()
    app, cache = _build_msal_app(config["client_id"])
    result = _get_token_silent(app)
    _save_cache(cache)

    if not result or "access_token" not in result:
        print("LIVE_CHECK_FAILED: Could not acquire token.")
        return False

    try:
        resp = requests.get(
            "https://graph.microsoft.com/v1.0/me",
            headers={"Authorization": f"Bearer {result['access_token']}"},
            timeout=15,
        )
        if resp.status_code == 200:
            data = resp.json()
            print(f"LIVE_CHECK_OK: Authenticated as {data.get('displayName', '?')} ({data.get('mail', data.get('userPrincipalName', '?'))})")
            return True
        else:
            print(f"LIVE_CHECK_FAILED: HTTP {resp.status_code}: {resp.text[:200]}")
            return False
    except Exception as e:
        print(f"LIVE_CHECK_FAILED: {e}")
        return False


def do_auth(client_id_arg: str | None = None):
    """Run the device code OAuth flow."""
    _ensure_deps()
    import msal

    config = _load_app_config()

    # Determine client ID: argument > existing config. No interactive fallback —
    # this script must work on headless platforms (Discord, Telegram, CI).
    if client_id_arg:
        client_id = client_id_arg
        _save_app_config(client_id)
    elif config:
        client_id = config["client_id"]
    else:
        print(
            "ERROR: No app config found. Pass the Azure AD Application (client) ID:\n"
            "  setup.py --auth CLIENT_ID"
        )
        sys.exit(1)

    app, cache = _build_msal_app(client_id)

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print(f"ERROR: Failed to initiate device flow: {flow.get('error_description', json.dumps(flow))}")
        sys.exit(1)

    # Print the device code instructions for the agent to relay
    print()
    print("=" * 60)
    print("DEVICE CODE AUTHENTICATION")
    print("=" * 60)
    print(f"URL:  {flow['verification_uri']}")
    print(f"Code: {flow['user_code']}")
    print()
    print("Open the URL above and enter the code to authorize.")
    print("Waiting for authorization...")
    print("=" * 60)
    print()

    result = app.acquire_token_by_device_flow(flow)
    _save_cache(cache)

    if "access_token" in result:
        account = result.get("id_token_claims", {})
        name = account.get("name", "unknown")
        print(f"OK: Authenticated as {name}")
        print(f"Token cache saved to {display_hermes_home()}/o365_token_cache.bin")
    else:
        error = result.get("error_description", result.get("error", "unknown error"))
        print(f"ERROR: Authentication failed: {error}")
        sys.exit(1)


def revoke():
    """Delete stored tokens and app config."""
    deleted = False
    for path in [TOKEN_CACHE_PATH, APP_CONFIG_PATH]:
        if path.exists():
            path.unlink()
            print(f"Deleted {path}")
            deleted = True
    if not deleted:
        print("No tokens to revoke.")
    else:
        print("Credentials removed. Re-run --auth to re-authenticate.")


def main():
    parser = argparse.ArgumentParser(description="Microsoft 365 OAuth setup for Hermes")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--check", action="store_true", help="Check if auth is valid (exit 0=yes, 1=no)")
    group.add_argument("--check-live", action="store_true", help="Check auth with a real Graph API call")
    group.add_argument("--auth", nargs="?", const=None, default=False, metavar="CLIENT_ID",
                       help="Start device code auth flow (optionally pass client ID)")
    group.add_argument("--revoke", action="store_true", help="Delete stored tokens")
    group.add_argument("--install-deps", action="store_true", help="Install Python dependencies")
    args = parser.parse_args()

    if args.check:
        sys.exit(0 if check_auth() else 1)
    elif args.check_live:
        sys.exit(0 if check_auth_live() else 1)
    elif args.auth is not False:
        do_auth(args.auth)
    elif args.revoke:
        revoke()
    elif args.install_deps:
        sys.exit(0 if install_deps() else 1)


if __name__ == "__main__":
    main()
