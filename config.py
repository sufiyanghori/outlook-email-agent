"""
Configuration and token management for Microsoft Graph API.
Loads credentials from .env file and handles token refresh automatically.
"""

import os
import json
import urllib.parse
import urllib.request

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_FILE = os.path.join(BASE_DIR, "outlook_tokens.json")
ENV_FILE = os.path.join(BASE_DIR, ".env")


def _load_env():
    """Load key=value pairs from .env file into a dict."""
    env = {}
    if not os.path.exists(ENV_FILE):
        raise FileNotFoundError(
            f".env file not found at {ENV_FILE}. "
            "Copy .env.example to .env and fill in your credentials."
        )
    with open(ENV_FILE) as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, value = line.split("=", 1)
                env[key.strip()] = value.strip()
    return env


_env = _load_env()

CLIENT_ID = _env.get("CLIENT_ID", "")
CLIENT_SECRET = _env.get("CLIENT_SECRET", "")
REDIRECT_URI = _env.get("REDIRECT_URI", "http://localhost:3000/callback")
SCOPES = _env.get(
    "SCOPES",
    "https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send offline_access",
)
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def load_tokens():
    """Load saved tokens from disk."""
    if not os.path.exists(TOKEN_FILE):
        return None
    with open(TOKEN_FILE) as f:
        return json.load(f)


def save_tokens(token_data):
    """Persist tokens to disk."""
    with open(TOKEN_FILE, "w") as f:
        json.dump(token_data, f, indent=2)


def refresh_access_token():
    """
    Use the refresh_token to get a new access_token.
    Returns the updated token dict and saves it to disk.
    """
    tokens = load_tokens()
    if not tokens or "refresh_token" not in tokens:
        raise RuntimeError(
            "No refresh token found. Run `python setup_oauth.py` first to authenticate."
        )

    data = urllib.parse.urlencode({
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": SCOPES,
        "refresh_token": tokens["refresh_token"],
        "grant_type": "refresh_token",
    }).encode()

    req = urllib.request.Request(
        "https://login.microsoftonline.com/common/oauth2/v2.0/token",
        data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )

    try:
        with urllib.request.urlopen(req) as resp:
            new_tokens = json.loads(resp.read())
    except urllib.error.HTTPError as e:
        error_body = json.loads(e.read())
        raise RuntimeError(f"Token refresh failed: {error_body}")

    if "access_token" not in new_tokens:
        raise RuntimeError(f"Token refresh failed: {new_tokens}")

    save_tokens(new_tokens)
    return new_tokens


def get_access_token():
    """
    Get a valid access token. Tries the saved token first,
    and if it fails a Graph API call, refreshes automatically.
    """
    tokens = load_tokens()
    if not tokens:
        raise RuntimeError(
            "No tokens found. Run `python setup_oauth.py` first to authenticate."
        )
    return tokens["access_token"]


def graph_request(endpoint, method="GET", body=None, auto_refresh=True):
    """
    Make an authenticated request to Microsoft Graph API.
    Automatically refreshes the token on 401.
    """
    token = get_access_token()
    url = f"{GRAPH_BASE}{endpoint}" if endpoint.startswith("/") else endpoint

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    data = json.dumps(body).encode() if body else None
    req = urllib.request.Request(url, data=data, headers=headers, method=method)

    try:
        with urllib.request.urlopen(req) as resp:
            if resp.status == 204:
                return {}
            return json.loads(resp.read())
    except urllib.error.HTTPError as e:
        if e.code == 401 and auto_refresh:
            refresh_access_token()
            return graph_request(endpoint, method, body, auto_refresh=False)
        error_body = e.read().decode()
        try:
            error_json = json.loads(error_body)
            raise RuntimeError(f"Graph API error ({e.code}): {json.dumps(error_json, indent=2)}")
        except json.JSONDecodeError:
            raise RuntimeError(f"Graph API error ({e.code}): {error_body}")
