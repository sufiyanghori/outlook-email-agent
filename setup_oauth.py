"""
OAuth setup script for Microsoft Graph API.

Run this once to authenticate your Outlook account:
    python setup_oauth.py

It will:
1. Start a local server on http://localhost:3000
2. Open your browser to Microsoft login
3. Capture the authorization code automatically
4. Exchange it for access + refresh tokens
5. Save tokens to outlook_tokens.json

After setup, use email_client.py to interact with your inbox.
"""

import http.server
import urllib.parse
import urllib.request
import webbrowser
import json
import threading
import os

from config import CLIENT_ID, CLIENT_SECRET, REDIRECT_URI, SCOPES, save_tokens

TOKEN_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outlook_tokens.json")


class OAuthHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/callback":
            params = urllib.parse.parse_qs(parsed.query)
            if "code" in params:
                code = params["code"][0]
                token_data = self._exchange_code(code)
                if "access_token" in token_data:
                    save_tokens(token_data)
                    self._respond(200, "<h1>Success! You can close this tab.</h1>"
                                       "<p>Tokens saved. Go back to your terminal / Cursor.</p>")
                    print("\n" + "=" * 50)
                    print("SUCCESS — OAuth tokens saved!")
                    print(f"  Token file: {TOKEN_FILE}")
                    print(f"  Scopes:     {token_data.get('scope', 'N/A')}")
                    print(f"  Expires in: {token_data.get('expires_in', '?')}s")
                    print("=" * 50 + "\n")
                    print("You can now use: python email_client.py inbox")
                else:
                    error_msg = json.dumps(token_data, indent=2)
                    self._respond(400, f"<h1>Error exchanging code</h1><pre>{error_msg}</pre>")
                    print(f"\nERROR: {error_msg}\n")
            elif "error" in params:
                error = params.get("error", ["unknown"])[0]
                desc = params.get("error_description", [""])[0]
                self._respond(400, f"<h1>Auth Error</h1><p>{error}: {desc}</p>")
                print(f"\nAuth error: {error} — {desc}\n")
            else:
                self._respond(400, "<h1>Error: No code received</h1>")

            threading.Thread(target=self.server.shutdown, daemon=True).start()

    def _exchange_code(self, code):
        data = urllib.parse.urlencode({
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope": SCOPES,
            "code": code,
            "redirect_uri": REDIRECT_URI,
            "grant_type": "authorization_code",
        }).encode()
        req = urllib.request.Request(
            "https://login.microsoftonline.com/common/oauth2/v2.0/token",
            data=data,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )
        try:
            with urllib.request.urlopen(req) as resp:
                return json.loads(resp.read())
        except urllib.error.HTTPError as e:
            return json.loads(e.read())

    def _respond(self, status, html):
        self.send_response(status)
        self.send_header("Content-Type", "text/html")
        self.end_headers()
        self.wfile.write(html.encode())

    def log_message(self, format, *args):
        pass


if __name__ == "__main__":
    port = int(REDIRECT_URI.split(":")[-1].split("/")[0])

    auth_url = (
        f"https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
        f"?client_id={CLIENT_ID}"
        f"&response_type=code"
        f"&redirect_uri={urllib.parse.quote(REDIRECT_URI)}"
        f"&response_mode=query"
        f"&scope={urllib.parse.quote(SCOPES)}"
    )

    server = http.server.HTTPServer(("localhost", port), OAuthHandler)
    print(f"Starting OAuth server on http://localhost:{port} ...")
    print("Opening browser for Microsoft login...")
    webbrowser.open(auth_url)
    print("Waiting for login callback...\n")
    server.serve_forever()
    print("Done.")
