# Outlook Email Agent

Access your personal Microsoft Outlook inbox programmatically using the Microsoft Graph API. Read, search, send, and manage emails from the command line or import as a Python module.

**Zero external dependencies** — uses only Python standard library.

---

## Quick Start

```bash
# 1. Clone the repo
git clone https://github.com/sufiyanghori/outlook-email-agent.git
cd outlook-email-agent

# 2. Copy and fill in your credentials
cp .env.example .env
# Edit .env with your CLIENT_ID and CLIENT_SECRET

# 3. Run OAuth setup (opens browser, logs you in, saves tokens)
python3 setup_oauth.py

# 4. Use it
python3 email_client.py inbox
python3 email_client.py search "invoice"
python3 email_client.py send "user@example.com" "Subject" "Hello!"
```

---

## Azure App Registration Setup

If you don't already have an Azure app registered, follow these steps:

### 1. Register the App

1. Go to [Azure Portal](https://portal.azure.com)
2. Search for **App registrations** → click **New registration**
3. Settings:
   - **Name:** `outlook-email-agent` (or anything)
   - **Supported account types:** "Accounts in any organizational directory and personal Microsoft accounts"
   - **Redirect URI:** Web → `http://localhost:3000/callback`
4. Click **Register**
5. Copy the **Application (client) ID** from the Overview page

### 2. Create a Client Secret

1. Go to **Certificates & secrets** → **New client secret**
2. Description: `cursor-access` (or anything), choose expiry
3. Click **Add**
4. **IMMEDIATELY copy the Value** (it's only shown once)
   - The "Value" column is the secret — NOT the "Secret ID" column

### 3. Add API Permissions

1. Go to **API permissions** → **Add a permission**
2. Choose **Microsoft Graph** → **Delegated permissions**
3. Add:
   - `Mail.Read`
   - `Mail.Send`
   - `offline_access`
   - `User.Read`
4. Click **Grant admin consent** if available

### 4. Configure .env

```bash
cp .env.example .env
```

Edit `.env`:

```
CLIENT_ID=your-application-client-id
CLIENT_SECRET=your-client-secret-value
REDIRECT_URI=http://localhost:3000/callback
SCOPES=https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send offline_access
```

### 5. Authenticate

```bash
python3 setup_oauth.py
```

This will:
1. Start a local server on `http://localhost:3000`
2. Open your browser to the Microsoft login page
3. After you log in and approve, it captures the auth code automatically
4. Exchanges the code for access + refresh tokens
5. Saves tokens to `outlook_tokens.json`

---

## CLI Usage

```bash
# List recent emails (default: 10)
python3 email_client.py inbox
python3 email_client.py inbox 20

# Read a specific email by ID
python3 email_client.py read <message_id>

# Search emails
python3 email_client.py search "quarterly report"

# Send an email
python3 email_client.py send "recipient@example.com" "Subject line" "Email body text"

# List mail folders
python3 email_client.py folders

# Show authenticated user info
python3 email_client.py me
```

---

## Python Module Usage

```python
from email_client import OutlookClient

client = OutlookClient()

# Read inbox
emails = client.list_inbox(top=5)
for email in emails:
    print(client.format_email(email))

# Search
results = client.search_emails("invoice")

# Send email
client.send_email("user@example.com", "Hello", "This is the body")

# Send HTML email
client.send_email("user@example.com", "Hello", "<h1>HTML body</h1>", body_type="HTML")

# Reply to an email
client.reply_to_email(message_id, "Thanks for your email!")

# Create a draft
client.create_draft("user@example.com", "Draft Subject", "Draft body")

# Mark as read
client.mark_as_read(message_id)

# List folders
folders = client.list_folders()
```

---

## How It Works

### Authentication Flow

```
User runs setup_oauth.py
       ↓
Local server starts on http://localhost:3000
       ↓
Browser opens → Microsoft login page
       ↓
User logs in and approves permissions
       ↓
Microsoft redirects to http://localhost:3000/callback?code=AUTH_CODE
       ↓
Local server captures the code
       ↓
Exchanges code for access_token + refresh_token
       ↓
Tokens saved to outlook_tokens.json
```

### Token Refresh

- Access tokens expire after ~1 hour
- The `config.py` module automatically refreshes using the refresh token on 401 errors
- No need to re-authenticate manually

### API Endpoints Used

| Action | Method | Endpoint |
|--------|--------|----------|
| List emails | GET | `/me/mailFolders/{folder}/messages` |
| Get email | GET | `/me/messages/{id}` |
| Search | GET | `/me/messages?$search=` |
| Send | POST | `/me/sendMail` |
| Reply | POST | `/me/messages/{id}/reply` |
| Create draft | POST | `/me/messages` |
| Mark read | PATCH | `/me/messages/{id}` |
| Delete | DELETE | `/me/messages/{id}` |
| List folders | GET | `/me/mailFolders` |
| User info | GET | `/me` |

---

## Project Structure

```
outlook-email-agent/
├── .env.example       # Template for credentials
├── .env               # Your credentials (git-ignored)
├── .gitignore         # Excludes tokens and secrets
├── config.py          # Configuration, token management, Graph API helper
├── email_client.py    # Email operations (read/search/send) + CLI
├── setup_oauth.py     # One-time OAuth authentication script
├── outlook_tokens.json # Saved tokens (git-ignored)
└── README.md          # This file
```

---

## Security Notes

- **Never commit** `.env` or `outlook_tokens.json` — they are in `.gitignore`
- Client secrets expire (check your Azure app settings)
- Refresh tokens can be revoked from [Microsoft account security](https://account.microsoft.com/security)
- The redirect URI (`localhost:3000`) only works locally — no remote exposure

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| `bash: !...: event not found` | Use single quotes around the auth code in curl commands |
| `invalid_client` / wrong secret | Go to Azure → Certificates & secrets → create a **new** secret → copy the **Value** (not Secret ID) |
| `invalid_grant` / code expired | Auth codes are single-use and expire in ~10 minutes. Re-run `setup_oauth.py` |
| `AADSTS700016` | Check CLIENT_ID in .env matches your Azure app |
| Token refresh fails | Re-run `setup_oauth.py` to get fresh tokens |
| `No module named 'requests'` | This project uses only stdlib — no pip install needed |

---

## Using with Cursor AI

Once set up, Cursor can interact with your Outlook inbox by importing `email_client.py`:

```python
from email_client import OutlookClient

client = OutlookClient()
emails = client.list_inbox(top=5)

# AI can now process, summarize, reply to emails
```

Or run CLI commands directly from Cursor's terminal:

```bash
python3 email_client.py inbox 5
python3 email_client.py search "meeting"
```
