"""
Outlook email client using Microsoft Graph API.

Usage:
    from email_client import OutlookClient

    client = OutlookClient()
    emails = client.list_inbox(top=10)
    client.send_email("user@example.com", "Subject", "Body text")
"""

from config import graph_request


class OutlookClient:
    """High-level client for reading and sending Outlook emails."""

    # ── Read ─────────────────────────────────────────────────

    def list_inbox(self, top=10, skip=0, folder="inbox"):
        """List recent emails. Returns a list of email summaries."""
        endpoint = (
            f"/me/mailFolders/{folder}/messages"
            f"?$top={top}&$skip={skip}"
            f"&$select=subject,from,toRecipients,receivedDateTime,isRead,bodyPreview"
            f"&$orderby=receivedDateTime%20desc"
        )
        result = graph_request(endpoint)
        return result.get("value", [])

    def get_email(self, message_id):
        """Get a single email with full body."""
        endpoint = f"/me/messages/{message_id}"
        return graph_request(endpoint)

    def search_emails(self, query, top=10):
        """
        Search emails by keyword.
        Searches subject, body, and sender fields.
        """
        endpoint = (
            f"/me/messages"
            f"?$search=\"{query}\""
            f"&$top={top}"
            f"&$select=subject,from,toRecipients,receivedDateTime,isRead,bodyPreview"
        )
        return graph_request(endpoint).get("value", [])

    def list_folders(self):
        """List all mail folders."""
        endpoint = "/me/mailFolders?$top=50"
        return graph_request(endpoint).get("value", [])

    # ── Write ────────────────────────────────────────────────

    def send_email(self, to, subject, body, body_type="Text", cc=None, bcc=None):
        """
        Send an email.

        Args:
            to: recipient email address (str) or list of addresses
            subject: email subject
            body: email body content
            body_type: "Text" or "HTML"
            cc: optional CC address(es)
            bcc: optional BCC address(es)
        """
        if isinstance(to, str):
            to = [to]
        if isinstance(cc, str):
            cc = [cc]
        if isinstance(bcc, str):
            bcc = [bcc]

        message = {
            "message": {
                "subject": subject,
                "body": {"contentType": body_type, "content": body},
                "toRecipients": [
                    {"emailAddress": {"address": addr}} for addr in to
                ],
            },
            "saveToSentItems": True,
        }

        if cc:
            message["message"]["ccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in cc
            ]
        if bcc:
            message["message"]["bccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in bcc
            ]

        return graph_request("/me/sendMail", method="POST", body=message)

    def create_draft(self, to, subject, body, body_type="Text"):
        """Create a draft email (does not send)."""
        if isinstance(to, str):
            to = [to]

        message = {
            "subject": subject,
            "body": {"contentType": body_type, "content": body},
            "toRecipients": [
                {"emailAddress": {"address": addr}} for addr in to
            ],
        }
        return graph_request("/me/messages", method="POST", body=message)

    def reply_to_email(self, message_id, comment, reply_all=False):
        """Reply to an email."""
        action = "replyAll" if reply_all else "reply"
        endpoint = f"/me/messages/{message_id}/{action}"
        return graph_request(endpoint, method="POST", body={"comment": comment})

    def mark_as_read(self, message_id):
        """Mark an email as read."""
        endpoint = f"/me/messages/{message_id}"
        return graph_request(endpoint, method="PATCH", body={"isRead": True})

    def delete_email(self, message_id):
        """Move an email to trash."""
        endpoint = f"/me/messages/{message_id}"
        return graph_request(endpoint, method="DELETE")

    # ── Helpers ──────────────────────────────────────────────

    def get_me(self):
        """Get the authenticated user's profile."""
        return graph_request("/me")

    @staticmethod
    def format_email(email):
        """Format an email dict into a readable string."""
        sender = email.get("from", {}).get("emailAddress", {})
        return (
            f"From:    {sender.get('name', 'Unknown')} <{sender.get('address', '')}>\n"
            f"Date:    {email.get('receivedDateTime', 'N/A')}\n"
            f"Subject: {email.get('subject', '(no subject)')}\n"
            f"Read:    {'Yes' if email.get('isRead') else 'No'}\n"
            f"Preview: {email.get('bodyPreview', '')[:200]}\n"
        )


# ── CLI usage ────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    client = OutlookClient()

    if len(sys.argv) < 2:
        print("Usage:")
        print("  python email_client.py inbox [count]")
        print("  python email_client.py read <message_id>")
        print("  python email_client.py search <query>")
        print("  python email_client.py send <to> <subject> <body>")
        print("  python email_client.py folders")
        print("  python email_client.py me")
        sys.exit(1)

    command = sys.argv[1]

    if command == "inbox":
        count = int(sys.argv[2]) if len(sys.argv) > 2 else 10
        emails = client.list_inbox(top=count)
        for i, email in enumerate(emails, 1):
            print(f"─── Email {i} ───")
            print(client.format_email(email))

    elif command == "read":
        msg = client.get_email(sys.argv[2])
        print(f"Subject: {msg.get('subject')}")
        print(f"From: {msg['from']['emailAddress']['name']} <{msg['from']['emailAddress']['address']}>")
        print(f"Date: {msg.get('receivedDateTime')}")
        print(f"\n{msg.get('body', {}).get('content', '')}")

    elif command == "search":
        query = " ".join(sys.argv[2:])
        emails = client.search_emails(query)
        for i, email in enumerate(emails, 1):
            print(f"─── Result {i} ───")
            print(client.format_email(email))

    elif command == "send":
        if len(sys.argv) < 5:
            print("Usage: python email_client.py send <to> <subject> <body>")
            sys.exit(1)
        client.send_email(sys.argv[2], sys.argv[3], " ".join(sys.argv[4:]))
        print("Email sent successfully!")

    elif command == "folders":
        folders = client.list_folders()
        for f in folders:
            print(f"  {f['displayName']} ({f.get('totalItemCount', '?')} emails, {f.get('unreadItemCount', '?')} unread)")

    elif command == "me":
        me = client.get_me()
        print(f"Name:  {me.get('displayName', 'N/A')}")
        print(f"Email: {me.get('mail') or me.get('userPrincipalName', 'N/A')}")

    else:
        print(f"Unknown command: {command}")
        sys.exit(1)
