"""Local month-end receipt requests. No email is sent during dry runs."""
import argparse
import calendar
import datetime as dt
import fcntl
import json
import os
from pathlib import Path
import sys
from zoneinfo import ZoneInfo

TZ = ZoneInfo('Australia/Melbourne')
START = (2026, 10)
STATE = Path.home() / 'Library/Application Support/AAMI-Month-End'


def due_months(now):
    now = now.astimezone(TZ)
    year, month = START
    due = []
    while (year, month) <= (now.year, now.month):
        end = dt.datetime(year, month, calendar.monthrange(year, month)[1], 18, tzinfo=TZ)
        if now >= end:
            due.append(f'{year:04d}-{month:02d}')
        year, month = (year + 1, 1) if month == 12 else (year, month + 1)
    return due


def message(month, member='62302014'):
    period = dt.date.fromisoformat(month + '-01').strftime('%B %Y')
    subject = f'Request for {period} payment receipts - Membership {member}'
    body = (
        '<p>Hi AAMI Health team,</p>'
        f'<p>Please email me a separate PDF invoice/receipt for each health insurance '
        f'premium payment made in {period}, including any payment on the last day '
        'of the month. These are needed for reimbursement, rather than a combined '
        'payment history.</p>'
        f'<p>Member: Muhammad Sufiyan Ghori<br>Membership number: {member}</p>'
        '<p>Thanks,<br>Sufiyan</p>'
    )
    return subject, body


def save(path, state):
    temp = path.with_suffix('.tmp')
    temp.write_text(json.dumps(state, indent=2))
    temp.chmod(0o600)
    temp.replace(path)


def run(now, directory, send, dry_run=False, member='62302014'):
    directory.mkdir(parents=True, exist_ok=True, mode=0o700)
    with (directory / 'job.lock').open('w') as lock:
        fcntl.flock(lock, fcntl.LOCK_EX)
        path = directory / 'state.json'
        state = json.loads(path.read_text()) if path.exists() else {}
        for month in due_months(now):
            status = state.get(month)
            if status == 'sent':
                continue
            if status == 'sending':
                raise RuntimeError(f'{month}: delivery uncertain; check Sent Items before resetting state')
            subject, body = message(month, member)
            if dry_run:
                print(subject, body, sep='\n')
                continue
            # Persist before sending: an uncertain network outcome must not cause duplicates.
            state[month] = 'sending'
            save(path, state)
            send(subject, body)
            state[month] = 'sent'
            save(path, state)
            print(f'{now.isoformat()}: {month} accepted by Outlook', flush=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dry-run', action='store_true')
    parser.add_argument('--preview-month')
    args = parser.parse_args()
    config = json.loads((STATE / 'config.json').read_text())
    if args.preview_month:
        print(*message(args.preview_month, config['member']), sep='\n')
        return

    def send(subject, body):
        sys.path.insert(0, config['client_directory'])
        from email_client import OutlookClient
        from config import graph_request
        # User.Read is not granted; verify sender using existing Mail.Read access.
        recent = graph_request('/me/mailFolders/sentitems/messages?$top=1&$select=from')
        sender = recent['value'][0]['from']['emailAddress']['address'].lower()
        if sender != config['sender'].lower():
            raise RuntimeError('Personal Outlook identity mismatch; refusing to send')
        OutlookClient().send_email(config['recipient'], subject, body, body_type='HTML')

    run(dt.datetime.now(TZ), STATE, send, args.dry_run, config['member'])


if __name__ == '__main__':
    os.umask(0o077)
    try:
        main()
    except Exception as error:
        print(f'AAMI automation failed: {type(error).__name__}. Check authentication and state.', file=sys.stderr)
        raise SystemExit(1)
