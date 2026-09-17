#!/usr/bin/env python3
"""
M365 Employee Data Exporter
Exports mailbox (email + attachments) and OneDrive files for specific employees.
Mirrors the workspace-classifier interface for Microsoft 365.

Usage examples:
  python3 run_m365_export.py --admin admin@company.com --user employee@company.com --export-only --local-only
  python3 run_m365_export.py --admin admin@company.com --only alice@company.com --only bob@company.com --export-only --local-only --modified-after 2024-01-01
  python3 run_m365_export.py --admin admin@company.com --user employee@company.com --export-only --local-only --since-days 90
"""

import argparse
import os
import json
import csv
import time
import sys
import base64
from datetime import datetime, timezone, timedelta
from pathlib import Path
from dotenv import load_dotenv
import msal
import requests

load_dotenv()

TENANT_ID     = os.getenv("TENANT_ID")
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
GRAPH_BASE    = "https://graph.microsoft.com/v1.0"
RETRY_STATUS  = {429, 500, 502, 503, 504}

# ── Fix 1: Build MSAL app once; MSAL caches + refreshes tokens internally ────

_MSAL_APP = None

def _get_app():
    global _MSAL_APP
    if _MSAL_APP is None:
        if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
            raise RuntimeError(
                "Missing credentials. Set TENANT_ID, CLIENT_ID, and CLIENT_SECRET in your .env file."
            )
        _MSAL_APP = msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            client_credential=CLIENT_SECRET,
        )
    return _MSAL_APP


def get_token():
    """Returns a valid token. MSAL serves from cache and refreshes near expiry."""
    result = _get_app().acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise RuntimeError(
            f"Authentication failed: {result.get('error_description', result)}"
        )
    return result["access_token"]


# ── Fix 2: Centralised retry/backoff for all HTTP calls ──────────────────────

def _request(method, url, *, auth=True, max_attempts=6, **kwargs):
    """
    Single HTTP call with exponential backoff.
    auth=False for OneDrive pre-signed URLs — adding Bearer breaks them.
    """
    for attempt in range(max_attempts):
        headers = dict(kwargs.pop("headers", {}) or {})
        if auth:
            headers["Authorization"] = f"Bearer {get_token()}"
            headers.setdefault("Accept", "application/json")
        r = requests.request(method, url, headers=headers, timeout=120, **kwargs)
        if r.status_code in RETRY_STATUS and attempt < max_attempts - 1:
            wait = int(r.headers.get("Retry-After", min(2 ** attempt, 60)))
            print(f"    HTTP {r.status_code}. Retrying in {wait}s...")
            time.sleep(wait)
            continue
        r.raise_for_status()
        return r
    raise RuntimeError(f"Gave up after {max_attempts} attempts: {url}")


# ── Fix 3: Page generator — stream instead of buffering whole mailbox ─────────

def graph_pages(url, params=None):
    """Yields one page (list of items) at a time. Never holds the whole set in RAM."""
    while url:
        data = _request("GET", url, params=params).json()
        value = data.get("value")
        yield value if value is not None else [data]
        url = data.get("@odata.nextLink")
        params = None


def graph_get(url, params=None):
    """Convenience wrapper that collects all pages — use only for small calls."""
    return [item for page in graph_pages(url, params) for item in page]


# ── Fix 4: Safe filenames — no path traversal, no null bytes ─────────────────

def safe_filename(name, fallback="file", max_len=120):
    name = os.path.basename(str(name or "")).replace("\x00", "")
    name = "".join(c if c.isalnum() or c in " .-_" else "_" for c in name).strip(" .")
    return (name or fallback)[:max_len]


def unique_path(directory, name):
    """Never clobber an existing file: foo.pdf → foo (1).pdf."""
    stem, ext = os.path.splitext(name)
    candidate, n = directory / name, 1
    while candidate.exists():
        candidate = directory / f"{stem} ({n}){ext}"
        n += 1
    return candidate


# ── Fix 6: Download with stale-URL recovery + atomic write (.part file) ──────

def download_file(upn, item, dest):
    """
    Download a OneDrive file. If the pre-signed URL has expired (403/401),
    re-fetch the item to get a fresh URL. Writes to .part first so a
    killed run never leaves a half-written file that resume then skips.
    """
    url = item.get("@microsoft.graph.downloadUrl")
    try:
        r = _request("GET", url, auth=False, stream=True)
    except requests.exceptions.HTTPError as e:
        if e.response is not None and e.response.status_code in (401, 403):
            print(f"    Download URL expired for {item.get('name')} — refreshing...")
            fresh = _request("GET", f"{GRAPH_BASE}/users/{upn}/drive/items/{item['id']}").json()
            r = _request("GET", fresh["@microsoft.graph.downloadUrl"], auth=False, stream=True)
        else:
            raise
    tmp = dest.with_suffix(dest.suffix + ".part")
    with open(tmp, "wb") as f:
        for chunk in r.iter_content(chunk_size=65536):
            f.write(chunk)
    tmp.replace(dest)  # atomic rename — .part stays if killed, dest is always complete


# ── Email export ──────────────────────────────────────────────────────────────

def write_message_file(m, email_dir, attach_dir, upn):
    msg_id = m.get("id", "unknown")
    sn = safe_filename(msg_id[:40], fallback="msg")
    msg_path = email_dir / f"{sn}.txt"

    # Fix 5: Real resume — skip messages already written
    if msg_path.exists():
        return

    body_content = m.get("body", {}).get("content", "") or m.get("bodyPreview", "")
    with open(msg_path, "w", encoding="utf-8") as f:
        f.write(f"Subject:    {m.get('subject', '')}\n")
        f.write(f"From:       {m.get('from', {}).get('emailAddress', {}).get('address', '')}\n")
        f.write(f"Sent:       {m.get('sentDateTime', '')}\n")
        f.write(f"Received:   {m.get('receivedDateTime', '')}\n")
        f.write(f"To:         {', '.join(r.get('emailAddress', {}).get('address', '') for r in m.get('toRecipients', []))}\n")
        f.write(f"CC:         {', '.join(r.get('emailAddress', {}).get('address', '') for r in m.get('ccRecipients', []))}\n")
        f.write(f"BCC:        {', '.join(r.get('emailAddress', {}).get('address', '') for r in m.get('bccRecipients', []))}\n")
        f.write("-" * 60 + "\n\n")
        f.write(body_content)

    if m.get("hasAttachments"):
        try:
            att_url = f"{GRAPH_BASE}/users/{upn}/messages/{m['id']}/attachments"
            attachments = graph_get(att_url)
            for att in attachments:
                if att.get("@odata.type") == "#microsoft.graph.fileAttachment":
                    content = att.get("contentBytes", "")
                    if content:
                        att_name = safe_filename(att.get("name", "attachment"))
                        att_path = unique_path(attach_dir, f"{sn}_{att_name}")
                        with open(att_path, "wb") as af:
                            af.write(base64.b64decode(content))
        except Exception as e:
            print(f"    Warning: could not fetch attachments for {msg_id[:20]}: {e}")


def export_emails(upn, out_dir, since_dt=None):
    print(f"  Exporting emails...")
    email_dir  = out_dir / "dump" / "emails"
    attach_dir = email_dir / "attachments"
    email_dir.mkdir(parents=True, exist_ok=True)
    attach_dir.mkdir(parents=True, exist_ok=True)

    params = {
        "$select": (
            "id,subject,from,toRecipients,ccRecipients,bccRecipients,"
            "sentDateTime,receivedDateTime,bodyPreview,hasAttachments,body"
        ),
        "$top": 100,
        "$orderby": "receivedDateTime desc",
    }
    if since_dt:
        params["$filter"] = (
            f"receivedDateTime ge {since_dt.strftime('%Y-%m-%dT%H:%M:%SZ')}"
        )

    url = f"{GRAPH_BASE}/users/{upn}/messages"

    # Fix 7: Fallback if tenant rejects $filter + $orderby together
    try:
        pages = graph_pages(url, params)
        first_page = next(pages, [])
    except requests.exceptions.HTTPError as e:
        if e.response is not None and e.response.status_code == 400:
            print("    Tenant rejected filter+sort; retrying without $orderby...")
            params.pop("$orderby", None)
            pages = graph_pages(url, params)
            first_page = next(pages, [])
        else:
            raise

    # Fix 3: Stream — write JSON and CSV incrementally, one page at a time
    json_path = email_dir / "emails_all.json"
    csv_path  = email_dir / "emails_metadata.csv"

    with open(json_path, "w", encoding="utf-8") as jf, \
         open(csv_path,  "w", newline="", encoding="utf-8") as cf:

        jf.write("[\n")
        writer = csv.writer(cf)
        writer.writerow(["id", "subject", "from_address", "sentDateTime",
                          "receivedDateTime", "hasAttachments"])

        first_json = True
        count = 0

        def process_page(page):
            nonlocal first_json, count
            for m in page:
                if not first_json:
                    jf.write(",\n")
                json.dump(m, jf, indent=2, ensure_ascii=False)
                first_json = False
                count += 1
                writer.writerow([
                    m.get("id", ""),
                    m.get("subject", ""),
                    m.get("from", {}).get("emailAddress", {}).get("address", ""),
                    m.get("sentDateTime", ""),
                    m.get("receivedDateTime", ""),
                    m.get("hasAttachments", False),
                ])
                write_message_file(m, email_dir, attach_dir, upn)

        process_page(first_page)
        for page in pages:
            process_page(page)

        jf.write("\n]\n")

    print(f"  ✓ {count} emails exported.")
    return count


# ── OneDrive export ───────────────────────────────────────────────────────────

def export_onedrive(upn, out_dir, since_dt=None, folder_path="/", local_dir=None, depth=0):
    """Recursively download OneDrive files, preserving folder structure."""
    if depth > 50:
        print(f"  Warning: max folder depth reached at {folder_path} — skipping subtree.")
        return
    if local_dir is None:
        local_dir = out_dir / "dump" / "files"
        local_dir.mkdir(parents=True, exist_ok=True)
        print(f"  Exporting OneDrive files...")

    if folder_path == "/":
        url = f"{GRAPH_BASE}/users/{upn}/drive/root/children"
    else:
        encoded = requests.utils.quote(folder_path, safe="")
        url = f"{GRAPH_BASE}/users/{upn}/drive/root:{encoded}:/children"

    try:
        items = graph_get(url)
    except requests.exceptions.HTTPError as e:
        print(f"  Warning: Could not list {folder_path}: {e}")
        return

    for item in items:
        # Fix 4: Sanitise folder and file names
        raw_name = item.get("name", "unknown")
        name = safe_filename(raw_name)

        if "folder" in item:
            sub_dir = local_dir / name
            sub_dir.mkdir(exist_ok=True)
            sub_path = folder_path.rstrip("/") + "/" + raw_name
            export_onedrive(upn, out_dir, since_dt, sub_path, sub_dir, depth + 1)

        elif "file" in item:
            if since_dt:
                modified_str = item.get("lastModifiedDateTime", "")
                if modified_str:
                    mod_dt = datetime.fromisoformat(modified_str.replace("Z", "+00:00"))
                    if mod_dt < since_dt:
                        continue

            dest = local_dir / name
            if dest.exists():
                continue  # resume: skip already-downloaded files

            try:
                download_file(upn, item, dest)
            except Exception as e:
                print(f"  Warning: could not download {name}: {e}")


# ── Fix 8: Pre-flight check using --admin ────────────────────────────────────

def preflight_check(admin_upn):
    """
    Verify the app registration and admin consent work before starting
    a potentially long export. Catches 403 Forbidden on second 1, not minute 40.
    """
    print("  Running pre-flight check...")
    try:
        r = _request("GET", f"{GRAPH_BASE}/users/{admin_upn}")
        name = r.json().get("displayName", admin_upn)
        print(f"  ✓ Auth confirmed. Signed in as: {name}\n")
    except requests.exceptions.HTTPError as e:
        if e.response is not None and e.response.status_code == 403:
            raise RuntimeError(
                "Pre-flight failed: 403 Forbidden.\n"
                "Admin consent was not granted. Go to Stage 1, Step 3 and click "
                "'Grant admin consent for [your org]'."
            )
        elif e.response is not None and e.response.status_code == 401:
            raise RuntimeError(
                "Pre-flight failed: 401 Unauthorized.\n"
                "Check TENANT_ID, CLIENT_ID, and CLIENT_SECRET in your .env file."
            )
        raise


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Export M365 employee mailbox and OneDrive files locally."
    )
    parser.add_argument(
        "--admin", required=True,
        help="UPN of the admin account — used for pre-flight auth check"
    )
    parser.add_argument(
        "--user",
        help="Single employee to export"
    )
    parser.add_argument(
        "--only", action="append",
        help="Specific employee to export — repeat for multiple"
    )
    parser.add_argument(
        "--since-days", type=int,
        help="Export only the last N days of data"
    )
    parser.add_argument(
        "--modified-after",
        help="Export files/emails modified or received after YYYY-MM-DD"
    )
    parser.add_argument(
        "--export-only", action="store_true",
        help="(accepted for compatibility; export is always local)"
    )
    parser.add_argument(
        "--local-only", action="store_true",
        help="(accepted for compatibility; data is never uploaded anywhere)"
    )
    args = parser.parse_args()

    users = args.only if args.only else ([args.user] if args.user else [])
    if not users:
        parser.error("Specify --user or at least one --only flag.")

    since_dt = None
    if args.since_days:
        since_dt = datetime.now(timezone.utc) - timedelta(days=args.since_days)
    elif args.modified_after:
        since_dt = datetime.fromisoformat(args.modified_after).replace(tzinfo=timezone.utc)

    print("Authenticating with Microsoft 365...")
    try:
        get_token()
    except RuntimeError as e:
        print(f"Error: {e}")
        sys.exit(1)

    # Fix 8: Pre-flight — fail fast before starting the real export
    try:
        preflight_check(args.admin)
    except RuntimeError as e:
        print(f"\nError: {e}")
        sys.exit(1)

    out_base = Path("out")
    out_base.mkdir(exist_ok=True)

    for upn in users:
        print(f"━━━ {upn} ━━━")
        safe_upn = upn.replace("@", "_at_").replace(".", "_")
        user_dir = out_base / safe_upn
        user_dir.mkdir(exist_ok=True)

        export_emails(upn, user_dir, since_dt)
        export_onedrive(upn, user_dir, since_dt)
        print(f"  ✓ OneDrive export complete.\n")

    print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print("All exports complete. Data is saved in the 'out' folder.")
    print("Nothing has been uploaded anywhere.")


if __name__ == "__main__":
    main()
