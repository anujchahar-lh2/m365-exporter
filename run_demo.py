#!/usr/bin/env python3
"""
M365 Exporter — DEMO MODE
Generates realistic fake data so you can see exactly what a real export
produces, without needing any Microsoft 365 credentials.

Usage:
    python3 run_demo.py
    python3 run_demo.py --user alice@lh2.ai --since-days 90
"""

import argparse
import json
import csv
import os
import random
from datetime import datetime, timedelta, timezone
from pathlib import Path

# ── Sample data pool ──────────────────────────────────────────────────────────

PEOPLE = [
    ("Alice Sharma",    "alice.sharma@lh2.ai"),
    ("Rohan Mehta",     "rohan.mehta@lh2holdings.com"),
    ("Priya Nair",      "priya.nair@lh2.ai"),
    ("Yash Choudhary",  "yash.choudhary@lh2holdings.com"),
    ("Sara Iyer",       "sara.iyer@lh2.ai"),
    ("Client Contact",  "contact@client-corp.com"),
    ("Vendor Rep",      "billing@vendor.io"),
]

EMAIL_SUBJECTS = [
    "Re: Q2 Dataset Delivery — LH2 AI Labs",
    "Follow-up: Medical Data Licensing Agreement",
    "Proposal Draft — Doctor-Patient Audio Dataset",
    "Action needed: Sign off on NDA by Friday",
    "Weekly sync notes — Data Catalogue team",
    "Fwd: Invoice #2047 from Annotation Vendor",
    "Re: OB/GYN dataset — scope confirmation",
    "Intro: New partnership with HealthTech Inc.",
    "LH2 Company Catalog — v3 feedback",
    "Reminder: Sprint review tomorrow 3pm",
    "Re: Codebase classifier results — batch 4",
    "External: Data request from Acme Analytics",
    "Quick question about the storefront sheet",
    "Contract renewal — Azure credits",
]

TEAMS_SUBJECTS = [
    "Teams chat: Sara Iyer, Alice Sharma",
    "Teams chat: Rohan Mehta, Yash Choudhary, Alice Sharma",
    "Teams channel: LH2 Data Team — General",
    "Teams channel: LH2 Data Team — Engineering",
    "Teams meeting: Sprint Review — 30 Aug 2024",
]

EMAIL_BODIES = [
    """Hi team,

Attached is the updated Q2 delivery schedule for the medical datasets. Please review
the timeline for the Doctor-Patient Audio section — we may need to push by a week
given the annotation backlog.

Let me know if you have questions.

Best,
Alice""",
    """Hi,

Following up on our call last Thursday — we're ready to proceed with the licensing
agreement for the multimodal clinical dataset. Legal has reviewed the NDA and it
looks good. Can you send the signed copy by EOD Friday?

Thanks,
Rohan""",
    """Team,

Notes from today's sync:
  • Dashboard v3 approved — Yash to share externally by next week
  • OB/GYN dataset still pending confirmation from hospital partner
  • Catalogue sheet: ~13 companies added, 2 more in pipeline
  • Next meeting: Monday 10am

Regards,
Priya""",
    """Hi,

Please find attached Invoice #2047 for annotation work completed in August.
Total: $4,200 USD. Net 30 terms as agreed.

Cheers,
Vendor Rep""",
]

TEAMS_BODIES = [
    """[Teams 1:1 Chat transcript]
Sara Iyer: Hey, did you get a chance to look at the OB/GYN dataset proposal?
Alice Sharma: Yes, reviewing now. A few questions on the scope — can we jump on a call?
Sara Iyer: Sure, free at 3pm?
Alice Sharma: Works for me. Sending invite now.""",
    """[Teams Group Chat transcript]
Rohan Mehta: Team — reminder that the external catalogue goes out to clients on Friday.
Yash Choudhary: Confirmed, v3 is ready. Sharing the sheet now.
Alice Sharma: Looks great! One small change on the About Company description for row 7.""",
    """[Teams Channel: LH2 Data Team — General]
Sara Iyer posted: Dashboard is live. Link in the pinned message.
Priya Nair replied: Looks clean! The navy/pink palette came out really well.
Alice Sharma replied: Agreed. Client should be happy with this.""",
]

ATTACHMENTS = [
    ("Q2_Delivery_Schedule.xlsx",  b"PK\x03\x04DEMO_XLSX"),
    ("NDA_Draft_v2.docx",          b"PK\x03\x04DEMO_DOCX"),
    ("Invoice_2047.pdf",           b"%PDF-1.4 DEMO"),
    ("Dataset_Proposal.pdf",       b"%PDF-1.4 DEMO"),
    ("Batch4_Compiled_Sheet.xlsx", b"PK\x03\x04DEMO_XLSX"),
]

ONEDRIVE = {
    "Documents": {
        "Contracts":     ["NDA_HealthTech_2024.docx", "MSA_Client_Corp_v3.pdf", "Vendor_Agreement_Aug24.docx"],
        "Proposals":     ["LH2_Dataset_Proposal_Q3.pptx", "Pricing_Overview.xlsx"],
        "Meeting Notes": ["Weekly_Sync_Sep2024.docx", "Sprint_Review_Aug2024.docx"],
    },
    "Datasets": {
        "Medical":            ["Doctor_Patient_Audio_Manifest.csv", "Dataset_Metadata_v2.xlsx"],
        "Company Catalogue":  ["External_Company_Catalog_v3.xlsx", "Card_View_Draft.xlsx"],
    },
    "Shared with Team": {
        "Codebase Profiler": ["Batch4_Results.csv", "Org_Pipeline_Summary.xlsx"],
    },
    "Personal": ["Notes_Sep2024.docx", "Action_Items.txt"],
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def rand_date(since, until):
    delta = until - since
    return since + timedelta(seconds=random.randint(0, int(delta.total_seconds())))

def fmt(dt):
    return dt.strftime("%Y-%m-%dT%H:%M:%SZ")

def safe_filename(name, fallback="file", max_len=120):
    name = os.path.basename(str(name or "")).replace("\x00", "")
    name = "".join(c if c.isalnum() or c in " .-_" else "_" for c in name).strip(" .")
    return (name or fallback)[:max_len]

def unique_path(directory, name):
    stem, ext = os.path.splitext(name)
    candidate, n = directory / name, 1
    while candidate.exists():
        candidate = directory / f"{stem} ({n}){ext}"
        n += 1
    return candidate

def print_tree(path, prefix=""):
    entries = sorted(Path(path).iterdir(), key=lambda p: (p.is_file(), p.name))
    for i, entry in enumerate(entries):
        connector = "└── " if i == len(entries) - 1 else "├── "
        print(f"{prefix}{connector}{entry.name}")
        if entry.is_dir():
            ext = "    " if i == len(entries) - 1 else "│   "
            print_tree(entry, prefix + ext)

# ── Generators ────────────────────────────────────────────────────────────────

def generate_emails(user_dir, since, until, count=12):
    print("    Pulling mailbox (Outlook + Teams chats)...")
    email_dir  = user_dir / "dump" / "emails"
    attach_dir = email_dir / "attachments"
    email_dir.mkdir(parents=True, exist_ok=True)
    attach_dir.mkdir(parents=True, exist_ok=True)

    messages = []
    att_count = 0

    for i in range(count):
        sender_name, sender_addr = random.choice(PEOPLE)
        recv_dt  = rand_date(since, until)
        sent_dt  = recv_dt - timedelta(minutes=random.randint(1, 30))
        msg_id   = f"AAMkABC{i:04d}XYZ{random.randint(10000,99999)}"
        is_teams = i >= (count - 3)

        subject  = random.choice(TEAMS_SUBJECTS) if is_teams else random.choice(EMAIL_SUBJECTS)
        body     = random.choice(TEAMS_BODIES)   if is_teams else random.choice(EMAIL_BODIES)
        to_list  = [random.choice(PEOPLE)[1] for _ in range(random.randint(1,2))]
        cc_list  = [random.choice(PEOPLE)[1]] if random.random() > 0.6 else []
        bcc_list = [random.choice(PEOPLE)[1]] if random.random() > 0.8 else []
        has_att  = (not is_teams) and (i % 3 == 0)

        msg = {
            "id":               msg_id,
            "subject":          subject,
            "from":             {"emailAddress": {"name": sender_name, "address": sender_addr}},
            "toRecipients":     [{"emailAddress": {"address": a}} for a in to_list],
            "ccRecipients":     [{"emailAddress": {"address": a}} for a in cc_list],
            "bccRecipients":    [{"emailAddress": {"address": a}} for a in bcc_list],
            "sentDateTime":     fmt(sent_dt),
            "receivedDateTime": fmt(recv_dt),
            "hasAttachments":   has_att,
            "bodyPreview":      body[:120].replace("\n", " "),
            "body":             {"contentType": "text", "content": body},
            "_type":            "Teams chat" if is_teams else "Email",
        }
        messages.append(msg)

        # Individual .txt file (matches real script output)
        sn = safe_filename(msg_id[:40], fallback="msg")
        msg_path = email_dir / f"{sn}.txt"
        with open(msg_path, "w", encoding="utf-8") as f:
            f.write(f"Subject:    {subject}\n")
            f.write(f"From:       {sender_name} <{sender_addr}>\n")
            f.write(f"Sent:       {fmt(sent_dt)}\n")
            f.write(f"Received:   {fmt(recv_dt)}\n")
            f.write(f"To:         {', '.join(to_list)}\n")
            f.write(f"CC:         {', '.join(cc_list)}\n")
            f.write(f"BCC:        {', '.join(bcc_list)}\n")
            f.write(f"Type:       {'Teams chat (journaled to mailbox)' if is_teams else 'Email'}\n")
            f.write("-" * 60 + "\n\n")
            f.write(body)

        if has_att:
            att_name, att_bytes = random.choice(ATTACHMENTS)
            att_name = safe_filename(att_name)
            att_path = unique_path(attach_dir, f"{sn}_{att_name}")
            with open(att_path, "wb") as af:
                af.write(att_bytes)
            att_count += 1

    # emails_all.json (streamed in real script; written at once in demo)
    with open(email_dir / "emails_all.json", "w", encoding="utf-8") as f:
        f.write("[\n")
        for i, m in enumerate(messages):
            if i > 0:
                f.write(",\n")
            json.dump(m, f, indent=2, ensure_ascii=False)
        f.write("\n]\n")

    # emails_metadata.csv (matches real script columns)
    with open(email_dir / "emails_metadata.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["id", "subject", "from_address", "sentDateTime",
                         "receivedDateTime", "hasAttachments"])
        for m in messages:
            writer.writerow([
                m["id"], m["subject"],
                m["from"]["emailAddress"]["address"],
                m["sentDateTime"], m["receivedDateTime"],
                m["hasAttachments"],
            ])

    print(f"    ✓ {count} items ({count-3} emails + 3 Teams chats) | {att_count} attachments")
    return count


def generate_onedrive(user_dir, since, until):
    print("    Pulling OneDrive files...")
    files_dir = user_dir / "dump" / "files"
    count = 0

    def write_folder(structure, parent):
        nonlocal count
        for key, value in structure.items():
            if isinstance(value, dict):
                sub = parent / safe_filename(key)
                sub.mkdir(exist_ok=True)
                write_folder(value, sub)
            elif isinstance(value, list):
                folder = parent / safe_filename(key)
                folder.mkdir(exist_ok=True)
                for filename in value:
                    mod_dt = rand_date(since, until)
                    safe_name = safe_filename(filename)
                    fpath = folder / safe_name
                    ext = filename.split(".")[-1].lower()
                    if ext == "txt":
                        fpath.write_text(f"File: {filename}\nModified: {fmt(mod_dt)}\n\nSample content.", encoding="utf-8")
                    elif ext == "csv":
                        fpath.write_text("id,name,value\n1,Sample,100\n2,Demo,200\n", encoding="utf-8")
                    else:
                        fpath.write_bytes(b"DEMO_FILE: " + filename.encode() + b"\nModified: " + fmt(mod_dt).encode())
                    count += 1

    files_dir.mkdir(parents=True, exist_ok=True)
    write_folder(ONEDRIVE, files_dir)

    for fname in ["Read_Me_First.txt", "Action_Items_Sep2024.txt"]:
        (files_dir / fname).write_text(f"Root file: {fname}\nGenerated by M365 export demo.\n", encoding="utf-8")
        count += 1

    print(f"    ✓ {count} files across {len(ONEDRIVE)} top-level folders")
    return count


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="M365 Exporter — DEMO MODE (no credentials needed)")
    parser.add_argument("--user",           default="employee@yourcompany.com")
    parser.add_argument("--since-days",     type=int, default=90)
    parser.add_argument("--modified-after", help="YYYY-MM-DD")
    # Accept real-script flags so demo and real script are interchangeable
    parser.add_argument("--admin",       default="admin@yourcompany.com")
    parser.add_argument("--only",        action="append")
    parser.add_argument("--export-only", action="store_true")
    parser.add_argument("--local-only",  action="store_true")
    args = parser.parse_args()

    upn = (args.only[0] if args.only else args.user)

    until = datetime.now(timezone.utc)
    if args.modified_after:
        since = datetime.fromisoformat(args.modified_after).replace(tzinfo=timezone.utc)
    else:
        since = until - timedelta(days=args.since_days)

    print("=" * 60)
    print("  M365 Exporter — DEMO MODE")
    print("  (No credentials needed — generating sample data)")
    print("=" * 60)
    print(f"\n  Employee  : {upn}")
    print(f"  Date range: {since.strftime('%d %b %Y')} → {until.strftime('%d %b %Y')}")
    print(f"  Scope     : Outlook + Teams chats + OneDrive files\n")

    out_base = Path("demo_out")
    safe_upn = upn.replace("@", "_at_").replace(".", "_")
    user_dir = out_base / safe_upn
    user_dir.mkdir(parents=True, exist_ok=True)

    print("  [1/3] Pre-flight check...")
    print("        → DEMO MODE: skipping real auth ✓\n")
    print(f"  [2/3] Exporting mailbox for {upn}:")
    email_count = generate_emails(user_dir, since, until)
    print(f"\n  [3/3] Exporting OneDrive for {upn}:")
    file_count  = generate_onedrive(user_dir, since, until)

    print("\n" + "=" * 60)
    print("  EXPORT COMPLETE")
    print("=" * 60)
    print(f"  Emails + Teams chats : {email_count} items")
    print(f"  OneDrive files       : {file_count} files")
    print(f"  Output folder        : demo_out/{safe_upn}/\n")
    print("  Folder structure:")
    print(f"  demo_out/")
    print(f"  └── {safe_upn}/")
    print_tree(user_dir, "      ")
    print()
    print("  Key files:")
    print("  ├─ dump/emails/emails_metadata.csv  → open in Excel")
    print("  ├─ dump/emails/emails_all.json       → full dump")
    print("  ├─ dump/emails/attachments/          → email attachments")
    print("  └─ dump/files/                       → OneDrive (original structure)")
    print()
    print("  Ready to run for real? Set up your .env and run:")
    print("  python3 run_m365_export.py --admin admin@company.com \\")
    print(f"    --user {upn} --export-only --local-only --since-days {args.since_days}")
    print()

if __name__ == "__main__":
    main()
