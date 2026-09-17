# m365-exporter

Export a specific employee's Microsoft 365 data — Outlook emails, OneDrive files, and Teams chats — to your computer, for a date range you choose.

Works on **Business Basic, Business Standard, and all higher M365 plans**. No eDiscovery licence needed.

---

## What it exports

| Data | Format | Notes |
|---|---|---|
| Outlook emails + attachments | JSON + original files | All emails in the date range |
| Teams 1:1 and group chats | JSON | Journaled to the mailbox — included automatically |
| OneDrive files (docs, PPTs, Excel, PDFs) | Native files | Original folder structure preserved |
| Email metadata summary | CSV | Open in Excel — one row per email |

> **Teams channel messages** (public team conversations) are not yet implemented. They require the `ChannelMessage.Read.All` permission and a separate API call. 1:1 chats and group chats are included automatically via the mailbox.

---

## Before you start

- **Microsoft 365 plan:** Business Basic or Standard (or higher)
- **Your role:** Global Administrator on the tenant
- **Python:** 3.10 or newer — [python.org](https://python.org)

---

## Try it first — no credentials needed

Run the demo to see exactly what the output looks like before touching any M365 account:

```bash
python3 run_demo.py --user employee@yourcompany.com --since-days 90
```

---

## Setup — do this once

### Step 1 — Clone the repo

```bash
git clone https://github.com/anujchahar-lh2/m365-exporter.git
cd m365-exporter
```

### Step 2 — Install dependencies

```bash
python3 -m pip install -r requirements.txt
```

### Step 3 — Register an app in Microsoft Azure

1. Go to [portal.azure.com](https://portal.azure.com) → sign in as Global Admin
2. Search **Microsoft Entra ID** → **App registrations** → **+ New registration**
3. Name it `m365-exporter` → click **Register**
4. Copy your **Application (client) ID** and **Directory (tenant) ID** from the Overview page
5. Go to **Certificates & secrets** → **+ New client secret** → copy the **Value** immediately
6. Go to **API permissions** → **+ Add a permission** → **Microsoft Graph** → **Application permissions**
7. Add: `Mail.Read`, `Files.Read.All`, `User.Read.All`
8. Click **Grant admin consent for [your org]** → confirm

### Step 4 — Add your credentials

```bash
cp .env.example .env
```

Open `.env` and fill in your three values:

```
TENANT_ID=your-directory-tenant-id
CLIENT_ID=your-application-client-id
CLIENT_SECRET=your-client-secret-value
```

> ⚠️ Never commit `.env` to Git. It is already in `.gitignore`.

---

## Usage

Replace the email addresses with real ones.

**One employee — everything:**
```bash
python3 run_m365_export.py \
  --admin admin@yourcompany.com \
  --user employee@yourcompany.com \
  --export-only \
  --local-only
```

**One employee — last 90 days:**
```bash
python3 run_m365_export.py \
  --admin admin@yourcompany.com \
  --user employee@yourcompany.com \
  --export-only \
  --local-only \
  --since-days 90
```

**One employee — from a specific date:**
```bash
python3 run_m365_export.py \
  --admin admin@yourcompany.com \
  --user employee@yourcompany.com \
  --export-only \
  --local-only \
  --modified-after 2024-01-01
```

**Multiple employees:**
```bash
python3 run_m365_export.py \
  --admin admin@yourcompany.com \
  --only alice@yourcompany.com \
  --only bob@yourcompany.com \
  --export-only \
  --local-only \
  --since-days 90
```

---

## Output structure

```
out/
└── employee_at_company_com/
    └── dump/
        ├── emails/
        │   ├── emails_all.json         ← full email dump
        │   ├── emails_metadata.csv     ← open in Excel
        │   ├── [message-id].txt        ← one file per email/chat
        │   └── attachments/            ← all email attachments
        └── files/
            └── ...                     ← OneDrive files, original folder structure
```

To hand off the data:

```bash
# Mac / Linux
zip -r export.zip out/

# Windows — right-click the out folder → Send to → Compressed (zipped) folder
```

---

## Flags

| Flag | Description |
|---|---|
| `--admin` | UPN of your admin account — used for pre-flight auth check |
| `--user` | Single employee to export |
| `--only` | Specific employee — repeat for multiple |
| `--since-days N` | Export only the last N days |
| `--modified-after YYYY-MM-DD` | Export from a specific date onward |
| `--export-only` | Accepted for compatibility; export is always local |
| `--local-only` | Accepted for compatibility; data is never uploaded anywhere |

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `401 Unauthorized` | Check TENANT_ID, CLIENT_ID, CLIENT_SECRET in `.env` |
| `403 Forbidden` | Admin consent not granted — redo Step 3, point 8 |
| `python3: command not found` | Install Python from python.org (Windows: tick "Add to PATH") |
| `ModuleNotFoundError` | Run `python3 -m pip install -r requirements.txt` |
| Script stopped midway | Run the same command again — it resumes automatically |
| Employee not found (404) | Check email spelling; account must be active in M365 |
| `out/` folder empty | Date range too narrow — try without `--since-days` |
| Secret expired | Create a new client secret in Azure → Certificates & secrets |

---

## Security notes

- The script is **read-only** — nothing in Microsoft 365 is changed or deleted
- All data stays on your machine — nothing is uploaded anywhere
- After handing off the export, delete the `out/` folder and revoke the client secret in Azure
- This tool is the Microsoft 365 equivalent of [workspace-classifier](https://github.com/data847/workspace-classifier) for Google Workspace

---

## Licence

MIT
