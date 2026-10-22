# AE Techron PST to SharePoint list import

Read Larry Shank's historical PST, keep the fields needed for later classification, and create one SharePoint list item per message.

This is the runnable form of the original sketch. Column internal names, the site, and the list still have to be set for the AE Techron tenant. Attachment bytes are not uploaded, and the AI columns are left blank.

## What you configure

| Concern | Where |
| --- | --- |
| PST path, site id, list id | `config.json` or the environment |
| Sign-in | `auth_mode`: `device_code`, `client_credentials`, or `access_token` |
| Column internal names | `field_map` |
| Text length | `limits` |
| Rows already imported | `state/imported.jsonl`, plus a scan of the list when `sync_existing` is true |
| Progress and failures | `logs/import.log` and the console |

`device_code` fits a one-time import: you sign in once and the refresh token stays in `state/token_cache.json`. `client_credentials` is the unattended option. `access_token` is only useful for a short pilot, because a Graph token expires while a full archive is still running.

Environment variables and a `.env` file next to the config override `config.json`. A variable already set in the shell wins over `.env`.

## Setup

Python 3.10 or newer.

```bash
cd pst-sharepoint-import
python -m pip install -r requirements.txt
copy config.example.json config.json
copy .env.example .env
```

On Windows, `copy` is fine. On macOS or Linux, use `cp`.

`pypff` is separate from `requirements.txt`. Install [libpff](https://github.com/libyal/libpff) and its Python bindings on the machine that has the PST, then check:

```bash
python -c "import pypff"
```

The default archive path is `C:\EmailArchive\Larry_Shank.pst`.

### App registration

Create an app registration in the AE Techron tenant.

Device code (interactive):

- Supported account type: this tenant
- Allow public client flows
- Delegated permission `Sites.ReadWrite.All`, with admin consent
- `auth_mode` = `device_code`
- Set `GRAPH_TENANT_ID` and `GRAPH_CLIENT_ID`

Client credentials (unattended):

- Application permission `Sites.ReadWrite.All` or `Sites.Selected` granted on this site
- Admin consent
- `auth_mode` = `client_credentials`
- Set `GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID`, and `GRAPH_CLIENT_SECRET`

Look up the ids in [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer):

```text
GET https://graph.microsoft.com/v1.0/sites/YOUR_TENANT.sharepoint.com:/sites/YOUR_SITE
GET https://graph.microsoft.com/v1.0/sites/{site-id}/lists
```

Put `site_id` and `list_id` in `.env`. Create a blank list first if you do not have one yet.

### Columns

`field_map` uses SharePoint **internal** names. `Title` is already on every list. Run this once to create the rest:

```bash
python -m pst_sharepoint_import --config config.json --ensure-columns --limit 1
```

That creates any missing columns and imports one message so you can confirm a row lands in the list.

Set a mapped name to `""` when that column should not be created or written. A real import stops if a mapped column is missing.

| Internal name | Type | Contents |
| --- | --- | --- |
| Title | text | Subject, or `(no subject)` |
| EmailID | indexed text | Stable id used to skip duplicates |
| InternetMessageId | text | RFC Message-ID when the PST has one |
| PstIdentifier | text | libpff item id, for a later attachment pass |
| EmailDate | date/time | Submit time, otherwise delivery or creation time, in UTC |
| Sender | text | SMTP address from the From header |
| SenderName | text | Display name |
| Recipients | multiline | To, Cc, and Bcc |
| Folder | multiline | Folder path inside the PST |
| EmailBody | multiline | Plain text. HTML is converted. RTF is marked and left unconverted |
| BodySource | text | `plain`, `html`, `rtf`, or `empty` |
| AttachmentNames | multiline | File names only |
| HasAttachments | yes/no | |
| SourceFile | text | PST file name |
| TechnicalRelevant | yes/no | Reserved. Written as false |
| TechnicalScore | number | Reserved. Written as 0 |
| Customer, Product, Application | text | Reserved. Written blank |
| Problem, Resolution, Keywords | multiline | Reserved. Written blank |
| AIProcessed | yes/no | Reserved. Written as false |

Sender and recipients come from the message transport headers. libpff does not expose `sender_email_address` or `display_to`.

A blank date is omitted so a DateTime column does not reject the row. Single-line values are capped (Title at 255). The body cap defaults to 60,000 characters and the stored text ends with `[truncated]` when it is cut.

## Run

Preview the PST without SharePoint:

```bash
python -m pst_sharepoint_import --config config.json --dry-run --no-sync-existing --limit 10
```

Pilot against the list:

```bash
python -m pst_sharepoint_import --config config.json --ensure-columns --limit 10
```

Full archive:

```bash
python -m pst_sharepoint_import --config config.json
```

Optional `folder_prefixes` limits the walk, for example `["Inbox"]`. A prefix matches a path segment, so `Inbox` includes `Inbox/Customers` and leaves `InboxArchive` alone.

The run prints a sign-in message for device code, then logs each folder. Every 100 messages it logs a progress line. Failures include the folder, subject, and Graph error, and the body is not written to the log.

## Duplicates and resume

Each message gets a SHA-256 id from the PST file name, folder, Message-ID, libpff identifier, subject, sender, date, and a fingerprint of the body. The same message in two folders stays two rows, because the folder is part of the id.

Before creating a row, the importer skips ids already in `state/imported.jsonl`. With `sync_existing` true (the example default), it also pages the list and records `EmailID` values that are already there. Turn that off only for a local preview.

A message that fails is not recorded, so the next run retries it. If the process dies after Graph accepts a row and before the state file is updated, the list scan picks that id up on the next run.

`state/imported.lock` stops two imports at once. Delete that file only after you have confirmed the other process is gone. `state/token_cache.json` is a secret. Do not commit `state/`, `.env`, `config.json`, or the PST.

Graph `429` and other transient status codes wait for `Retry-After` (or a short backoff) and then retry. Retries stop at `max_retries`. For client credentials and device code, a `401` refreshes the token once. A pasted access token cannot be refreshed.

## Tests

```bash
cd pst-sharepoint-import
python -m pip install -r requirements.txt
python -m unittest discover -s tests
```

The tests use fake PST messages and a fake Graph client. They do not read `Larry_Shank.pst` and they do not call Microsoft.

## Still left for a production pass

- Confirm the live internal names if the list was created by hand in SharePoint
- Decide whether `Sites.Selected` should replace `Sites.ReadWrite.All`
- Upload attachment bytes, or pull them later with `PstIdentifier`
- Fill TechnicalRelevant, Customer, Product, and the other AI columns in a separate job
- Raise `limits.body` if the EmailBody column accepts more text, or lower it if Graph rejects a row
