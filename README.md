# ExtractOutlook

Extracts emails from an Outlook `.pst` file into JSON files (one per folder) and a summary CSV.

## Requirements

**Python packages**

```
pip install pypff pdfplumber docx2txt
```

**System tools**

- `antiword` — extracts text from old Word `.doc` files. On Windows, install [Git for Windows](https://git-scm.com/download/win); antiword is bundled at `C:\Program Files\Git\mingw64\bin\antiword.exe` and is found automatically.

## Usage

```
python extract_pst.py <path_to_pst> [--output-dir <dir>]
```

Default output directory is `extracted_emails/`.

## Output

### Per-folder JSON files

Each folder in the PST produces one JSON file containing an array of email objects:

| Field | Description |
|---|---|
| `subject` | Email subject |
| `sender_name` | Display name of sender |
| `sender_email` | Email address of sender |
| `to` | To header |
| `cc` | Cc header |
| `delivery_time` | Received time (ISO 8601) |
| `plain_body` | Plain-text body (MIME structure stripped) |
| `attachments` | Array of attachment objects (see below) |

Each attachment object:

| Field | Description |
|---|---|
| `filename` | Original filename |
| `size` | Size in bytes |
| `attachment_text` | Extracted text (`.doc`, `.docx`, `.pdf` only) |

### summary.csv

One row per email with columns: `folder`, `subject`, `sender_name`, `sender_email`, `to`, `cc`, `delivery_time`, `num_attachments`.

## How it works

`pypff` reads the PST file and walks its folder tree. Some emails store attachments as base64-encoded MIME parts inside the plain-text body; the script detects and parses this structure, decoding attachments directly. Text is extracted from attachments using:

- `.doc` — `antiword` (subprocess)
- `.docx` — `docx2txt` (Python)
- `.pdf` — `pdfplumber` (Python)
