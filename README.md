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

## Summarizing emails with Claude

`summarize_emails.py` reads the JSON files produced by `extract_pst.py`, sends each one to Claude Haiku 4.5, and collects the results as markdown tables and an Excel workbook.

**Additional requirements**

```
pip install anthropic python-dotenv openpyxl
```

Set `ANTHROPIC_API_KEY` in your environment or in a `.env` file.

**Usage**

```
python summarize_emails.py [--input-dir <dir>] [--output-dir <dir>] [--xlsx <path>] [--overwrite] [--xlsx-only]
```

| Flag | Default | Description |
|---|---|---|
| `--input-dir` | `extracted_emails/` | Directory containing the JSON files from `extract_pst.py` |
| `--output-dir` | `email_summaries/` | Where per-folder markdown summaries are written |
| `--xlsx` | `claude_summary.xlsx` | Path for the aggregated Excel output |
| `--overwrite` | off | Re-process files that already have a markdown summary |
| `--xlsx-only` | off | Skip Claude calls; rebuild the Excel file from existing markdown |

**Output**

For each JSON file the script writes a markdown file to `--output-dir` containing a table with columns: `DateTime`, `From`, `To`, `Subject`, and `Summary`. The summary combines the email body and any extracted attachment text into bullet points. All per-folder tables are then merged into a single Excel workbook (`claude_summary.xlsx`) with an added `Folder` column.

## How it works

Each JSON file is sent to Claude Haiku 4.5 with a fixed system prompt that instructs the model to produce a markdown table. Files larger than ~120,000 characters are split into chunks to stay within the Haiku 4.5 rate limits; the resulting tables are concatenated into a single markdown file. The system prompt is marked ephemeral so it is cached across chunks and files, reducing token costs. The SDK is configured with up to 10 retries and exponential back-off to ride out rate-limit windows automatically.
