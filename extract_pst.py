"""
Extract emails from an Outlook .pst file using pypff.
Outputs extracted emails as JSON, one file per folder, and a summary CSV.

Usage:
    python extract_pst.py <path_to_pst> [--output-dir <dir>]
"""

import pypff
import json
import csv
import os
import re
import sys
import argparse
import tempfile
from email import message_from_string
from email.utils import parsedate_to_datetime

import io
import shutil
import subprocess

try:
    import pdfplumber as _pdfplumber
except ImportError:
    _pdfplumber = None

try:
    import docx2txt as _docx2txt
except ImportError:
    _docx2txt = None

# Locate antiword — it ships with Git for Windows but lives outside the normal PATH
_ANTIWORD = shutil.which("antiword") or shutil.which("antiword.exe") or (
    r"C:\Program Files\Git\mingw64\bin\antiword.exe"
    if os.path.exists(r"C:\Program Files\Git\mingw64\bin\antiword.exe") else None
)

_EXTRACTABLE_EXTENSIONS = {".pdf", ".doc", ".docx"}


def _extract_attachment_text(data: bytes, filename: str) -> str:
    ext = os.path.splitext(filename)[1].lower()
    if ext not in _EXTRACTABLE_EXTENSIONS:
        return ""
    try:
        if ext == ".pdf":
            if _pdfplumber is None:
                return ""
            with _pdfplumber.open(io.BytesIO(data)) as pdf:
                return "\n".join(p.extract_text() or "" for p in pdf.pages).strip()

        # Write to temp file for tools that need a path
        with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
            tmp.write(data)
            tmp_path = tmp.name
        try:
            if ext == ".docx":
                if _docx2txt is None:
                    return ""
                return (_docx2txt.process(tmp_path) or "").strip()

            if ext == ".doc":
                if _ANTIWORD is None:
                    return ""
                result = subprocess.run(
                    [_ANTIWORD, tmp_path],
                    capture_output=True,
                )
                return result.stdout.decode("utf-8", errors="replace").strip()
        finally:
            os.unlink(tmp_path)
    except Exception as e:
        print(f"  Warning: could not extract text from {filename!r}: {e}")
    return ""


def _parse_mime_body(raw: str) -> tuple:
    """Return (text_body, attachments_list) by parsing raw as a MIME message.

    attachments_list entries: {filename, size, pdf_text?}
    Returns (raw, None) if raw doesn't look like MIME.
    """
    # Detect bare multipart body (no envelope headers): first non-empty line is "--boundary"
    # or preamble followed by "--boundary"
    boundary = None
    for line in raw.splitlines():
        line = line.strip()
        if line.startswith("--") and len(line) > 2 and not line.endswith("--"):
            boundary = line[2:]
            break
        if line and not line.startswith("--"):
            continue  # preamble text — keep scanning

    if boundary:
        # Synthesize envelope so message_from_string can parse it
        envelope = (
            f"MIME-Version: 1.0\r\n"
            f'Content-Type: multipart/mixed; boundary="{boundary}"\r\n'
            f"\r\n"
        )
        raw_to_parse = envelope + raw
    elif "Content-Type:" in raw or "MIME-Version:" in raw:
        raw_to_parse = raw
    else:
        return raw, None

    try:
        msg = message_from_string(raw_to_parse)
    except Exception:
        return raw, None

    text_parts = []
    attachments = []

    for part in msg.walk():
        ct = part.get_content_type()
        cd = part.get("Content-Disposition", "")
        filename = part.get_filename()

        if filename or "attachment" in cd.lower():
            payload = part.get_payload(decode=True)
            if payload is None:
                continue
            entry = {"filename": filename or "unnamed", "size": len(payload)}
            ext = os.path.splitext(filename or "")[1].lower()
            if ext in _EXTRACTABLE_EXTENSIONS:
                entry["attachment_text"] = _extract_attachment_text(payload, filename)
            attachments.append(entry)
        elif ct == "text/plain":
            payload = part.get_payload(decode=True)
            if payload:
                charset = part.get_content_charset() or "utf-8"
                text_parts.append(payload.decode(charset, errors="replace"))

    return "\n".join(text_parts).strip(), attachments or None


def safe_get(fn, default=""):
    try:
        result = fn()
        return result if result is not None else default
    except Exception:
        return default


def parse_header_field(headers, field):
    """Extract a single header field value from raw transport headers."""
    if not headers:
        return ""
    m = re.search(rf"^{field}:\s*(.+)$", headers, re.IGNORECASE | re.MULTILINE)
    return m.group(1).strip() if m else ""


def parse_sender_email(headers):
    if not headers:
        return ""
    m = re.search(r"^From:\s*<?([^>\r\n]+?)>?\s*$", headers, re.IGNORECASE | re.MULTILINE)
    return m.group(1).strip() if m else ""


def extract_message(message):
    subject = safe_get(message.get_subject)
    sender_name = safe_get(message.get_sender_name)

    headers = safe_get(message.get_transport_headers)
    sender_email = parse_sender_email(headers)
    to = parse_header_field(headers, "To")
    cc = parse_header_field(headers, "Cc")

    # Prefer delivery_time (receive), fall back to client_submit_time (send), then Date: header
    delivery_time = ""
    try:
        dt = message.get_delivery_time()
        if dt:
            delivery_time = dt.isoformat()
    except Exception:
        pass

    if not delivery_time:
        try:
            dt = message.get_client_submit_time()
            if dt:
                delivery_time = dt.isoformat()
        except Exception:
            pass

    if not delivery_time:
        try:
            headers = message.get_transport_headers()
            m = re.search(r"^Date:\s*(.+)$", headers, re.IGNORECASE | re.MULTILINE)
            if m:
                delivery_time = parsedate_to_datetime(m.group(1).strip()).isoformat()
        except Exception:
            pass

    # Body — plain text only (html_body is always empty)
    plain_body = safe_get(message.get_plain_text_body)
    if isinstance(plain_body, bytes):
        plain_body = plain_body.decode("utf-8", errors="replace")

    # Parse MIME structure if present — extracts clean text and decodes attachments
    plain_body, mime_attachments = _parse_mime_body(plain_body)

    if mime_attachments is not None:
        attachments = mime_attachments
    else:
        # Fall back to pypff attachment metadata (no content available)
        attachments = []
        try:
            for i in range(message.get_number_of_attachments()):
                att = message.get_attachment(i)
                attachments.append({
                    "filename": safe_get(att.get_name),
                    "size": safe_get(att.get_size, 0),
                })
        except Exception:
            pass

    return {
        "subject": subject,
        "sender_name": sender_name,
        "sender_email": sender_email,
        "to": to,
        "cc": cc,
        "delivery_time": delivery_time,
        "plain_body": plain_body,
        "attachments": attachments,
    }


def walk_folder(folder, folder_path, output_dir, summary_rows):
    """Recursively walk folders and extract messages."""
    folder_name = safe_get(folder.get_name) or "unnamed"
    current_path = os.path.join(folder_path, folder_name)

    # Extract messages in this folder
    num_messages = folder.get_number_of_sub_messages()
    if num_messages > 0:
        messages = []
        for i in range(num_messages):
            try:
                msg = folder.get_sub_message(i)
                data = extract_message(msg)
                messages.append(data)
                summary_rows.append({
                    "folder": current_path,
                    "subject": data["subject"],
                    "sender_name": data["sender_name"],
                    "sender_email": data["sender_email"],
                    "to": data["to"],
                    "cc": data["cc"],
                    "delivery_time": data["delivery_time"],
                    "num_attachments": len(data["attachments"]),
                })
            except Exception as e:
                print(f"  Warning: could not read message {i} in '{current_path}': {e}")

        # Write folder's messages to JSON
        safe_folder_path = current_path.replace("\\", "_").replace("/", "_").lstrip("_")
        out_file = os.path.join(output_dir, f"{safe_folder_path}.json")
        os.makedirs(os.path.dirname(out_file) if os.path.dirname(out_file) else ".", exist_ok=True)
        with open(out_file, "w", encoding="utf-8") as f:
            json.dump(messages, f, indent=2, ensure_ascii=False)
        print(f"  {current_path}: {num_messages} messages -> {out_file}")

    # Recurse into sub-folders
    num_subfolders = folder.get_number_of_sub_folders()
    for i in range(num_subfolders):
        try:
            sub = folder.get_sub_folder(i)
            walk_folder(sub, current_path, output_dir, summary_rows)
        except Exception as e:
            print(f"  Warning: could not read sub-folder {i} of '{current_path}': {e}")


def extract_pst(pst_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)

    pst = pypff.file()
    pst.open(pst_path)

    print(f"Opened: {pst_path}")
    root = pst.get_root_folder()

    summary_rows = []
    num_subfolders = root.get_number_of_sub_folders()
    print(f"Root has {num_subfolders} top-level folder(s)\n")

    for i in range(num_subfolders):
        folder = root.get_sub_folder(i)
        walk_folder(folder, "", output_dir, summary_rows)

    # Write summary CSV
    csv_path = os.path.join(output_dir, "summary.csv")
    if summary_rows:
        fieldnames = ["folder", "subject", "sender_name", "sender_email", "to", "cc", "delivery_time", "num_attachments"]
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(summary_rows)
        print(f"\nSummary CSV: {csv_path} ({len(summary_rows)} emails total)")
    else:
        print("\nNo messages found.")

    pst.close()


def main():
    parser = argparse.ArgumentParser(description="Extract emails from a .pst file using pypff.")
    parser.add_argument("pst_file", help="Path to the .pst file")
    parser.add_argument("--output-dir", default="extracted_emails", help="Output directory (default: extracted_emails)")
    args = parser.parse_args()

    if not os.path.exists(args.pst_file):
        print(f"Error: file not found: {args.pst_file}")
        sys.exit(1)

    extract_pst(args.pst_file, args.output_dir)


if __name__ == "__main__":
    main()
