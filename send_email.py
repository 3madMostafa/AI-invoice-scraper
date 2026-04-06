#!/usr/bin/env python3
"""
Email Sender for Egyptian eInvoicing Results
=============================================
Sends ONE zip file:

    invoices_YYYY-MM-DD.zip
    ├── results/
    │   ├── GAP/YYYY-MM-DD.xlsx
    │   ├── GPI/YYYY-MM-DD.xlsx
    │   └── GNP/YYYY-MM-DD.xlsx
    └── pdf/
        ├── GAP/YYYY-MM-DD/*.pdf
        ├── GPI/YYYY-MM-DD/*.pdf
        └── GNP/YYYY-MM-DD/*.pdf
"""

import smtplib
import sys
import argparse
import logging
import zipfile
import re

# Fix Windows console encoding
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

from pathlib import Path
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from typing import List, Optional, Tuple

# ============================================================================
# EMAIL CONFIGURATION
# ============================================================================

SMTP_SERVER   = "smtp.gmail.com"
SMTP_PORT     = 587
SMTP_EMAIL    = "autofinancialalerts@gmail.com"
SMTP_PASSWORD = "phjn zdwb htpm lije"

RECIPIENT_EMAILS = [
    # "Payables@globalnapi.com"
    # "Mohamedzenhomsayed@gmail.com",
    "emadmostafa1442002@gmail.com"
]

COMPANIES = ["GAP", "GPI", "GNP"]

# ============================================================================
# LOGGING
# ============================================================================

def setup_logging() -> logging.Logger:
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)

    logger = logging.getLogger('emailer')
    logger.setLevel(logging.INFO)
    for h in logger.handlers[:]:
        logger.removeHandler(h)

    fh = logging.FileHandler(log_dir / 'email.log', encoding='utf-8', mode='a')
    ch = logging.StreamHandler()
    fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    fh.setFormatter(fmt); ch.setFormatter(fmt)
    logger.addHandler(fh); logger.addHandler(ch)
    return logger

logger = setup_logging()

# ============================================================================
# ZIP PACKAGE
# ============================================================================

def find_latest_date() -> Optional[str]:
    """Find the most recent date label across results/<COMPANY>/
    Only considers stems that are pure dates (YYYY-MM-DD), ignoring
    combined files like invoices_YYYY-MM-DD.
    """
    results_path = Path("results")
    if not results_path.exists():
        return None

    DATE_RE = re.compile(r'^\d{4}-\d{2}-\d{2}$')

    dates = set()
    for xlsx in results_path.glob("**/*.xlsx"):
        if DATE_RE.match(xlsx.stem):
            dates.add(xlsx.stem)

    if not dates:
        return None

    latest = sorted(dates, reverse=True)[0]
    logger.info(f"Latest date label: {latest}")
    return latest


def create_zip_package(date_label: Optional[str] = None) -> Optional[Path]:
    """
    Build one ZIP file containing:

        results/<COMPANY>/<date_label>.xlsx   (for each company)
        pdf/<COMPANY>/<date_label>/*.pdf      (for each company)

    Returns the Path to the zip, or None on failure.
    """
    if date_label is None:
        date_label = find_latest_date()
        if date_label is None:
            logger.error("No date label found — cannot build ZIP")
            return None

    output_dir = Path("api_outputs")
    output_dir.mkdir(parents=True, exist_ok=True)
    zip_path = output_dir / f"invoices_{date_label}.zip"

    added = 0
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:

        # ── Excel files ──────────────────────────────────────────────────────
        results_root = Path("results")
        for company in COMPANIES:
            xlsx = results_root / company / f"{date_label}.xlsx"
            if xlsx.exists():
                arcname = f"results/{company}/{date_label}.xlsx"
                zipf.write(xlsx, arcname)
                logger.info(f"Added Excel: {arcname}")
                added += 1
            else:
                logger.warning(f"Excel not found, skipping: {xlsx}")

        # ── General combined CSV ─────────────────────────────────────────────
        general_csv = results_root / f"invoices_{date_label}.csv"
        if general_csv.exists():
            arcname = f"results/invoices_{date_label}.csv"
            zipf.write(general_csv, arcname)
            logger.info(f"Added general CSV: {arcname}")
            added += 1
        else:
            logger.warning(f"General CSV not found, skipping: {general_csv}")

        # ── PDF files ────────────────────────────────────────────────────────
        pdf_root = Path("pdf")
        for company in COMPANIES:
            pdf_dir = pdf_root / company / date_label
            if pdf_dir.exists():
                for pdf_file in sorted(pdf_dir.glob("*.pdf")):
                    arcname = f"pdf/{company}/{date_label}/{pdf_file.name}"
                    zipf.write(pdf_file, arcname)
                    added += 1
                logger.info(f"Added PDFs from {pdf_dir}")
            else:
                logger.warning(f"PDF folder not found, skipping: {pdf_dir}")

    if added == 0:
        logger.error("ZIP created but contains no files")
        zip_path.unlink(missing_ok=True)
        return None

    size_mb = zip_path.stat().st_size / (1024 * 1024)
    logger.info(f"ZIP ready: {zip_path.name} ({size_mb:.2f} MB, {added} entries)")
    return zip_path


# ============================================================================
# EMAIL CONTENT
# ============================================================================

def zip_summary(zip_path: Path) -> Tuple[int, int, float]:
    """Returns (excel_count, pdf_count, size_mb) from a zip file."""
    excel_count = 0
    pdf_count   = 0
    try:
        with zipfile.ZipFile(zip_path, 'r') as z:
            for name in z.namelist():
                if name.endswith('.xlsx'):
                    excel_count += 1
                elif name.endswith('.pdf'):
                    pdf_count += 1
    except Exception:
        pass
    size_mb = zip_path.stat().st_size / (1024 * 1024)
    return excel_count, pdf_count, size_mb


def create_email_content(zip_path: Path) -> Tuple[str, str]:
    excel_count, pdf_count, size_mb = zip_summary(zip_path)

    subject = (
        f"Egyptian eInvoicing Results — {datetime.now().strftime('%d/%m/%Y')} "
        f"({excel_count} Excel, {pdf_count} PDF)"
    )

    html = f"""
    <html><head><style>
        body {{ font-family: 'Segoe UI', sans-serif; color: #333;
                max-width: 800px; margin: auto; padding: 20px; }}
        .header {{ background: linear-gradient(135deg, #667eea, #764ba2);
                   color: white; padding: 30px; border-radius: 10px;
                   text-align: center; margin-bottom: 30px; }}
        .header h1 {{ margin: 0; font-size: 26px; }}
        .header p  {{ margin: 8px 0 0; opacity: 0.9; }}
        .section   {{ background: #f8f9fa; padding: 20px; border-radius: 10px;
                      border-left: 4px solid #667eea; margin-bottom: 20px; }}
        .section h3 {{ margin-top: 0; }}
        .item {{ padding: 10px 15px; margin: 8px 0; background: white;
                 border-radius: 6px; border-left: 3px solid #28a745; font-size: 14px; }}
        .item.zip {{ border-left-color: #667eea; }}
        .tree {{ font-family: monospace; background: #1e1e1e; color: #d4d4d4;
                 padding: 16px; border-radius: 8px; font-size: 13px;
                 line-height: 1.6; white-space: pre; }}
        .footer {{ text-align: center; color: #888; font-size: 12px; margin-top: 20px; }}
    </style></head><body>

    <div class="header">
        <h1>Egyptian eInvoicing Results</h1>
        <p>{datetime.now().strftime('%A, %B %d, %Y at %H:%M')}</p>
    </div>

    <div class="section">
        <p>Hello,</p>
        <p>Please find attached the complete extraction package for all three companies.</p>

        <h3>📦 Attached Package</h3>
        <div class="item zip">
            <strong>{zip_path.name}</strong>
            &nbsp;•&nbsp; {excel_count} Excel files
            &nbsp;•&nbsp; {pdf_count} PDF invoices
            &nbsp;•&nbsp; {size_mb:.2f} MB
        </div>

    <div class="footer">
        <p>Egyptian eInvoicing Data Extraction System — automated email, do not reply.</p>
    </div>
    </body></html>"""

    return subject, html


# ============================================================================
# SEND
# ============================================================================

def attach_file(msg: MIMEMultipart, file_path: Path,
                custom_name: Optional[str] = None) -> bool:
    try:
        with open(file_path, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
        encoders.encode_base64(part)
        name = custom_name or file_path.name
        part.add_header('Content-Disposition', f'attachment; filename="{name}"')
        msg.attach(part)
        return True
    except Exception as e:
        logger.error(f"Failed to attach {file_path.name}: {e}")
        return False


def send_email(zip_path: Path, recipients: List[str]) -> bool:
    try:
        subject, html_body = create_email_content(zip_path)

        msg = MIMEMultipart('mixed')
        msg['From']    = SMTP_EMAIL
        msg['To']      = ", ".join(recipients)
        msg['Subject'] = subject

        alt = MIMEMultipart('alternative')
        alt.attach(MIMEText(html_body, 'html', 'utf-8'))
        msg.attach(alt)

        if not attach_file(msg, zip_path):
            logger.error("Failed to attach ZIP")
            return False

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_EMAIL, SMTP_PASSWORD)
            server.send_message(msg, to_addrs=recipients)

        logger.info(f"Email sent — {zip_path.name} to {', '.join(recipients)}")
        return True

    except Exception as e:
        logger.error(f"Failed to send email: {e}", exc_info=True)
        return False


# ============================================================================
# MAIN
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description='Send eInvoicing results as a single ZIP via email',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python send_email.py                            # auto-build ZIP from results/ + pdf/
  python send_email.py --file invoices.zip        # send a ready-made zip directly
  python send_email.py --date 2026-03-09          # specific date label
  python send_email.py --email someone@x.com     # custom recipient
        """
    )
    parser.add_argument('--file',   type=str,  help='Send a specific zip directly (skip auto-build)')
    parser.add_argument('--date',   type=str,  help='Date label e.g. 2026-03-09')
    parser.add_argument('--email',  type=str,  help='Override recipient email')
    parser.add_argument('--emails', nargs='+', help='Multiple recipient emails')
    args = parser.parse_args()

    recipients = (
        [args.email]  if args.email  else
        args.emails   if args.emails else
        RECIPIENT_EMAILS
    )

    logger.info("Starting email process")
    logger.info(f"Recipients: {', '.join(recipients)}")

    # ── Mode 1: ready-made file passed in ──────────────────────────────────
    if args.file:
        zip_path = Path(args.file)
        if not zip_path.exists():
            logger.error(f"File not found: {args.file}")
            sys.exit(1)

        ok = send_email(zip_path, recipients)

    # ── Mode 2: auto-build ZIP from results/ + pdf/ ─────────────────────────
    else:
        zip_path = create_zip_package(args.date)
        if zip_path is None:
            logger.error("Could not build ZIP — no files found")
            sys.exit(1)

        ok = send_email(zip_path, recipients)

    if ok:
        excel_count, pdf_count, size_mb = zip_summary(zip_path)
        print(f"\n{'='*60}")
        print(f"EMAIL SENT SUCCESSFULLY")
        print(f"{'='*60}")
        print(f"Package:    {zip_path.name}  ({size_mb:.2f} MB)")
        print(f"Excel:      {excel_count} files")
        print(f"PDF:        {pdf_count} files")
        print(f"Recipients: {', '.join(recipients)}")
        print(f"{'='*60}\n")
    else:
        print("\nFailed. Check logs/email.log\n")
        sys.exit(1)


if __name__ == "__main__":
    main()