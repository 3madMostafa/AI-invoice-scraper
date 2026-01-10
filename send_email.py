#!/usr/bin/env python3
import smtplib
import os
import sys
import argparse
import logging
from pathlib import Path
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import zipfile

# Email Configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_EMAIL = "autofinancialalerts@gmail.com"
SMTP_PASSWORD = "phjn zdwb htpm lije"

# Multiple recipients
RECIPIENT_EMAILS = [
    # "pola_reffat@globalnapi.com",
    # "doaa_mohamed@globalnapi.com", 
    # "hussein.alshreef@ifssgroup.com",
    # "Mohamedzenhomsayed@gmail.com",
    "emadmostafa1442002@gmail.com"
]

# Setup centralized logging with proper UTF-8 encoding
def setup_logging():
    log_dir = os.path.join(os.getcwd(), "logs")
    os.makedirs(log_dir, exist_ok=True)
    
    logger = logging.getLogger('emailer')
    logger.setLevel(logging.INFO)
    
    # Remove existing handlers
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    
    # File handler with UTF-8 encoding
    file_handler = logging.FileHandler(
        os.path.join(log_dir, 'email.log'), 
        encoding='utf-8',
        mode='a'
    )
    file_handler.setLevel(logging.INFO)
    
    # Console handler with UTF-8 support  
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # Enhanced formatter
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

logger = setup_logging()

def find_latest_output_folder():
    """Find the most recent output folder with new structure: outputs/<date>/Excel/<SupplierName>/"""
    outputs_path = Path("outputs")
    
    if not outputs_path.exists():
        logger.error("Outputs directory does not exist")
        return None
    
    # Find the most recent date folder
    date_folders = [d for d in outputs_path.iterdir() if d.is_dir()]
    if not date_folders:
        logger.error("No date folders found in outputs directory")
        return None
    
    # Sort by name (assuming dd-mm-yyyy format)
    try:
        date_folders.sort(key=lambda x: datetime.strptime(x.name, "%d-%m-%Y"), reverse=True)
        latest_date_folder = date_folders[0]
        logger.info(f"Found latest date folder: {latest_date_folder}")
        
        # NEW STRUCTURE: Look for Excel subfolder
        excel_folder = latest_date_folder / "Excel"
        if not excel_folder.exists():
            logger.error(f"Excel folder not found in {latest_date_folder}")
            return None
        
        # Find supplier folders within Excel folder
        supplier_folders = [d for d in excel_folder.iterdir() if d.is_dir()]
        if not supplier_folders:
            logger.error(f"No supplier folders found in {excel_folder}")
            return None
        
        logger.info(f"Found {len(supplier_folders)} supplier folders in Excel directory")
        return latest_date_folder, supplier_folders
        
    except ValueError as e:
        logger.error(f"Error parsing date folders: {e}")
        return None

def find_results_files(search_path):
    """Find all results.xlsx files in the given path (supports new structure)"""
    results_files = []
    
    if isinstance(search_path, tuple):
        # If we got a tuple from find_latest_output_folder
        date_folder, supplier_folders = search_path
        logger.info(f"Searching for results.xlsx in {len(supplier_folders)} supplier folders")
        
        for supplier_folder in supplier_folders:
            results_file = supplier_folder / "results.xlsx"
            if results_file.exists():
                results_files.append(results_file)
                logger.info(f"Found results file: {results_file}")
            else:
                logger.warning(f"No results.xlsx found in {supplier_folder}")
    else:
        # Direct path provided
        search_path = Path(search_path)
        
        if search_path.is_file() and search_path.name == "results.xlsx":
            results_files.append(search_path)
            logger.info(f"Using direct file: {search_path}")
        elif search_path.is_dir():
            # Check if this is a date folder with Excel subfolder
            excel_folder = search_path / "Excel"
            if excel_folder.exists() and excel_folder.is_dir():
                logger.info(f"Found Excel folder: {excel_folder}")
                # Look in Excel/<SupplierName>/ subdirectories
                for supplier_folder in excel_folder.iterdir():
                    if supplier_folder.is_dir():
                        results_file = supplier_folder / "results.xlsx"
                        if results_file.exists():
                            results_files.append(results_file)
                            logger.info(f"Found results file: {results_file}")
            else:
                # Old structure or single folder - look for results.xlsx directly
                results_file = search_path / "results.xlsx"
                if results_file.exists():
                    results_files.append(results_file)
                else:
                    # Look for results.xlsx in subdirectories
                    for subdir in search_path.iterdir():
                        if subdir.is_dir():
                            results_file = subdir / "results.xlsx"
                            if results_file.exists():
                                results_files.append(results_file)
    
    logger.info(f"Total results files found: {len(results_files)}")
    return results_files

def create_pdf_zip(output_folder_path, date_str):
    """Create a ZIP file containing all PDFs from all suppliers"""
    try:
        # Determine the base path
        if isinstance(output_folder_path, tuple):
            date_folder, _ = output_folder_path
            base_path = date_folder
        else:
            base_path = Path(output_folder_path)
        
        # Look for PDF folder
        pdf_folder = base_path / "PDF"
        
        if not pdf_folder.exists():
            logger.warning(f"PDF folder not found at {pdf_folder}")
            return None
        
        # Create zip file name
        clean_date = date_str.replace("/", "-").replace("\\", "-").replace(":", "-")
        zip_filename = f"_PDFs.zip"
        zip_path = base_path / zip_filename
        
        # Create ZIP file
        pdf_count = 0
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Iterate through supplier folders in PDF directory
            for supplier_folder in pdf_folder.iterdir():
                if supplier_folder.is_dir():
                    supplier_name = supplier_folder.name
                    # Add all PDFs from this supplier
                    for pdf_file in supplier_folder.glob("*.pdf"):
                        # Create archive name: SupplierName/filename.pdf
                        arcname = f"{supplier_name}/{pdf_file.name}"
                        zipf.write(pdf_file, arcname)
                        pdf_count += 1
                        logger.info(f"Added to ZIP: {arcname}")
        
        if pdf_count > 0:
            zip_size = zip_path.stat().st_size / (1024 * 1024)  # Size in MB
            logger.info(f"Created ZIP file: {zip_path} ({zip_size:.2f} MB, {pdf_count} PDFs)")
            return zip_path
        else:
            logger.warning("No PDF files found to zip")
            zip_path.unlink()  # Delete empty zip
            return None
            
    except Exception as e:
        logger.error(f"Error creating PDF ZIP: {e}")
        return None

def attach_file_to_email(msg, file_path, custom_filename=None, date_str=None):
    """Attach a single file to the email message"""
    try:
        file_path = Path(file_path)
        
        if not file_path.exists():
            logger.error(f"File does not exist: {file_path}")
            return False, None
        
        # Determine MIME type based on file extension
        file_extension = file_path.suffix.lower()
        
        # Read file and get content
        with open(file_path, "rb") as attachment:
            file_content = attachment.read()
        
        # Create appropriate MIME object based on file type
        if file_extension in ['.xlsx', '.xls']:
            part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        elif file_extension == '.pdf':
            part = MIMEBase('application', 'pdf')
        elif file_extension in ['.txt', '.log']:
            part = MIMEBase('text', 'plain')
        elif file_extension in ['.jpg', '.jpeg']:
            part = MIMEBase('image', 'jpeg')
        elif file_extension == '.png':
            part = MIMEBase('image', 'png')
        else:
            part = MIMEBase('application', 'octet-stream')
        
        # Set payload
        part.set_payload(file_content)
        
        # Encode file in ASCII characters to send by email    
        encoders.encode_base64(part)
        
        # Create filename with date and supplier information
        if custom_filename:
            filename = custom_filename
        else:
            # Extract supplier name from parent directory
            # NEW STRUCTURE: file_path is outputs/<date>/Excel/<SupplierName>/results.xlsx
            supplier_name = file_path.parent.name
            
            # Get file extension
            file_ext = file_path.suffix
            base_name = file_path.stem
            
            # Create formatted filename: Date_SupplierName_OriginalName.ext
            if date_str and date_str != "Unknown Date":
                filename = f"{date_str}_{supplier_name}_{base_name}{file_ext}"
            else:
                filename = f"{supplier_name}_{base_name}{file_ext}"
        
        # Clean filename (remove invalid characters for email attachments)
        filename = filename.replace("/", "-").replace("\\", "-").replace(":", "-").replace(" ", "_")
        
        # Handle Arabic text properly
        import unicodedata
        
        # Convert Arabic characters to transliteration or remove them
        try:
            filename_ascii = ""
            for char in filename:
                if ord(char) < 128:  # ASCII characters
                    filename_ascii += char
                elif char in 'ابتثجحخدذرزسشصضطظعغفقكلمنهوي':
                    arabic_to_latin = {
                        'ا': 'a', 'ب': 'b', 'ت': 't', 'ث': 'th', 'ج': 'j', 
                        'ح': 'h', 'خ': 'kh', 'د': 'd', 'ذ': 'dh', 'ر': 'r',
                        'ز': 'z', 'س': 's', 'ش': 'sh', 'ص': 's', 'ض': 'd',
                        'ط': 't', 'ظ': 'z', 'ع': 'a', 'غ': 'gh', 'ف': 'f',
                        'ق': 'q', 'ك': 'k', 'ل': 'l', 'م': 'm', 'ن': 'n',
                        'ه': 'h', 'و': 'w', 'ي': 'y'
                    }
                    filename_ascii += arabic_to_latin.get(char, '_')
                else:
                    filename_ascii += '_'
            
            filename = filename_ascii
            
        except:
            # Fallback: use only ASCII characters
            filename = ''.join(c if ord(c) < 128 else '_' for c in filename)
        
        # Ensure filename doesn't have problematic characters
        import re
        filename = re.sub(r'[^\w\-_\.]', '_', filename)
        
        # Remove multiple consecutive underscores
        filename = re.sub(r'_+', '_', filename)
        
        logger.info(f"Final attachment filename: {filename}")
        
        # Add headers
        part.add_header(
            'Content-Disposition',
            f'attachment; filename="{filename}"'
        )
        
        if file_extension in ['.xlsx', '.xls']:
            part.replace_header('Content-Type', f'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; name="{filename}"')
        
        part.set_param('name', filename)
        
        msg.attach(part)
        logger.info(f"Successfully attached file: {filename} ({file_path.stat().st_size} bytes)")
        return True, filename
        
    except Exception as e:
        logger.error(f"Error attaching file {file_path}: {e}")
        return False, None

def attach_multiple_files(msg, file_paths, date_str=None):
    """Attach multiple files to the email message"""
    attached_count = 0
    attached_filenames = []
    
    for file_path in file_paths:
        file_path = Path(file_path)
        
        # Extract supplier name (parent directory in new structure)
        supplier_name = file_path.parent.name
        file_ext = file_path.suffix
        base_name = file_path.stem
        
        # Create formatted filename: Date_SupplierName_OriginalName.ext
        if date_str and date_str != "Unknown Date":
            clean_date = date_str.replace("/", "-").replace("\\", "-").replace(":", "-")
            custom_filename = f"{clean_date}_{supplier_name}_{base_name}{file_ext}"
        else:
            custom_filename = f"{supplier_name}_{base_name}{file_ext}"
        
        # Clean filename further
        custom_filename = custom_filename.replace(" ", "_").replace("(", "").replace(")", "")
        
        result = attach_file_to_email(msg, file_path, custom_filename, date_str)
        if isinstance(result, tuple) and result[0]:
            attached_count += 1
            attached_filenames.append(result[1])
        elif result is True:
            attached_count += 1
            attached_filenames.append(custom_filename)
    
    logger.info(f"Successfully attached {attached_count} out of {len(file_paths)} files")
    return attached_count, attached_filenames

def create_email_content(results_files, output_folder_path, attached_filenames=None):
    """Create email subject and body content"""
    # Determine date from folder path
    date_str = "Unknown Date"
    if isinstance(output_folder_path, Path):
        # NEW: Get date from outputs/<date>/ structure
        date_str = output_folder_path.name
    elif isinstance(output_folder_path, tuple):
        date_folder, _ = output_folder_path
        date_str = date_folder.name
    
    subject = f"Invoice Processing Report - {date_str}"
    
    # Count suppliers and files
    supplier_count = len(results_files)
    total_records = 0
    
    # Try to count records from Excel files
    try:
        import pandas as pd
        for file in results_files:
            try:
                df = pd.read_excel(file)
                total_records += len(df)
            except:
                pass
    except ImportError:
        logger.warning("pandas not available for record counting")
    
    # Create HTML body
    html_body = f"""
    <html>
    <body>
    <h2>Invoice Processing Pipeline - Completion Report</h2>
    
    <p><strong>Report Generated:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    <li><strong>PDF Archive:</strong> All supplier PDFs in one ZIP file</li>
    <p><strong>Processing Date:</strong> {date_str}</p>
    
    <h3>🎉 Pipeline Status: COMPLETED SUCCESSFULLY</h3>
    
    <p>The invoice processing pipeline has completed successfully. All stages (scraping, parsing, and processing) have been executed.</p>
    
    <h3>📊 Summary</h3>
    <ul>
    <li><strong>Suppliers Processed:</strong> {supplier_count}</li>
    <li><strong>Excel Reports Generated:</strong> {len(results_files)}</li>
    """
    
    if total_records > 0:
        html_body += f"<li><strong>Total Invoice Records:</strong> {total_records}</li>"
    
    html_body += """
    </ul>
    
    <h3>📁 Output Location</h3>
    """
    
    # Add output folder path information
    if isinstance(output_folder_path, Path):
        absolute_path = output_folder_path.absolute()
        html_body += f"""
        <p><strong>Output Folder:</strong> <code>{absolute_path}</code></p>
        <p><strong>Excel Files:</strong> <code>{absolute_path}/Excel/&lt;SupplierName&gt;/results.xlsx</code></p>
        <p><strong>PDF Files:</strong> <code>{absolute_path}/PDF/&lt;SupplierName&gt;/*.pdf</code></p>
        """
    elif isinstance(output_folder_path, tuple):
        date_folder, _ = output_folder_path
        absolute_path = date_folder.absolute()
        html_body += f"""
        <p><strong>Output Folder:</strong> <code>{absolute_path}</code></p>
        <p><strong>Excel Files:</strong> <code>{absolute_path}/Excel/&lt;SupplierName&gt;/results.xlsx</code></p>
        <p><strong>PDF Files:</strong> <code>{absolute_path}/PDF/&lt;SupplierName&gt;/*.pdf</code></p>
        """
    
    html_body += """
    <h3>📋 Attached Files</h3>
    <p>The following Excel reports are attached to this email:</p>
    <ul>
    """
    
    # List attached files
    if attached_filenames:
        for i, filename in enumerate(attached_filenames):
            if i < len(results_files):
                file_path = results_files[i]
                supplier_name = file_path.parent.name
                file_size = file_path.stat().st_size / 1024  # Size in KB
                html_body += f"<li><strong>{filename}</strong> - {supplier_name} ({file_size:.1f} KB)</li>"
            else:
                html_body += f"<li><strong>{filename}</strong></li>"
    else:
        # Fallback to original format
        for file in results_files:
            supplier_name = file.parent.name
            file_size = file.stat().st_size / 1024
            clean_date = date_str.replace("/", "-").replace("\\", "-").replace(":", "-")
            if date_str != "Unknown Date":
                display_filename = f"{clean_date}_{supplier_name}_results.xlsx"
            else:
                display_filename = f"{supplier_name}_results.xlsx"
            html_body += f"<li><strong>{display_filename}</strong> - {supplier_name} ({file_size:.1f} KB)</li>"
    
    html_body += """
    </ul>
    
    <h3>📝 File Naming Convention</h3>
    <p>Attached files follow the naming format: <code>Date_SupplierName_results.xlsx</code></p>
    <p>This makes it easy to identify the processing date and supplier for each file.</p>
    
    <h3>📝 Next Steps</h3>
    <p>Please review the attached Excel files for detailed invoice information. Each file contains processed data for a specific supplier including:</p>
    <ul>
    <li>Invoice IDs and dates</li>
    <li>Total values</li>
    <li>Issuer information (with Arabic text segmentation)</li>
    <li>Purchase Order numbers (automatically extracted)</li>
    <li>Processing status</li>
    </ul>
    
    <hr>
    <p><em>This is an automated report generated by the invoice processing system.</em></p>
    <p><em>For any questions or issues, please check the system logs in the 'logs' folder.</em></p>
    
    </body>
    </html>
    """
    
    return subject, html_body

def send_email_with_attachments(results_files, output_folder_path, recipient_emails=None):
    """Send email with Excel attachments to multiple recipients"""
    # استخدام الإيميلات المُدخلة أو الافتراضية
    if recipient_emails is None:
        recipient_emails = RECIPIENT_EMAILS
    
    try:
        # Determine date string for filename formatting
        date_str = "Unknown Date"
        if isinstance(output_folder_path, Path):
            date_str = output_folder_path.name
        elif isinstance(output_folder_path, tuple):
            date_folder, _ = output_folder_path
            date_str = date_folder.name
        
        # Create email content
        subject, html_body = create_email_content(results_files, output_folder_path)
        
        # Create message
        msg = MIMEMultipart('mixed')
        msg['From'] = SMTP_EMAIL
        msg['To'] = ", ".join(recipient_emails)
        msg['Subject'] = subject
        
        # Create alternative container for HTML content
        msg_alternative = MIMEMultipart('alternative')
        
        # Attach HTML content
        html_part = MIMEText(html_body, 'html', 'utf-8')
        msg_alternative.attach(html_part)
        
        # Attach the alternative container to main message
        msg.attach(msg_alternative)
        
        # Attach files
        attached_filenames = []
        
        # Create and attach PDF ZIP file
        pdf_zip_path = create_pdf_zip(output_folder_path, date_str)
        if pdf_zip_path:
            # Create custom filename for ZIP with date
            zip_custom_name = f"{date_str.replace('/', '-').replace('\\', '-').replace(':', '-')}_PDFs.zip"
            result = attach_file_to_email(msg, pdf_zip_path, zip_custom_name, date_str)
            if result[0]:
                attached_filenames.append(result[1])
                logger.info(f"Successfully attached PDF ZIP: {result[1]}")
        
        for file_path in results_files:
            file_path = Path(file_path)
            
            # Extract supplier name
            supplier_name = file_path.parent.name
            file_ext = file_path.suffix
            base_name = file_path.stem
            
            # Create formatted filename
            if date_str and date_str != "Unknown Date":
                clean_date = date_str.replace("/", "-").replace("\\", "-").replace(":", "-")
                custom_filename = f"{clean_date}_{supplier_name}_{base_name}{file_ext}"
            else:
                custom_filename = f"{supplier_name}_{base_name}{file_ext}"
            
            custom_filename = custom_filename.replace(" ", "_").replace("(", "").replace(")", "")
            
            result = attach_file_to_email(msg, file_path, custom_filename, date_str)
            if result[0]:
                attached_filenames.append(result[1])
                logger.info(f"Successfully attached: {result[1]}")
        
        # Send email
        logger.info("Connecting to SMTP server...")
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_EMAIL, SMTP_PASSWORD)
            server.send_message(msg, to_addrs=recipient_emails)
            
        logger.info(f"Email sent successfully to {len(RECIPIENT_EMAILS)} recipients")
        logger.info(f"Recipients: {', '.join(RECIPIENT_EMAILS)}")
        logger.info(f"Subject: {subject}")
        logger.info(f"Total attachments sent: {len(attached_filenames)} files")
        
        return True
        
    except Exception as e:
        logger.error(f"Failed to send email: {e}")
        return False

def main():
    """Main execution function"""
    parser = argparse.ArgumentParser(description='Send email with invoice processing results')
    parser.add_argument('--path', type=str, help='Path to output folder or specific results.xlsx file')
    parser.add_argument('--date', type=str, help='Specific date folder to process (dd-mm-yyyy format)')
    parser.add_argument('--files', nargs='+', help='Specific files to attach')
    parser.add_argument('--email', type=str, help='Recipient email address')
    
    args = parser.parse_args()
    
    # استخدام الإيميل المُدخل أو الافتراضي
    recipient_emails = [args.email] if args.email else RECIPIENT_EMAILS
    
    logger.info("Starting email sending process")
    logger.info(f"Will send to {len(recipient_emails)} recipients: {', '.join(recipient_emails)}")
    
    output_folder_path = None
    results_files = []
    
    try:
        if args.files:
            logger.info(f"Using specific files: {args.files}")
            for file_path in args.files:
                file_path = Path(file_path)
                if file_path.exists():
                    results_files.append(file_path)
                    logger.info(f"Added file: {file_path}")
                else:
                    logger.error(f"File does not exist: {file_path}")
            output_folder_path = Path.cwd()
            
        elif args.path:
            logger.info(f"Using provided path: {args.path}")
            output_folder_path = Path(args.path)
            if not output_folder_path.exists():
                logger.error(f"Provided path does not exist: {args.path}")
                sys.exit(1)
            results_files = find_results_files(output_folder_path)
            
        elif args.date:
            logger.info(f"Looking for date folder: {args.date}")
            date_folder = Path("outputs") / args.date
            if not date_folder.exists():
                logger.error(f"Date folder does not exist: {date_folder}")
                sys.exit(1)
            output_folder_path = date_folder
            results_files = find_results_files(date_folder)
            
        else:
            logger.info("Auto-detecting latest output folder...")
            output_folder_path = find_latest_output_folder()
            if output_folder_path is None:
                logger.error("Could not find output folder")
                sys.exit(1)
            results_files = find_results_files(output_folder_path)
        
        if not results_files:
            logger.error("No files found to attach")
            sys.exit(1)
        
        logger.info(f"Found {len(results_files)} files to attach")
        for file in results_files:
            file_size = file.stat().st_size / 1024
            logger.info(f"  - {file.name} ({file_size:.1f} KB)")
        
        # Send email
        if send_email_with_attachments(results_files, output_folder_path, recipient_emails):
            print(f"\nEmail sent successfully!")
            print(f"Recipients ({len(RECIPIENT_EMAILS)}):")
            for email in RECIPIENT_EMAILS:
                print(f"   • {email}")
            print(f"Attachments: {len(results_files)} files")
            
            # Show attachment names
            date_str = "Unknown Date"
            if isinstance(output_folder_path, Path):
                date_str = output_folder_path.name
            elif isinstance(output_folder_path, tuple):
                date_folder, _ = output_folder_path
                date_str = date_folder.name
            
            clean_date = date_str.replace("/", "-").replace("\\", "-").replace(":", "-")
            
            for file in results_files:
                supplier_name = file.parent.name
                file_size = file.stat().st_size / 1024
                if date_str != "Unknown Date":
                    attachment_name = f"{clean_date}_{supplier_name}_results.xlsx"
                else:
                    attachment_name = f"{supplier_name}_results.xlsx"
                print(f"   • {attachment_name} ({file_size:.1f} KB)")
                
            if isinstance(output_folder_path, Path):
                print(f"Output folder: {output_folder_path.absolute()}")
            elif isinstance(output_folder_path, tuple):
                date_folder, _ = output_folder_path
                print(f"Output folder: {date_folder.absolute()}")
        else:
            print(f"\nFailed to send email. Check logs for details.")
            sys.exit(1)
            
    except Exception as e:
        logger.error(f"Error in main process: {e}")
        print(f"\nError: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
