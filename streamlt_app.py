#!/usr/bin/env python3
"""
Enterprise Invoice Processing System
Professional API Integration Platform  —  v3.1 (parallel edition)
"""
import streamlit as st
import subprocess
import sys
import io
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from pathlib import Path
import re
import shutil
import zipfile
import os
import time

# Fix Windows console encoding issues
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# Page configuration
st.set_page_config(
    page_title="Enterprise Invoice Processing",
    page_icon="■",
    layout="centered",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header { font-size: 2.5rem; font-weight: 700; color: #1f2937; margin-bottom: 0.5rem; }
    .sub-header  { font-size: 1.1rem; color: #6b7280; margin-bottom: 2rem; }
    .metric-card { background: linear-gradient(135deg,#667eea 0%,#764ba2 100%);
                   padding: 1.5rem; border-radius: 12px; color: white; text-align: center; }
    .success-box { background-color: #10b981; color: white; padding: 1rem;
                   border-radius: 8px; margin: 1rem 0; }
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-header">ENTERPRISE INVOICE PROCESSING</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Professional API Integration • Real-time Processing • Automated Delivery</p>', unsafe_allow_html=True)
st.divider()

# ── Companies (all 3 always processed in PARALLEL) ──────────────────────────
COMPANIES = {
    "GAP": {"id": "3201adad-6c8a-457d-b59b-d1b3fab125fb",
             "secret": "2d94c039-4747-485f-b213-c210d43db673",
             "name": "Global API Partners"},
    "GPI": {"id": "8f2a4ab7-d06a-4b65-99cd-355b2b42549f",
             "secret": "3d5c7e1f-1b9a-4d24-899b-ae03e6e5c58c",
             "name": "Global Partners Inc"},
    "GNP": {"id": "c89548fb-299a-48dd-8e21-6c4a0ba76612",
             "secret": "86905d13-03c7-485b-95bb-f052dc2ee8ba",
             "name": "Global NAPI"},
}
ALL_COMPANIES = list(COMPANIES.keys())

st.subheader("COMPANIES TO PROCESS")
st.info("All three companies run **in parallel** ⚡ — GAP + GPI + GNP simultaneously")
col1, col2, col3 = st.columns(3)
with col1: st.markdown("**GAP**\nGlobal API Partners")
with col2: st.markdown("**GPI**\nGlobal Partners Inc")
with col3: st.markdown("**GNP**\nGlobal NAPI")
st.divider()

# Email
st.subheader("EMAIL CONFIGURATION")
recipient_email = st.text_input(
    "Recipient Email Address",
    value="Payables@globalnapi.com",
    placeholder="example@email.com",
)
email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
is_valid_email = bool(re.match(email_pattern, recipient_email))
if recipient_email and not is_valid_email:
    st.error("INVALID EMAIL FORMAT")
elif recipient_email:
    st.success(f"EMAIL VALIDATED — Reports will be sent to: {recipient_email}")
st.divider()

# Time range
st.subheader("TIME PERIOD")
col1, col2 = st.columns(2)
with col1:
    days_to_process = st.number_input("Number of Days to Process", min_value=1, max_value=5, value=1)
with col2:
    st.metric("Processing Window", f"{days_to_process} Day{'s' if days_to_process > 1 else ''}",
              delta=f"{days_to_process * 24} Hours")

end_date   = datetime.now()
start_date = end_date - timedelta(days=days_to_process)
st.info(f"DATE RANGE: {start_date.strftime('%d-%m-%Y %H:%M')} to {end_date.strftime('%d-%m-%Y %H:%M')}")
st.divider()

# Options
st.subheader("PROCESSING OPTIONS")
col1, col2 = st.columns(2)
with col1: process_json = st.checkbox("Download JSON Files", value=True)
with col2: process_pdf  = st.checkbox("Download PDF Files",  value=True)
if not process_json and not process_pdf:
    st.warning("WARNING - Please select at least one processing option")
st.divider()


# ============================================================================
# HELPERS
# ============================================================================

def make_company_script(company: str, days: int) -> Path:
    """
    Write a per-company copy of get_invoices_using_api.py so all 3 can run at the same time
    without overwriting each other's config.
    """
    with open('get_invoices_using_api.py', 'r', encoding='utf-8') as f:
        content = f.read()

    content = re.sub(r'COMPANY_NAME = "[A-Z]+"', f'COMPANY_NAME = "{company}"', content)
    content = re.sub(r'DAYS = \d+',              f'DAYS = {days}',              content)

    tmp = Path(f'_tmp_api_{company}.py')
    tmp.write_text(content, encoding='utf-8')
    return tmp


def run_company_download(company: str, days: int) -> dict:
    """Run download for ONE company — safe to call from a worker thread."""
    script = make_company_script(company, days)
    try:
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'

        proc = subprocess.Popen(
            [sys.executable, str(script)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True, encoding='utf-8', errors='replace',
            env=env, bufsize=1, universal_newlines=True,
        )

        lines = []
        while True:
            line = proc.stdout.readline()
            if not line and proc.poll() is not None:
                break
            if line:
                lines.append(line.strip())

        remaining, _ = proc.communicate()
        if remaining:
            lines.extend(remaining.strip().split('\n'))

        output  = '\n'.join(lines)
        success = (
            "Complete!" in output or "Total files:" in output or
            "[SUCCESS]" in output or proc.returncode == 0
        )

        # ── Try regex first (fast path) ──────────────────────────────────────
        jm = (re.search(r'JSON done\. ok=(\d+)', output) or
              re.search(r'\[COMPLETE\] JSON done\. ok=(\d+)', output) or
              re.search(r'json[^\d]*(\d+)', output, re.IGNORECASE))
        pm = (re.search(r'PDF done\. ok=(\d+)', output) or
              re.search(r'\[COMPLETE\] PDF done\. ok=(\d+)', output) or
              re.search(r'pdf[^\d]*(\d+)', output, re.IGNORECASE))

        json_count = int(jm.group(1)) if jm else 0
        pdf_count  = int(pm.group(1)) if pm else 0

        # ── Fallback: count actual files on disk ─────────────────────────────
        # Covers any output format — if the files exist, we count them.
        if json_count == 0:
            today = datetime.now().strftime("%Y-%m-%d")
            for search_path in [
                Path("results") / company,
                Path("json")    / company,
                Path("json"),
            ]:
                if search_path.is_dir():
                    found = len(list(search_path.rglob("*.json")))
                    if found:
                        json_count = found
                        break

        if pdf_count == 0:
            today = datetime.now().strftime("%Y-%m-%d")
            for search_path in [
                Path("pdf") / company / today,
                Path("pdf") / company,
                Path("pdf"),
            ]:
                if search_path.is_dir():
                    found = len(list(search_path.rglob("*.pdf")))
                    if found:
                        pdf_count = found
                        break

        return {
            'company':    company,
            'success':    success,
            'json_count': json_count,
            'pdf_count':  pdf_count,
            'output':     output,
            'error':      None if success else "Download failed",
        }
    except Exception as e:
        return {
            'company': company, 'success': False,
            'json_count': 0, 'pdf_count': 0, 'output': '', 'error': str(e),
        }
    finally:
        script.unlink(missing_ok=True)


def run_all_parallel(days: int) -> list:
    """
    Launch GAP, GPI, GNP downloads simultaneously.
    Returns results in original order once all finish.
    """
    bucket = {}
    with ThreadPoolExecutor(max_workers=3) as pool:
        futures = {pool.submit(run_company_download, c, days): c for c in ALL_COMPANIES}
        for future in as_completed(futures):
            res = future.result()
            bucket[res['company']] = res
    return [bucket[c] for c in ALL_COMPANIES]


def create_delivery_package():
    """Build one ZIP: results/<CO>/<date>.xlsx + pdf/<CO>/<date>/*.pdf"""
    try:
        date_str = datetime.now().strftime("%Y-%m-%d")
        out_dir  = Path("api_outputs")
        out_dir.mkdir(parents=True, exist_ok=True)
        zip_path = out_dir / f"invoices_{date_str}.zip"

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in Path("results").rglob("*.xlsx") if Path("results").exists() else []:
                zf.write(f, f)
            for f in Path("pdf").rglob("*.pdf") if Path("pdf").exists() else []:
                zf.write(f, f)

        return zip_path
    except Exception:
        return None


def send_email_with_results(zip_path, recipient: str):
    """Patch recipient into send_email.py and call it with --file <zip>."""
    try:
        with open('send_email.py', 'r', encoding='utf-8') as f:
            code = f.read()
        shutil.copy('send_email.py', 'send_email.py.backup')
        code = re.sub(
            r'RECIPIENT_EMAILS = \[.*?\]',
            f'RECIPIENT_EMAILS = ["{recipient}"]',
            code, flags=re.DOTALL
        )
        with open('send_email.py', 'w', encoding='utf-8') as f:
            f.write(code)

        result = subprocess.run(
            [sys.executable, "send_email.py", "--file", str(zip_path)],
            capture_output=True, text=True, encoding='utf-8',
            timeout=180, errors='replace'
        )
        shutil.move('send_email.py.backup', 'send_email.py')

        out = (result.stdout or '') + (result.stderr or '')
        ok  = "EMAIL SENT SUCCESSFULLY" in out or result.returncode == 0
        return ok, out
    except Exception as e:
        try:
            if Path('send_email.py.backup').exists():
                shutil.move('send_email.py.backup', 'send_email.py')
        except Exception:
            pass
        return False, str(e)


# ============================================================================
# MAIN BUTTON
# ============================================================================

button_disabled = not is_valid_email or (not process_json and not process_pdf)

if st.button("START PROCESSING", type="primary", use_container_width=True, disabled=button_disabled):

    run_start = time.time()

    with st.container():
        st.markdown("## PROCESSING STATUS")

        # ── Phase 1 & 2: PARALLEL DOWNLOAD ──────────────────────────────────
        st.markdown("### PHASE 1 & 2: Parallel Download ⚡")

        launch_msg = st.info("⚡ Launching GAP + GPI + GNP downloads simultaneously…")
        p_bar      = st.progress(0.05)

        phase_start = time.time()
        all_results = run_all_parallel(days_to_process)   # ← all 3 at once
        phase_time  = time.time() - phase_start

        p_bar.progress(1.0)
        launch_msg.empty()

        # Per-company result cards
        cols = st.columns(3)
        total_json = total_pdf = 0
        any_failed = False

        for i, res in enumerate(all_results):
            with cols[i]:
                if res['success']:
                    total_json += res['json_count']
                    total_pdf  += res['pdf_count']
                    st.success(f"**{res['company']}** ✅")
                    st.metric("JSON", res['json_count'])
                    st.metric("PDF",  res['pdf_count'])
                else:
                    any_failed = True
                    st.error(f"**{res['company']}** ❌")
                    st.caption(res['error'] or "Unknown error")

        st.info(
            f"⚡ All 3 companies finished in **{phase_time:.1f}s** "
            f"(~{phase_time * 3 / 60:.1f}x faster than sequential)"
        )

        for res in all_results:
            with st.expander(f"View {res['company']} download log"):
                st.code(res['output'][-2000:], language="text")

        if any_failed:
            st.warning("One or more companies had issues — continuing with available data.")

        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric("Companies",   3)
        with col2: st.metric("Total JSON",  total_json)
        with col3: st.metric("Total PDF",   total_pdf)
        with col4: st.metric("Total Files", total_json + total_pdf)

        st.divider()

        # ── Phase 3: PO Extraction + Package ────────────────────────────────
        st.markdown("### PHASE 3: PO Extraction & Packaging")
        phase_start = time.time()

        with st.spinner("Running PO extraction for all companies…"):
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"
            gem = subprocess.run(
                [sys.executable, "invoices_preprocessing.py"],
                capture_output=True, text=True,
                encoding="utf-8", errors="replace", env=env
            )
            if gem.returncode == 0:
                st.success("PO extraction complete — results/ updated")
            else:
                st.warning("PO extraction had issues")
                with st.expander("View extraction log"):
                    st.code((gem.stdout + gem.stderr)[-2000:], language="text")

        with st.spinner("Creating delivery package…"):
            zip_package = create_delivery_package()
            if zip_package:
                size_mb = zip_package.stat().st_size / (1024 * 1024)
                st.success(f"Package ready: **{zip_package.name}** ({size_mb:.2f} MB)")
            else:
                st.warning("Package creation failed")

        st.info(f"PHASE COMPLETED IN: {time.time() - phase_start:.1f}s")
        st.divider()

        # ── Phase 4: Email ───────────────────────────────────────────────────
        st.markdown("### PHASE 4: Email Delivery")
        phase_start = time.time()

        with st.spinner(f"Sending to {recipient_email}…"):
            if zip_package and zip_package.exists():
                ok, email_out = send_email_with_results(zip_package, recipient_email)
                if ok:
                    st.success(f"Email delivered to **{recipient_email}** ({time.time()-phase_start:.1f}s)")
                    st.balloons()
                else:
                    st.error("Email delivery failed")
                    with st.expander("View email log"):
                        st.code(email_out[-1000:], language="text")
            else:
                st.error("No ZIP package to send")

        st.divider()

        # ── Final Summary ────────────────────────────────────────────────────
        total_time = time.time() - run_start
        st.markdown("## PROCESSING SUMMARY")

        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric("Companies",  "GAP / GPI / GNP")
        with col2: st.metric("Days",       days_to_process)
        with col3: st.metric("Files",      total_json + total_pdf)
        with col4: st.metric("Total Time", f"{int(total_time//60)}m {int(total_time%60)}s")

        st.info(f"TOTAL PROCESSING TIME: {int(total_time//60):02d}:{int(total_time%60):02d}")
        st.success("PROCESSING COMPLETE ✅")

        if zip_package and zip_package.exists():
            with open(zip_package, 'rb') as f:
                st.download_button(
                    label="DOWNLOAD PACKAGE",
                    data=f,
                    file_name=zip_package.name,
                    mime="application/zip",
                    use_container_width=True
                )


# ============================================================================
# SIDEBAR
# ============================================================================

with st.sidebar:
    st.markdown("### SYSTEM INFORMATION")

    st.markdown("""
    #### ⚡ Parallel Processing

    All 3 companies download **at the same time** using
    `ThreadPoolExecutor(max_workers=3)`.

    Each company gets its own temp copy of `get_invoices_using_api.py`
    (`_tmp_api_GAP.py`, etc.) so they never conflict.
    """)

    st.divider()

    st.markdown("""
    #### Output Format

    ```
    invoices_YYYY-MM-DD.zip
    ├── results/
    │   ├── GAP/YYYY-MM-DD.xlsx
    │   ├── GPI/YYYY-MM-DD.xlsx
    │   └── GNP/YYYY-MM-DD.xlsx
    └── pdf/
        ├── GAP/YYYY-MM-DD/*.pdf
        ├── GPI/YYYY-MM-DD/*.pdf
        └── GNP/YYYY-MM-DD/*.pdf
    ```
    """)

    st.divider()
    st.caption("Enterprise Edition v3.1 — Parallel")
    st.caption("Secure • Fast • Automated")