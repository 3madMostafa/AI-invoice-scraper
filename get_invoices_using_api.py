import os, time, requests, sys
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock

# Fix Windows console encoding issues
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except AttributeError:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = "https://api.invoicing.eta.gov.eg"
TOKEN_URL = "https://id.eta.gov.eg/connect/token"

# Company credentials
COMPANIES = {
    "GAP": {
        "id": "3201adad-6c8a-457d-b59b-d1b3fab125fb",
        "secret": "2d94c039-4747-485f-b213-c210d43db673"
    },
    "GPI": {
        "id": "8f2a4ab7-d06a-4b65-99cd-355b2b42549f",
        "secret": "3d5c7e1f-1b9a-4d24-899b-ae03e6e5c58c"
    },
    "GNP": {
        "id": "c89548fb-299a-48dd-8e21-6c4a0ba76612",
        "secret": "86905d13-03c7-485b-95bb-f052dc2ee8ba"
    }
}

DAYS = 1
PAGE_SIZE = 100

# Reduced parallelism to avoid rate limits
MAX_WORKERS_JSON = 2      # Reduced from 6 to 2
TIMEOUT = 90

# Increased delays to respect rate limits
JSON_SLEEP = 0.5          # Small delay between JSON requests
PDF_SLEEP = 1.5           # Increased from 1.2
PDF_MAX_RETRIES = 15      # Increased retries

def get_token(session: requests.Session, client_id: str, client_secret: str):
    if not client_id or not client_secret:
        raise RuntimeError("Missing client_id / client_secret")

    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret
    }
    r = session.post(TOKEN_URL, data=data, timeout=60)
    r.raise_for_status()
    return r.json()["access_token"]

def request_with_retry(session, method, url, headers=None, params=None,
                       max_retries=15, base_delay=2.0, max_delay=60):
    delay = base_delay
    for attempt in range(1, max_retries + 1):
        r = session.request(method, url, headers=headers, params=params, timeout=TIMEOUT)

        if r.status_code < 400:
            return r

        if r.status_code == 429:
            ra = r.headers.get("Retry-After")
            if ra:
                sleep_s = float(ra)
                print(f"[Rate Limit] Server says wait {sleep_s}s")
            else:
                sleep_s = delay
                print(f"[Rate Limit] Attempt {attempt}/{max_retries}, waiting {sleep_s:.1f}s")
            
            time.sleep(sleep_s)
            delay = min(delay * 2, max_delay)
            continue

        if r.status_code == 401:
            raise PermissionError("401 Unauthorized")

        raise RuntimeError(f"{r.status_code} {r.reason}: {r.text}")

    raise RuntimeError(f"429 Too Many Requests persisted after {max_retries} retries")

def fmt_utc(dt: datetime) -> str:
    return dt.strftime("%Y-%m-%dT%H:%M:%SZ")

def search_documents_uuids(session, headers, direction, from_dt, to_dt):
    url = f"{BASE}/api/v1.0/documents/search"
    token = ""
    uuids = []
    page = 1

    base_params = {
        "submissionDateFrom": fmt_utc(from_dt),
        "submissionDateTo": fmt_utc(to_dt),
        "pageSize": PAGE_SIZE,
    }
    if direction:
        base_params["direction"] = direction

    while True:
        params = dict(base_params)
        if token:
            params["continuationToken"] = token

        r = request_with_retry(session, "GET", url, headers=headers, params=params)
        data = r.json()

        docs = data.get("result", []) or []
        meta = data.get("metadata", {}) or {}
        token = (meta.get("continuationToken") or "").strip()

        uuids.extend([d.get("uuid") for d in docs if d.get("uuid")])
        next_status = 'End' if token=='EndofResultSet' else 'More'
        print(f"{direction or 'ALL'}: page {page} -> {len(docs)} docs | nextToken={next_status}")

        if token == "EndofResultSet":
            break
        if not docs and not token:
            break

        page += 1
        time.sleep(0.3)  # Small delay between search pages

    return uuids

def main(COMPANY_NAME):
    CLIENT_ID = COMPANIES[COMPANY_NAME]["id"]
    CLIENT_SECRET = COMPANIES[COMPANY_NAME]["secret"]

    print(f"\n{'='*60}")
    print(f"[INFO] Running for company: {COMPANY_NAME}")
    print(f"[INFO] Client ID: {CLIENT_ID}")
    print(f"{'='*60}")
    
    # Use Cairo local time (UTC+2) to determine "yesterday"
    CAIRO_OFFSET = timedelta(hours=2)
    today         = (datetime.now(timezone.utc) + CAIRO_OFFSET).date()
    to_date_d     = today - timedelta(days=1)      # yesterday in Cairo time

    # API range: yesterday Cairo 00:00 → 23:59 converted to UTC
    from_dt = datetime(to_date_d.year, to_date_d.month, to_date_d.day,
                       0, 0, 0, tzinfo=timezone.utc) - CAIRO_OFFSET   # 22:00 UTC prev day
    to_dt   = datetime(to_date_d.year, to_date_d.month, to_date_d.day,
                       23, 59, 59, tzinfo=timezone.utc) - CAIRO_OFFSET # 21:59 UTC same day

    # Folder name = yesterday in Cairo time (always a single clean date)
    date_folder = to_date_d.strftime("%Y-%m-%d")

    json_dir = os.path.join("json", COMPANY_NAME, date_folder)
    pdf_dir  = os.path.join("pdf",  COMPANY_NAME, date_folder)
    os.makedirs(json_dir, exist_ok=True)
    os.makedirs(pdf_dir,  exist_ok=True)

    session = requests.Session()
    token = get_token(session, CLIENT_ID, CLIENT_SECRET)
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    try:
        received = search_documents_uuids(session, headers, "Received", from_dt, to_dt)
    except PermissionError:
        print("[AUTH] Token expired, refreshing...")
        token = get_token(session, CLIENT_ID, CLIENT_SECRET)
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        received = search_documents_uuids(session, headers, "Received", from_dt, to_dt)

    all_uuids = received
    print(f"\n[STATS] Total RECEIVED UUIDs: {len(all_uuids)}")


    tasks = []
    seen = {}
    for u in all_uuids:
        seen[u] = seen.get(u, 0) + 1
        tasks.append((u, seen[u]))

    token_lock = Lock()
    request_lock = Lock()  # Add lock to control request timing
    
    def ensure_token():
        nonlocal token, headers
        with token_lock:
            token = get_token(session, CLIENT_ID, CLIENT_SECRET)
            headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    def save_bytes(path, content: bytes):
        with open(path, "wb") as f:
            f.write(content)

    def download_json(uuid, idx):
        suffix = f"_{idx}" if idx > 1 else ""
        path = os.path.join(json_dir, f"{uuid}{suffix}.json")
        if os.path.exists(path):
            return ("skip", os.path.basename(path))

        # Add small delay before each request
        with request_lock:
            time.sleep(JSON_SLEEP)

        url = f"{BASE}/api/v1.0/documents/{uuid}/raw"
        h = dict(headers); h["Accept"] = "application/json"

        try:
            r = request_with_retry(session, "GET", url, headers=h, max_retries=15)
        except PermissionError:
            ensure_token()
            h = dict(headers); h["Accept"] = "application/json"
            r = request_with_retry(session, "GET", url, headers=h, max_retries=15)

        save_bytes(path, r.content)
        return ("ok", os.path.basename(path))

    print("\n[DOWNLOAD] Starting JSON downloads...")
    ok_json = skip_json = fail_json = 0
    with ThreadPoolExecutor(max_workers=MAX_WORKERS_JSON) as ex:
        futures = [ex.submit(download_json, u, idx) for (u, idx) in tasks]
        for i, fut in enumerate(as_completed(futures), 1):
            try:
                status, _ = fut.result()
                if status == "ok":
                    ok_json += 1
                else:
                    skip_json += 1
                
                # Progress indicator
                if i % 10 == 0:
                    print(f"[PROGRESS] {i}/{len(tasks)} - OK: {ok_json}, Skip: {skip_json}, Fail: {fail_json}")
            except Exception as e:
                fail_json += 1
                print(f"[JSON FAIL] {e}")

    print(f"\n[COMPLETE] JSON done. ok={ok_json}, skip={skip_json}, fail={fail_json}")
    if fail_json > 0:
        raise RuntimeError(f"{fail_json} JSON files failed to download")

    print("\n[DOWNLOAD] Starting PDF downloads...")
    ok_pdf = skip_pdf = fail_pdf = 0

    def download_pdf_one(uuid, idx):
        suffix = f"_{idx}" if idx > 1 else ""
        path = os.path.join(pdf_dir, f"{uuid}{suffix}.pdf")
        if os.path.exists(path):
            return "skip"

        for endpoint in ("pdf", "printout"):
            url = f"{BASE}/api/v1.0/documents/{uuid}/{endpoint}"
            h = dict(headers); h["Accept"] = "application/pdf"

            try:
                r = request_with_retry(
                    session, "GET", url, headers=h,
                    max_retries=PDF_MAX_RETRIES,
                    base_delay=2.0, max_delay=60
                )
                save_bytes(path, r.content)
                return "ok"
            except RuntimeError as e:
                if "404" in str(e) and endpoint == "pdf":
                    continue
                raise

        raise RuntimeError("PDF endpoint not found (tried /pdf and /printout)")

    for i, (u, idx) in enumerate(tasks, 1):
        try:
            status = download_pdf_one(u, idx)
            if status == "ok":
                ok_pdf += 1
            else:
                skip_pdf += 1
        except PermissionError:
            ensure_token()
            try:
                status = download_pdf_one(u, idx)
                if status == "ok":
                    ok_pdf += 1
                else:
                    skip_pdf += 1
            except Exception as e:
                fail_pdf += 1
                print(f"[PDF FAIL] {u}: {e}")
        except Exception as e:
            fail_pdf += 1
            print(f"[PDF FAIL] {u}: {e}")

        # Progress indicator
        if i % 10 == 0:
            print(f"[PROGRESS] {i}/{len(tasks)} - OK: {ok_pdf}, Skip: {skip_pdf}, Fail: {fail_pdf}")
        
        time.sleep(PDF_SLEEP)

    print(f"\n[COMPLETE] PDF done. ok={ok_pdf}, skip={skip_pdf}, fail={fail_pdf}")
    if fail_pdf > 0:
        raise RuntimeError(f"{fail_pdf} PDF files failed to download")
    print(f"\n[OUTPUT] json/{COMPANY_NAME}/{date_folder}")
    print(f"[OUTPUT] pdf/{COMPANY_NAME}/{date_folder}")
    print(f"\n[SUCCESS] Complete! Total files: {ok_json + ok_pdf}")

if __name__ == "__main__":
    MAX_COMPANY_RETRIES = 5
    RETRY_WAIT = 30  # seconds between company-level retries

    for company in COMPANIES:
        for attempt in range(1, MAX_COMPANY_RETRIES + 1):
            print(f"\n[ATTEMPT {attempt}/{MAX_COMPANY_RETRIES}] Company: {company}")
            try:
                main(company)
                print(f"[OK] {company} completed successfully on attempt {attempt}")
                break  # success → move to next company
            except Exception as e:
                print(f"[RETRY] {company} attempt {attempt} failed: {e}")
                if attempt < MAX_COMPANY_RETRIES:
                    print(f"[WAIT] Retrying in {RETRY_WAIT}s...")
                    time.sleep(RETRY_WAIT)
                else:
                    print(f"[FAIL] {company} failed after {MAX_COMPANY_RETRIES} attempts — moving on")