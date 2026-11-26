import asyncio
import ctypes
import html as html_lib
import json
import os
import random
import re
import socket
import subprocess
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse, urlunparse
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from playwright.async_api import TimeoutError as PlaywrightTimeoutError, async_playwright

TARGET_URL = "https://www.upwork.com/nx/s/find-work/most-recent"

now_local = datetime.now()
FOLDER_NAME = f"{now_local:%d}d-{now_local:%m}m-{now_local:%H}h-{now_local:%M}m"
OUTPUT_DIR = Path(FOLDER_NAME)

JSON_OUTPUT_FILE = OUTPUT_DIR / "upwork_jobs.json"
XLSX_OUTPUT_FILE = OUTPUT_DIR / "upwork_jobs.xlsx"
HTML_OUTPUT_FILE = OUTPUT_DIR / "upwork_jobs.html"

TARGET_JOB_COUNT = 30
MAX_LOAD_MORE_CLICKS = 3

REAL_EDGE_USER_DATA_DIR = Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Edge" / "User Data"
DEVTOOLS_PORT_FILE = REAL_EDGE_USER_DATA_DIR / "DevToolsActivePort"

STRUCTURED_FIELDS = [
    "scraped_at",
    "url",
    "title",
    "posted_at",
    "area",
    "time_required",
    "month_required",
    "experience",
    "rate",
    "connects_required",
    "client_from",
    "hired_till_now",
    "joined_at",
    "skills",
    "description",
    "proposals",
    "interviewing",
    "invites_sent",
    "unanswered_invites",
    "payment_info",
    "verified_or_not",
    "phone_verified_or_not",
    "last_view",
    "hire_rate",
    "open_jobs",
    "member_since",
]

CHALLENGE_MARKERS = (
    "verify you are human",
    "verify you're human",
    "unusual traffic",
    "automated requests",
    "captcha",
    "security check",
    "access denied",
    "temporarily blocked",
    "just a moment",
)


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


def position_edge_right_half_and_focus():
    if os.name != "nt":
        return

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    hwnds = []

    def enum_cb(hwnd, _lparam):
        if user32.IsWindowVisible(hwnd):
            length = user32.GetWindowTextLengthW(hwnd)
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                user32.GetWindowTextW(hwnd, buf, length + 1)
                if "Microsoft Edge" in buf.value or "Microsoft\u200b Edge" in buf.value:
                    hwnds.append(hwnd)
        return True

    EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    user32.EnumWindows(EnumWindowsProc(enum_cb), 0)
    if not hwnds:
        return

    hwnd = hwnds[0]
    if user32.IsIconic(hwnd) or user32.IsZoomed(hwnd):
        user32.ShowWindow(hwnd, 9)

    rect = RECT()
    if user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):
        left, top = rect.left, rect.top
        width, height = rect.right - rect.left, rect.bottom - rect.top
    else:
        left, top = 0, 0
        width, height = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)

    half_w = width // 2
    fg_hwnd = user32.GetForegroundWindow()
    fg_tid = user32.GetWindowThreadProcessId(fg_hwnd, None)
    tgt_tid = user32.GetWindowThreadProcessId(hwnd, None)
    cur_tid = kernel32.GetCurrentThreadId()

    user32.keybd_event(0x12, 0, 0, 0)
    user32.keybd_event(0x12, 0, 2, 0)

    attached_fg = bool(fg_tid and fg_tid != cur_tid and user32.AttachThreadInput(cur_tid, fg_tid, True))
    attached_tgt = bool(tgt_tid and tgt_tid != cur_tid and user32.AttachThreadInput(cur_tid, tgt_tid, True))
    try:
        user32.SetWindowPos(hwnd, 0, left + half_w, top, width - half_w, height, 0x0040)
        user32.SetForegroundWindow(hwnd)
        user32.BringWindowToTop(hwnd)
    finally:
        if attached_fg:
            user32.AttachThreadInput(cur_tid, fg_tid, False)
        if attached_tgt:
            user32.AttachThreadInput(cur_tid, tgt_tid, False)


def is_port_open(host: str, port: int, timeout: float = 0.5) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def read_live_devtools_ws_url() -> str | None:
    if not DEVTOOLS_PORT_FILE.exists():
        return None
    try:
        lines = DEVTOOLS_PORT_FILE.read_text(encoding="utf-8").strip().splitlines()
        if len(lines) >= 2 and lines[0].strip().isdigit() and lines[1].strip().startswith("/devtools/browser/"):
            port = int(lines[0].strip())
            if is_port_open("127.0.0.1", port):
                return f"ws://127.0.0.1:{port}{lines[1].strip()}"
    except OSError:
        pass
    try:
        DEVTOOLS_PORT_FILE.unlink()
    except OSError:
        pass
    return None


def ensure_edge_remote_debugging_ready() -> str:
    ws_url = read_live_devtools_ws_url()
    if ws_url:
        position_edge_right_half_and_focus()
        return ws_url

    print("\n=================================================================")
    print("[!] Edge Remote Debugging is currently OFF.")
    print("    Opening edge://inspect/#remote-debugging in your Edge window...")
    print("    1. Check the 'Enable remote debugging' box in that tab.")
    print("    2. Wait until 'Server running at: 127.0.0.1:...' appears.")
    print("=================================================================\n")

    subprocess.Popen(["cmd", "/c", "start", "", "msedge", "edge://inspect/#remote-debugging"])
    time.sleep(1.0)
    position_edge_right_half_and_focus()

    while True:
        time.sleep(1.5)
        ws_url = read_live_devtools_ws_url()
        if ws_url:
            print(f"[+] Live Edge debugging server confirmed at: {ws_url}")
            position_edge_right_half_and_focus()
            return ws_url


async def human_delay(min_sec: float = 1.5, max_sec: float = 4.0):
    mean = (min_sec + max_sec) / 2
    std_dev = (max_sec - min_sec) / 4
    delay = max(min_sec, min(max_sec, random.gauss(mean, std_dev)))
    await asyncio.sleep(delay)


async def human_mouse_move(page, target_x: float, target_y: float, steps: int = 25):
    start_x = random.randint(120, 750)
    start_y = random.randint(120, 550)

    ctrl_x = (start_x + target_x) / 2 + random.randint(-140, 140)
    ctrl_y = (start_y + target_y) / 2 + random.randint(-140, 140)

    for i in range(1, steps + 1):
        t = i / steps
        t_eased = t * t * (3 - 2 * t)
        x = (1 - t_eased) ** 2 * start_x + 2 * (1 - t_eased) * t_eased * ctrl_x + t_eased ** 2 * target_x
        y = (1 - t_eased) ** 2 * start_y + 2 * (1 - t_eased) * t_eased * ctrl_y + t_eased ** 2 * target_y
        await page.mouse.move(x, y)
        await asyncio.sleep(random.uniform(0.008, 0.022))


async def human_click_element(page, locator):
    await locator.scroll_into_view_if_needed()
    await human_delay(0.5, 1.2)
    box = await locator.bounding_box()
    if box:
        target_x = box["x"] + random.uniform(box["width"] * 0.25, box["width"] * 0.75)
        target_y = box["y"] + random.uniform(box["height"] * 0.25, box["height"] * 0.75)
        await human_mouse_move(page, target_x, target_y, steps=random.randint(20, 32))
        await human_delay(0.2, 0.5)
        await page.mouse.down()
        await asyncio.sleep(random.uniform(0.05, 0.14))
        await page.mouse.up()
    else:
        await locator.click()


async def human_scroll(page, scrolls: int = 3):
    for _ in range(scrolls):
        scroll_amount = random.randint(300, 680)
        await page.mouse.wheel(0, scroll_amount)
        await human_delay(0.8, 2.0)


async def detect_challenge(page) -> str | None:
    try:
        title = (await page.title() or "").lower()
        body = (await page.locator("body").inner_text(timeout=3000) or "").lower()
    except Exception:
        return None

    sample = f"{title}\n{body[:15000]}"
    for marker in CHALLENGE_MARKERS:
        if marker in sample:
            return marker
    return None


async def check_and_wait_for_captcha(page):
    marker = await detect_challenge(page)
    if marker or await page.locator("iframe[src*='challenges.cloudflare.com']").count() > 0:
        position_edge_right_half_and_focus()
        await page.bring_to_front()
        print(f"\n[!] Security verification detected ({marker or 'captcha'}). Please complete it manually in Edge...")
        while True:
            await asyncio.sleep(2.0)
            if not await detect_challenge(page) and await page.locator("iframe[src*='challenges.cloudflare.com']").count() == 0:
                break
        print("[+] Verification cleared! Resuming...\n")
        await human_delay(2.0, 3.5)


def canonical_job_url(href: str) -> str | None:
    parsed = urlparse(href)
    if parsed.netloc.lower() not in {"www.upwork.com", "upwork.com"}:
        return None
    if not parsed.path.startswith("/jobs/") or "~" not in parsed.path:
        return None
    clean_path = parsed.path.rstrip("/") + "/"
    return urlunparse(("https", "www.upwork.com", clean_path, "", "", ""))


async def ensure_logged_in(page):
    await check_and_wait_for_captcha(page)
    await human_delay(2.0, 3.0)

    needs_login = (
        "account-security/login" in page.url
        or await page.locator("input#login_username").count() > 0
        or await page.locator("input#login_password").count() > 0
    )

    if needs_login:
        position_edge_right_half_and_focus()
        await page.bring_to_front()
        print("\n=================================================================")
        print("[!] NOT LOGGED IN")
        print("    Please log in to Upwork in your Edge window now.")
        print("    Script will resume automatically once you reach the Find Work feed.")
        print("=================================================================\n")

        await page.wait_for_url("**/find-work/**", timeout=600000)
        await check_and_wait_for_captcha(page)
        print("[+] Login detected! Resuming automation...\n")
        await human_delay(2.0, 3.5)


async def collect_visible_job_urls(page) -> list[str]:
    raw_links = await page.evaluate("""() => {
        const anchors = Array.from(document.querySelectorAll('a[href*="/jobs/"]'));
        return anchors.map(a => a.href);
    }""")

    urls: list[str] = []
    seen_ids: set[str] = set()
    for link in raw_links:
        canon = canonical_job_url(link)
        if canon:
            job_id = canon.rstrip("/").split("~")[-1]
            if job_id and job_id not in seen_ids:
                seen_ids.add(job_id)
                urls.append(canon)
    return urls


async def click_load_more_jobs(page, previous_count: int) -> bool:
    await human_scroll(page, scrolls=random.randint(3, 5))

    load_more_btn = page.locator(
        'button:has-text("Load More Jobs"), button:has-text("Load more jobs"), [data-test="load-more-button"]'
    ).first

    if await load_more_btn.count() == 0:
        load_more_btn = page.get_by_role("button", name=re.compile(r"load\s+more", re.I)).first

    if await load_more_btn.count() == 0 or not await load_more_btn.is_visible():
        print("[!] 'Load More Jobs' button not found or not visible.")
        return False

    if await load_more_btn.is_disabled():
        print("[!] 'Load More Jobs' button is currently disabled.")
        return False

    print("[*] Clicking 'Load More Jobs' button...")
    await human_click_element(page, load_more_btn)
    await check_and_wait_for_captcha(page)

    deadline = time.monotonic() + 20.0
    while time.monotonic() < deadline:
        await asyncio.sleep(1.0)
        current_urls = await collect_visible_job_urls(page)
        if len(current_urls) > previous_count:
            print(f"[+] Loaded more jobs! Total visible jobs now: {len(current_urls)}")
            await human_delay(1.2, 2.5)
            return True

    print("[!] Timed out waiting for additional jobs to load after clicking 'Load More Jobs'.")
    return False


async def extract_job_urls_with_pagination(page, target_count: int = 30, max_clicks: int = 3) -> list[str]:
    print("[*] Waiting for Upwork Nuxt feed to hydrate...")
    await check_and_wait_for_captcha(page)

    try:
        await page.wait_for_selector('a[href*="/jobs/"]', timeout=30000)
    except PlaywrightTimeoutError:
        print("[!] Timed out waiting for job links on the feed page.")
        return []

    await human_mouse_move(page, random.randint(250, 850), random.randint(250, 650))
    await human_scroll(page, scrolls=2)

    collected = await collect_visible_job_urls(page)
    print(f"[*] Initial batch extracted: {len(collected)} unique jobs.")

    clicks = 0
    while len(collected) < target_count and clicks < max_clicks:
        clicks += 1
        print(f"[*] Paginating (Attempt {clicks}/{max_clicks}) to reach {target_count} jobs...")
        loaded = await click_load_more_jobs(page, previous_count=len(collected))
        if not loaded:
            break
        collected = await collect_visible_job_urls(page)

    return collected[:target_count]


def extract_connects(soup: BeautifulSoup, rendered_text: str = "") -> str:
    page_copy = BeautifulSoup(str(soup), "html.parser")
    for tag in page_copy.select("script, style, noscript"):
        tag.decompose()

    dom_text = " ".join(page_copy.get_text(" ", strip=True).split())
    combined = f"{rendered_text}\n{dom_text}"

    patterns = [
        r"required connects to submit a proposal[:\s]*(\d+)",
        r"send a proposal for[:\s]*(\d+)\s*connects?",
        r"connects?\s+required[:\s]*(\d+)",
        r"\b(\d+)\s+connects?\b",
    ]
    for pat in patterns:
        m = re.search(pat, combined, flags=re.I)
        if m:
            return f"{m.group(1).strip()} Connects"
    return "N/A"


def parse_job_details(html_content: str, url: str, rendered_text: str = "") -> dict:
    soup = BeautifulSoup(html_content, "html.parser")

    title_el = soup.select_one(".job-details-content h1, .job-details-content h4, main h1, main h4, h1")
    title = title_el.get_text(" ", strip=True) if title_el and title_el.get_text(strip=True) else (
        soup.title.string.strip() if soup.title and soup.title.string else "N/A"
    )

    posted_line = soup.select_one(".posted-on-line, [data-test='posted-on']")
    posted_raw = posted_line.get_text(" | ", strip=True) if posted_line else ""
    posted_parts = [p.strip() for p in posted_raw.split("|") if p.strip() and p.strip().lower() != "posted"]
    posted_at = "N/A"
    area = "Worldwide"
    for p_item in posted_parts:
        m_time = re.search(r"\b(\d+\s+(?:second|minute|hour|day|week|month)s?\s+ago|yesterday|today)\b", p_item, re.I)
        if m_time and posted_at == "N/A":
            posted_at = m_time.group(1).strip()
        elif p_item.lower() not in {"posted"} and area == "Worldwide":
            area = p_item.strip()

    features = [
        li.get_text(" ", strip=True)
        for li in soup.select("ul.features > li")
        if not li.find_parent(class_="cfe-ui-job-about-client") and li.get_text(strip=True)
    ]

    time_required = "N/A"
    month_required = "N/A"
    experience = "N/A"
    rate = "N/A"

    for feat in features:
        low = feat.lower()
        if "hrs/week" in low or "hours to be determined" in low:
            time_required = re.sub(r"\s+Hourly$", "", feat, flags=re.I).strip()
        elif "duration" in low or "months" in low or "less than 1 month" in low:
            m_dur = re.match(r"^(Less than 1 month|More than 6 months|\d+\s+to\s+\d+\s+months)", feat, re.I)
            month_required = m_dur.group(1) if m_dur else re.sub(r"\s*Duration$", "", feat, flags=re.I).strip()
        elif "experience level" in low or feat.startswith(("Entry level", "Intermediate", "Expert")):
            for lvl in ("Entry level", "Intermediate", "Expert"):
                if lvl.lower() in low:
                    experience = lvl
                    break
        elif "$" in feat and ("hourly" in low or "fixed-price" in low or "fixed price" in low):
            rate = feat.strip()

    if rate == "N/A" and any("hourly" in f.lower() for f in features):
        rate = "Hourly (Open Rate)"

    connects_required = extract_connects(soup, rendered_text=rendered_text)

    skills = list(dict.fromkeys(
        s.get_text(strip=True)
        for s in soup.select(".skills-list .badge, [data-test='Skill']")
        if s.get_text(strip=True)
    ))

    desc_el = soup.select_one('[data-test="Description"], .job-details-content .multiline-text, .multiline-text')
    description = desc_el.get_text("\n", strip=True) if desc_el else "N/A"
    description = re.sub(r"^Summary\s*\n+", "", description, flags=re.I).strip()

    proposals = "N/A"
    last_view = "N/A"
    activity_hires = None
    interviewing = "0"
    invites_sent = "0"
    unanswered_invites = "0"

    for item in soup.select(".client-activity-items .ca-item"):
        item_copy = BeautifulSoup(str(item), "html.parser")
        for pop in item_copy.select(".air3-popover, [role='tooltip'], .sr-only"):
            pop.decompose()
        txt = re.sub(
            r"Close the tooltip.*?(?=\d+\s+to\s+\d+|\d+\+|yesterday|\d+\s+(?:minute|hour|day))",
            "",
            item_copy.get_text(" ", strip=True),
            flags=re.I,
        ).strip()
        low_txt = txt.lower()
        if low_txt.startswith("proposals:"):
            proposals = txt.split(":", 1)[1].strip()
        elif low_txt.startswith("last viewed by client:"):
            last_view = txt.split(":", 1)[1].strip()
        elif low_txt.startswith("hires:"):
            activity_hires = txt.split(":", 1)[1].strip()
        elif low_txt.startswith("interviewing:"):
            interviewing = txt.split(":", 1)[1].strip()
        elif low_txt.startswith("invites sent:"):
            invites_sent = txt.split(":", 1)[1].strip()
        elif low_txt.startswith("unanswered invites:"):
            unanswered_invites = txt.split(":", 1)[1].strip()

    client_section = soup.select_one(".cfe-ui-job-about-client")
    client_raw = ""
    if client_section:
        client_copy = BeautifulSoup(str(client_section), "html.parser")
        for pop in client_copy.select(".air3-popover, [role='tooltip'], .sr-only"):
            pop.decompose()
        client_raw = client_copy.get_text(" | ", strip=True)

    is_pay_ver = "payment method verified" in client_raw.lower() and "not verified" not in client_raw.lower()
    payment_info = "Payment method verified" if is_pay_ver else "Payment method not verified"
    verified_or_not = "Verified" if is_pay_ver else "Not Verified"
    phone_verified_or_not = "Verified" if "phone number verified" in client_raw.lower() else "Not Verified"

    hire_m = re.search(r"(\d+%)\s*hire rate", client_raw, re.I)
    hire_rate = hire_m.group(1) if hire_m else "0%"

    open_m = re.search(r"(\d+)\s*open jobs?", client_raw, re.I)
    open_jobs = open_m.group(1) if open_m else "0"

    hires_m = re.search(r"(\d+\s*hires?(?:,\s*\d+\s*active)?)", client_raw, re.I)
    spent_m = re.search(r"(\$[\d,.]+[KMB]?)\s*\|\s*total spent|\$[\d,.]+[KMB]?\s*total spent", client_raw, re.I)
    if activity_hires is not None:
        hired_till_now = f"{activity_hires} {'hire' if activity_hires == '1' else 'hires'}"
    elif hires_m:
        hired_till_now = f"{hires_m.group(1)} ({spent_m.group(1) if spent_m and spent_m.group(1) else 'spent'})" if spent_m else hires_m.group(1)
    elif hire_rate != "0%":
        hired_till_now = f"Hired ({hire_rate} hire rate)"
    else:
        hired_till_now = "0 hires"

    member_m = re.search(r"Member since\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})", client_raw, re.I)
    member_since = member_m.group(1) if member_m else "N/A"

    client_from = "N/A"
    cleaned_for_loc = re.sub(
        r"(?:Phone number verified|Payment method (?:not )?verified|Close the tooltip|About the client|Rating is [\d.]+ out of 5|[\d.]+\s*of\s*\d+\s*reviews?)",
        "",
        client_raw,
        flags=re.I,
    )
    loc_m = re.search(
        r"\|\s*([A-Za-z][A-Za-z\s.-]+?)(?:\s*\|\s*([A-Za-z][A-Za-z\s.-]+?))?\s*\|\s*\d{1,2}:\d{2}\s*[AP]M",
        cleaned_for_loc,
        re.I,
    )
    if loc_m:
        c1 = loc_m.group(1).strip(" |")
        c2 = loc_m.group(2).strip(" |") if loc_m.group(2) else ""
        client_from = f"{c1}, {c2}" if c2 else c1

    return {
        "scraped_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "url": url,
        "title": title,
        "posted_at": posted_at,
        "area": area,
        "time_required": time_required,
        "month_required": month_required,
        "experience": experience,
        "rate": rate,
        "connects_required": connects_required,
        "client_from": client_from,
        "hired_till_now": hired_till_now,
        "joined_at": member_since,
        "skills": skills,
        "description": description,
        "proposals": proposals,
        "interviewing": interviewing,
        "invites_sent": invites_sent,
        "unanswered_invites": unanswered_invites,
        "payment_info": payment_info,
        "verified_or_not": verified_or_not,
        "phone_verified_or_not": phone_verified_or_not,
        "last_view": last_view,
        "hire_rate": hire_rate,
        "open_jobs": open_jobs,
        "member_since": member_since,
    }


def save_xlsx_sheet(results: list[dict]):
    wb = Workbook()
    ws = wb.active
    ws.title = "Jobs Data"
    ws.views.sheetView[0].showGridLines = True

    navy_header = "1E2D3D"
    zebra_bg = "F7F9FB"
    white = "FFFFFF"
    border_gray = "D5DCE4"

    thin_side = Side(style="thin", color=border_gray)
    med_navy = Side(style="medium", color=navy_header)
    grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    header_border = Border(left=thin_side, right=thin_side, top=med_navy, bottom=med_navy)

    ws.append(STRUCTURED_FIELDS)
    for col_idx in range(1, len(STRUCTURED_FIELDS) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(name="Calibri", size=11, bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=navy_header)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = header_border
    ws.row_dimensions[1].height = 28

    for r_idx, item in enumerate(results, start=2):
        fill_color = zebra_bg if r_idx % 2 == 1 else white
        for c_idx, col_name in enumerate(STRUCTURED_FIELDS, start=1):
            val = item.get(col_name, "")
            if col_name == "skills" and isinstance(val, list):
                val = ", ".join(val)

            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = Font(name="Calibri", size=10, color="1A202C")
            cell.fill = PatternFill("solid", fgColor=fill_color)
            cell.border = grid_border

            if col_name == "url" and val:
                cell.hyperlink = val
                cell.font = Font(name="Calibri", size=10, color="1B5E20", underline="single")
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=False)
            elif col_name == "description":
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        desc_len = len(str(item.get("description", "")))
        ws.row_dimensions[r_idx].height = min(140, max(24, (desc_len // 75 + 1) * 15))

    col_widths = {
        "scraped_at": 22,
        "url": 42,
        "title": 36,
        "posted_at": 15,
        "area": 14,
        "time_required": 24,
        "month_required": 18,
        "experience": 15,
        "rate": 24,
        "connects_required": 18,
        "client_from": 22,
        "hired_till_now": 26,
        "joined_at": 16,
        "skills": 36,
        "description": 52,
        "proposals": 14,
        "interviewing": 14,
        "invites_sent": 14,
        "unanswered_invites": 18,
        "payment_info": 26,
        "verified_or_not": 16,
        "phone_verified_or_not": 20,
        "last_view": 18,
        "hire_rate": 13,
        "open_jobs": 12,
        "member_since": 16,
    }
    for c_idx, col_name in enumerate(STRUCTURED_FIELDS, start=1):
        ws.column_dimensions[get_column_letter(c_idx)].width = col_widths.get(col_name, 18)

    ws.freeze_panes = "D2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(STRUCTURED_FIELDS))}{max(2, len(results) + 1)}"

    dv_ver = DataValidation(type="list", formula1='"Verified,Not Verified"', allow_blank=True)
    ws.add_data_validation(dv_ver)
    ver_col_letter = get_column_letter(STRUCTURED_FIELDS.index("verified_or_not") + 1)
    dv_ver.add(f"{ver_col_letter}2:{ver_col_letter}{max(2, len(results) + 1)}")

    wb.save(XLSX_OUTPUT_FILE)


def save_results(results: list[dict]):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    JSON_OUTPUT_FILE.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    save_xlsx_sheet(results)

    cols = ["#"] + STRUCTURED_FIELDS
    th_row = "".join(f"<th onclick='sortCol({i})'>{c}</th>" for i, c in enumerate(cols))
    filter_row = "".join(
        f"<th><input type='text' class='f-input' data-col='{i}' placeholder='Filter...' oninput='filterRows()' onclick='event.stopPropagation()'></th>"
        for i in range(len(cols))
    )

    body_rows = []
    for idx, item in enumerate(results, start=1):
        cells = [f"<td>{idx}</td>"]
        for field in STRUCTURED_FIELDS:
            val = ", ".join(item.get(field, [])) if field == "skills" and isinstance(item.get(field), list) else str(item.get(field, ""))
            esc = html_lib.escape(val)
            if field == "url" and val:
                cells.append(f"<td><a href='{esc}' target='_blank'>{esc}</a></td>")
            elif field == "description":
                cells.append(f"<td><div class='desc'>{esc}</div></td>")
            else:
                cells.append(f"<td>{esc}</td>")
        body_rows.append("<tr>" + "".join(cells) + "</tr>")

    html_doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Upwork Jobs - {FOLDER_NAME}</title>
<style>
  body {{ font-family: Arial, sans-serif; margin: 12px; background: #12161f; color: #e6edf3; }}
  .bar {{ display: flex; gap: 10px; align-items: center; margin-bottom: 10px; }}
  input {{ background: #1b2230; color: #e6edf3; border: 1px solid #303a4f; border-radius: 4px; padding: 5px 8px; }}
  button {{ background: #1f8a4c; color: #fff; border: none; border-radius: 4px; padding: 6px 12px; cursor: pointer; }}
  .wrap {{ overflow: auto; max-height: 85vh; border: 1px solid #303a4f; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 12px; }}
  th, td {{ border: 1px solid #303a4f; padding: 6px 8px; text-align: left; vertical-align: top; white-space: nowrap; }}
  thead tr.h th {{ background: #195c36; color: #fff; position: sticky; top: 0; cursor: pointer; z-index: 2; }}
  thead tr.f th {{ background: #18202c; position: sticky; top: 30px; z-index: 1; padding: 4px; }}
  .f-input {{ width: 100%; min-width: 90px; box-sizing: border-box; }}
  .desc {{ max-width: 360px; max-height: 90px; overflow: auto; white-space: pre-wrap; }}
  a {{ color: #58a6ff; }}
</style>
</head>
<body>
  <div class="bar">
    <strong>{FOLDER_NAME}</strong>
    <input id="qAll" type="text" placeholder="Search all..." oninput="filterRows()">
    <button onclick="resetFilters()">Reset</button>
    <span id="countLbl"></span>
  </div>
  <div class="wrap">
    <table id="tbl">
      <thead>
        <tr class="h">{th_row}</tr>
        <tr class="f">{filter_row}</tr>
      </thead>
      <tbody>{"".join(body_rows)}</tbody>
    </table>
  </div>
<script>
function cellVal(tr, c) {{
  return (tr.children[c]?.innerText || "").replace(/\\s+/g, " ").trim().toLowerCase();
}}
function filterRows() {{
  const allQ = document.getElementById("qAll").value.trim().toLowerCase();
  const filters = Array.from(document.querySelectorAll(".f-input"))
    .map(el => [Number(el.dataset.col), el.value.trim().toLowerCase()])
    .filter(([, v]) => v.length > 0);
  const rows = Array.from(document.querySelectorAll("#tbl tbody tr"));
  let shown = 0;
  for (const r of rows) {{
    const full = (r.innerText || "").replace(/\\s+/g, " ").trim().toLowerCase();
    const okAll = !allQ || full.includes(allQ);
    const okCols = filters.every(([col, v]) => cellVal(r, col).includes(v));
    const vis = okAll && okCols;
    r.style.display = vis ? "" : "none";
    if (vis) shown++;
  }}
  document.getElementById("countLbl").textContent = shown + " / " + rows.length;
}}
function resetFilters() {{
  document.getElementById("qAll").value = "";
  document.querySelectorAll(".f-input").forEach(el => el.value = "");
  filterRows();
}}
function sortCol(c) {{
  const tb = document.querySelector("#tbl tbody");
  const rows = Array.from(tb.querySelectorAll("tr"));
  const asc = tb.dataset.sortCol != c || tb.dataset.sortDir !== "asc";
  tb.dataset.sortCol = c;
  tb.dataset.sortDir = asc ? "asc" : "desc";
  rows.sort((a, b) => asc
    ? cellVal(a, c).localeCompare(cellVal(b, c), undefined, {{ numeric: true }})
    : cellVal(b, c).localeCompare(cellVal(a, c), undefined, {{ numeric: true }}));
  rows.forEach(r => tb.appendChild(r));
}}
filterRows();
</script>
</body>
</html>"""
    HTML_OUTPUT_FILE.write_text(html_doc, encoding="utf-8")


async def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    ws_url = ensure_edge_remote_debugging_ready()

    results: list[dict] = []
    save_results(results)

    async with async_playwright() as p:
        print("[*] Connecting to your existing Edge session (Person 1 / Default)...")
        position_edge_right_half_and_focus()
        print("    >>> LOOK AT YOUR EDGE WINDOW AND CLICK 'ALLOW' WHEN THE POPUP APPEARS <<<")
        try:
            browser = await p.chromium.connect_over_cdp(ws_url, timeout=60000)
        except Exception as e:
            print(
                "\n[!] Connection failed or timed out.\n"
                "    - Make sure you clicked 'Allow' on the Edge permission popup.\n"
                f"    Error: {e}"
            )
            return

        context = browser.contexts[0]
        page = context.pages[0] if context.pages else await context.new_page()

        position_edge_right_half_and_focus()
        await page.bring_to_front()

        try:
            print(f"[*] Navigating to {TARGET_URL}...")
            await page.goto(TARGET_URL, wait_until="domcontentloaded", timeout=45000)

            await ensure_logged_in(page)

            job_urls = await extract_job_urls_with_pagination(
                page,
                target_count=TARGET_JOB_COUNT,
                max_clicks=MAX_LOAD_MORE_CLICKS,
            )
            print(f"[+] Total unique job URLs collected: {len(job_urls)}")

            if not job_urls:
                print("[!] No job URLs found. Check if Upwork loaded properly in the browser tab.")
                return

            for idx, job_url in enumerate(job_urls, start=1):
                nav_url = f"{job_url}?referrer_url_path=find_work_home"
                print(f"\n[{idx}/{len(job_urls)}] Visiting: {job_url}")
                await human_delay(2.5, 5.0)

                await human_mouse_move(page, random.randint(200, 950), random.randint(200, 650))

                try:
                    await page.goto(nav_url, wait_until="domcontentloaded", referer=TARGET_URL, timeout=45000)
                except PlaywrightTimeoutError:
                    print(f"[!] Navigation timeout on {job_url}, skipping...")
                    continue

                await check_and_wait_for_captcha(page)

                hydrated = False
                try:
                    await page.wait_for_selector(
                        '.job-details-content, [data-test="Description"], .job-details-card, .cfe-ui-job-about-client, main h1, main h4',
                        timeout=25000,
                    )
                    hydrated = True
                except PlaywrightTimeoutError:
                    page_title = await page.title()
                    print(f"[!] Job details did not hydrate within 25s (Title: {page_title!r}, URL: {page.url})")

                if not hydrated:
                    continue

                await human_delay(2.0, 4.0)
                await human_scroll(page, scrolls=random.randint(2, 4))
                await human_mouse_move(page, random.randint(300, 750), random.randint(250, 650))

                html = await page.content()
                rendered_text = await page.evaluate("() => document.body ? document.body.innerText : ''")
                job_data = parse_job_details(html, job_url, rendered_text=rendered_text)
                results.append(job_data)
                save_results(results)
                print(f"    -> Scraped: {job_data['title']} | Connects: {job_data['connects_required']} | Hires: {job_data['hired_till_now']}")

            print(
                f"\n[+] Finished! Saved {len(results)} jobs inside folder '{OUTPUT_DIR.resolve()}':\n"
                f"    - JSON : {JSON_OUTPUT_FILE.resolve()}\n"
                f"    - XLSX : {XLSX_OUTPUT_FILE.resolve()}\n"
                f"    - Sheet: {HTML_OUTPUT_FILE.resolve()}"
            )

        finally:
            await page.close()


if __name__ == "__main__":
    asyncio.run(main())