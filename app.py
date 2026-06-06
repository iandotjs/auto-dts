import argparse
import getpass
import os
import sys
import time

import pandas as pd
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# --- Configurations ---
EXCEL_FILE = "timesheet_data.xlsx"
BASE_URL = "http://www.ntsp.nec.co.jp/projects/web"
LOGIN_URL = f"{BASE_URL}/login_check"
LOGIN_PAGE_URL = f"{BASE_URL}/login"
TIMESHEET_URL = f"{BASE_URL}/dtstimesheet/"
UPDATE_URL = f"{BASE_URL}/dtstimesheet/update"
DEFAULT_USERNAME = "dinopol.ij"
DEFAULT_PROJECT_ID = "484"
DEFAULT_WORKORDER = "S26-000-12V-00"

REQUIRED_COLUMNS = ["Date", "Type", "Activity Name", "Regular Hours", "OT Hours"]

# Keep these dictionaries complete and synced with internal BSS master data IDs.
DIRECT_ACTIVITIES = {
    "Coding > [Tool] Source Code Modification": 42398,
    "Testing > UT Execution": 42383,
}

INDIRECT_ACTIVITIES = {
    "Progress meeting": {"category": 1, "activity": 7},
    "Company or group activities": {"category": 32, "activity": 92},
}


def parse_args():
    parser = argparse.ArgumentParser(
        description="Auto-DTS: submit daily timesheets from an Excel file."
    )
    parser.add_argument("--excel", default=EXCEL_FILE, help="Path to timesheet Excel file")
    parser.add_argument(
        "--username",
        default=os.getenv("AUTO_DTS_USERNAME", DEFAULT_USERNAME),
        help="HR portal username (or set AUTO_DTS_USERNAME)",
    )
    parser.add_argument(
        "--password",
        default=os.getenv("AUTO_DTS_PASSWORD"),
        help="HR portal password (or set AUTO_DTS_PASSWORD)",
    )
    parser.add_argument(
        "--project",
        default=DEFAULT_PROJECT_ID,
        help="Direct activity project ID to send in payload",
    )
    parser.add_argument(
        "--workorder",
        default=DEFAULT_WORKORDER,
        help="Direct activity workorder to send in payload",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=1.5,
        help="Delay (seconds) between day submissions",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Build and validate payloads without submitting",
    )
    return parser.parse_args()


def create_robust_session():
    session = requests.Session()
    retries = Retry(
        total=5,
        connect=5,
        read=5,
        backoff_factor=1,
        status_forcelist=[500, 502, 503, 504],
        allowed_methods=frozenset(["GET", "POST"]),
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("http://", adapter)
    session.mount("https://", adapter)

    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Origin": "http://www.ntsp.nec.co.jp",
            "Upgrade-Insecure-Requests": "1",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        }
    )
    return session


def login(session, username, password):
    print("\n[1/3] Initiating login sequence...")

    login_page = session.get(LOGIN_PAGE_URL, timeout=30)
    login_page.raise_for_status()
    soup = BeautifulSoup(login_page.text, "html.parser")

    login_token_input = soup.find("input", {"name": "_csrf_token"})
    if not login_token_input or not login_token_input.get("value"):
        print("[ERROR] Login CSRF token not found on login page.")
        return False

    payload = {
        "_username": username,
        "_password": password,
        "_csrf_token": login_token_input.get("value"),
    }

    response = session.post(LOGIN_URL, data=payload, timeout=30)

    if "Invalid" in response.text or "Bad credentials" in response.text:
        print("[ERROR] Login failed: invalid credentials.")
        return False

    cookie_names = set(session.cookies.keys())
    expected = {"ASP.NET_SessionId", "PHPSESSID"}
    if not expected.issubset(cookie_names):
        # Some environments only issue one cookie or rename session cookies behind gateways.
        print(
            "[WARN] Login response did not include all expected session cookies "
            f"(got: {', '.join(sorted(cookie_names)) or 'none'}). Verifying session via protected page..."
        )

    probe_date = time.strftime("%Y-%m-%d")
    probe_url = f"{TIMESHEET_URL}?dtsdate={probe_date}"
    try:
        auth_probe = session.get(probe_url, timeout=30, allow_redirects=True)
    except requests.RequestException as exc:
        print(f"[WARN] Could not verify authenticated session via probe URL: {exc}")
        print("[WARN] Continuing; authentication will be validated again during per-date processing.")
        print("[OK] Login accepted (probe skipped due server/network instability).")
        return True

    if auth_probe.status_code >= 500:
        print(
            f"[WARN] Probe endpoint returned HTTP {auth_probe.status_code}. "
            "Continuing; server may be temporarily unstable."
        )
        print("[OK] Login accepted (probe not conclusive).")
        return True

    probe_text_lower = auth_probe.text.lower()
    if (
        "name=\"_username\"" in probe_text_lower
        or "name=\"_password\"" in probe_text_lower
        or "name=\"_csrf_token\"" in probe_text_lower
    ):
        print("[ERROR] Authentication did not persist (still seeing login form).")
        return False

    print("[OK] Login successful.")
    return True


def load_and_validate_excel(excel_path):
    print("\n[2/3] Reading and validating Excel data...")
    df = pd.read_excel(excel_path)

    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

    df = df[REQUIRED_COLUMNS].copy()

    parsed_dates = pd.to_datetime(df["Date"], errors="coerce")
    invalid_dates = df[parsed_dates.isna()]
    if not invalid_dates.empty:
        raise ValueError(
            "Invalid Date values found. Ensure dates are valid and can be formatted as YYYY-MM-DD."
        )
    df["Date"] = parsed_dates.dt.strftime("%Y-%m-%d")

    df["Regular Hours"] = pd.to_numeric(df["Regular Hours"], errors="coerce")
    df["OT Hours"] = pd.to_numeric(df["OT Hours"], errors="coerce").fillna(0.0)
    if df["Regular Hours"].isna().any():
        raise ValueError("Regular Hours contains non-numeric values.")

    df["Type"] = df["Type"].astype(str).str.strip().str.title()
    df["Activity Name"] = df["Activity Name"].astype(str).str.strip()

    grouped = list(df.sort_values("Date").groupby("Date", sort=True))
    print(f"Found {len(grouped)} unique day(s) to process.")
    return grouped


def is_timesheet_locked(page_html, soup):
    approved_flags = ["DTS_STATUS_APPROVED = ''", "DTS_STATUS_APPROVED === ''"]
    if any(flag in page_html for flag in approved_flags):
        return True
    return soup.find("button", id="oss_dtsbundle_timesheet_submit") is None


def build_daily_payload(group, csrf_token, target_date, project_id, workorder):
    payload = {
        "_method": "PUT",
        "oss_dtsbundle_timesheet[date]": target_date,
        "oss_dtsbundle_timesheet[submit]": "",
        "oss_dtsbundle_timesheet[_token]": csrf_token,
    }

    direct_idx = 1
    indirect_idx = 1
    mapped_regular_total = 0.0

    for _, row in group.iterrows():
        act_type = row["Type"]
        act_name = row["Activity Name"]
        reg_hours = float(row["Regular Hours"])
        ot_hours = float(row["OT Hours"])

        reg_hrs = f"{reg_hours:.2f}"
        ot_hrs = f"{ot_hours:.2f}"

        if act_type == "Direct":
            act_id = DIRECT_ACTIVITIES.get(act_name)
            if not act_id:
                print(f"  [WARN] Unknown direct activity '{act_name}'. Row skipped.")
                continue

            payload[
                f"oss_dtsbundle_timesheet[dtsDirectActivity][{direct_idx}][project]"
            ] = str(project_id)
            payload[
                f"oss_dtsbundle_timesheet[dtsDirectActivity][{direct_idx}][workorder]"
            ] = str(workorder)
            payload[
                f"oss_dtsbundle_timesheet[dtsDirectActivity][{direct_idx}][projectActivity]"
            ] = str(act_id)
            payload[
                f"oss_dtsbundle_timesheet[dtsDirectActivity][{direct_idx}][regularHours]"
            ] = reg_hrs
            payload[
                f"oss_dtsbundle_timesheet[dtsDirectActivity][{direct_idx}][OTHours]"
            ] = ot_hrs
            direct_idx += 1
            mapped_regular_total += reg_hours
            continue

        if act_type == "Indirect":
            mapped = INDIRECT_ACTIVITIES.get(act_name)
            if not mapped:
                print(f"  [WARN] Unknown indirect activity '{act_name}'. Row skipped.")
                continue

            payload[
                f"oss_dtsbundle_timesheet[dtsIndirectActivity][{indirect_idx}][activityCategory]"
            ] = str(mapped["category"])
            payload[
                f"oss_dtsbundle_timesheet[dtsIndirectActivity][{indirect_idx}][activity]"
            ] = str(mapped["activity"])
            payload[
                f"oss_dtsbundle_timesheet[dtsIndirectActivity][{indirect_idx}][regularHours]"
            ] = reg_hrs
            payload[
                f"oss_dtsbundle_timesheet[dtsIndirectActivity][{indirect_idx}][OTHours]"
            ] = ot_hrs
            indirect_idx += 1
            mapped_regular_total += reg_hours
            continue

        print(f"  [WARN] Unsupported Type '{act_type}' for activity '{act_name}'. Row skipped.")

    payload["oss_dtsbundle_timesheet[dailyHours]"] = f"{mapped_regular_total:.2f}"
    has_entries = direct_idx > 1 or indirect_idx > 1
    return payload, mapped_regular_total, has_entries


def submit_timesheets(session, grouped_data, project_id, workorder, delay, dry_run):
    print("\n[3/3] Submitting timesheets...")
    print("-" * 60)

    summary = {
        "submitted": 0,
        "skipped_locked": 0,
        "skipped_no_entries": 0,
        "failed": 0,
    }

    for target_date, group in grouped_data:
        print(f"\n[DATE] {target_date} ({len(group)} row(s))")
        page_url = f"{TIMESHEET_URL}?dtsdate={target_date}"

        try:
            page_response = session.get(page_url, timeout=30)
            page_response.raise_for_status()
        except requests.RequestException as exc:
            print(f"[ERROR] Failed to open timesheet page: {exc}")
            summary["failed"] += 1
            continue

        soup = BeautifulSoup(page_response.text, "html.parser")
        if is_timesheet_locked(page_response.text, soup):
            print("[SKIP] Timesheet is approved/locked.")
            summary["skipped_locked"] += 1
            continue

        token_input = soup.find("input", {"name": "oss_dtsbundle_timesheet[_token]"})
        if not token_input or not token_input.get("value"):
            print("[ERROR] CSRF token not found. Session may be expired.")
            summary["failed"] += 1
            continue

        payload, total_hours, has_entries = build_daily_payload(
            group,
            token_input.get("value"),
            target_date,
            project_id,
            workorder,
        )

        if not has_entries:
            print("[SKIP] No valid mapped entries remained after filtering.")
            summary["skipped_no_entries"] += 1
            continue

        if dry_run:
            print(
                f"[DRY-RUN] Payload ready with dailyHours={payload['oss_dtsbundle_timesheet[dailyHours]']}"
            )
            continue

        headers = {
            "Referer": page_url,
            "Content-Type": "application/x-www-form-urlencoded",
        }

        try:
            response = session.post(UPDATE_URL, data=payload, headers=headers, timeout=30)
            if response.status_code in (200, 302):
                print(f"[OK] Submitted {total_hours:.2f} regular hour(s).")
                summary["submitted"] += 1
            else:
                print(f"[ERROR] Submit failed with HTTP {response.status_code}.")
                summary["failed"] += 1
        except requests.RequestException as exc:
            print(f"[ERROR] Submit request failed: {exc}")
            summary["failed"] += 1

        if delay > 0:
            time.sleep(delay)

    print("\n" + "-" * 60)
    print("Run summary")
    print(f"  Submitted: {summary['submitted']}")
    print(f"  Skipped (locked): {summary['skipped_locked']}")
    print(f"  Skipped (no valid entries): {summary['skipped_no_entries']}")
    print(f"  Failed: {summary['failed']}")
    return summary


def main():
    args = parse_args()

    password = args.password
    if not password:
        password = getpass.getpass(prompt="Enter your HR portal password: ")

    session = create_robust_session()
    try:
        if not login(session, args.username, password):
            return 1

        grouped_data = load_and_validate_excel(args.excel)
        summary = submit_timesheets(
            session=session,
            grouped_data=grouped_data,
            project_id=args.project,
            workorder=args.workorder,
            delay=args.delay,
            dry_run=args.dry_run,
        )

        if summary["failed"] > 0:
            return 2
        return 0
    except FileNotFoundError:
        print(f"[ERROR] Excel file not found: {args.excel}")
        return 1
    except ValueError as exc:
        print(f"[ERROR] Data validation failed: {exc}")
        return 1
    except requests.RequestException as exc:
        print(f"[ERROR] Network failure: {exc}")
        return 1
    finally:
        session.close()


if __name__ == "__main__":
    sys.exit(main())