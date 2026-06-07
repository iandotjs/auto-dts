import argparse
import getpass
import json
import os
import re
import sys
import time


def _load_dotenv(path=".env"):
    """Load key=value pairs from a .env file into os.environ (no-op if missing)."""
    try:
        with open(path, encoding="utf-8") as fp:
            for line in fp:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _, value = line.partition("=")
                key = key.strip()
                value = value.strip()
                if key and key not in os.environ:
                    os.environ[key] = value
    except FileNotFoundError:
        pass


_load_dotenv()

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
DEFAULT_USERNAME = os.getenv("AUTO_DTS_USERNAME", "dinopol.ij")
DEFAULT_PROJECT_ID = "484"
DEFAULT_WORKORDER = "S26-000-12V-00"
DIRECT_MAP_FILE = "direct_activities.json"
INDIRECT_MAP_FILE = "indirect_activities.json"

REQUIRED_COLUMNS = ["Date", "Type", "Activity Name", "Regular Hours", "OT Hours"]


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
    parser.add_argument(
        "--direct-map",
        default=DIRECT_MAP_FILE,
        help="Path to direct activity mapping JSON",
    )
    parser.add_argument(
        "--indirect-map",
        default=INDIRECT_MAP_FILE,
        help="Path to indirect activity mapping JSON",
    )
    parser.add_argument(
        "--validate-mappings",
        action="store_true",
        help="Validate mapping files and report unknown activities in Excel, then exit",
    )
    return parser.parse_args()


def normalize_activity_name(name):
    return re.sub(r"\s+", " ", str(name).strip())


def load_json(path, label):
    with open(path, "r", encoding="utf-8") as fp:
        try:
            data = json.load(fp)
        except json.JSONDecodeError as exc:
            raise ValueError(f"{label} JSON is invalid: {exc}") from exc

    if not isinstance(data, dict):
        raise ValueError(f"{label} must be a JSON object of activity-name keys.")
    return data


def validate_direct_mapping(raw_map):
    validated = {}
    for key, value in raw_map.items():
        normalized_key = normalize_activity_name(key)
        if not normalized_key:
            raise ValueError("DIRECT_ACTIVITIES contains an empty key.")
        if not isinstance(value, int):
            raise ValueError(
                f"DIRECT_ACTIVITIES value for '{key}' must be an integer, got {type(value).__name__}."
            )
        if normalized_key in validated and validated[normalized_key] != value:
            raise ValueError(
                f"DIRECT_ACTIVITIES has duplicate normalized key '{normalized_key}' with different IDs."
            )
        validated[normalized_key] = value
    return validated


def validate_indirect_mapping(raw_map):
    validated = {}
    for key, value in raw_map.items():
        normalized_key = normalize_activity_name(key)
        if not normalized_key:
            raise ValueError("INDIRECT_ACTIVITIES contains an empty key.")
        if not isinstance(value, dict):
            raise ValueError(
                f"INDIRECT_ACTIVITIES value for '{key}' must be an object with 'category' and 'activity'."
            )
        if "category" not in value or "activity" not in value:
            raise ValueError(
                f"INDIRECT_ACTIVITIES '{key}' must include both 'category' and 'activity'."
            )
        category = value["category"]
        activity = value["activity"]
        if not isinstance(category, int) or not isinstance(activity, int):
            raise ValueError(
                f"INDIRECT_ACTIVITIES '{key}' category/activity must be integers."
            )
        normalized_value = {"category": category, "activity": activity}
        if normalized_key in validated and validated[normalized_key] != normalized_value:
            raise ValueError(
                f"INDIRECT_ACTIVITIES has duplicate normalized key '{normalized_key}' with different values."
            )
        validated[normalized_key] = normalized_value
    return validated


def load_activity_mappings(direct_map_path, indirect_map_path):
    print("[MAP] Loading activity mappings...")
    direct_raw = load_json(direct_map_path, "Direct mapping")
    indirect_raw = load_json(indirect_map_path, "Indirect mapping")
    direct_activities = validate_direct_mapping(direct_raw)
    indirect_activities = validate_indirect_mapping(indirect_raw)
    print(
        "[MAP] Loaded "
        f"{len(direct_activities)} direct and {len(indirect_activities)} indirect activity mappings."
    )
    return direct_activities, indirect_activities


def create_robust_session():
    session = requests.Session()
    retries = Retry(
        total=5,
        connect=5,
        read=5,
        backoff_factor=1,
        status_forcelist=[502, 503, 504],
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
    login_token = login_token_input.get("value")

    payload = {
        "_username": username,
        "_password": password,
        "_csrf_token": login_token,
        "_remember_me": "on",
    }
    login_headers = {
        "Referer": LOGIN_PAGE_URL,
        "Content-Type": "application/x-www-form-urlencoded",
    }

    # Symfony returns 302 on success, 200 on failure.
    response = session.post(
        LOGIN_URL, data=payload, headers=login_headers, timeout=30, allow_redirects=False
    )

    if response.status_code not in (301, 302, 303, 307, 308):
        print("[ERROR] Login failed: server did not redirect after credential POST.")
        return False

    redirect_url = response.headers.get("Location", "")
    if not redirect_url.startswith("http"):
        redirect_url = BASE_URL.rstrip("/") + "/" + redirect_url.lstrip("/")

    # A redirect back to the login page means credentials were rejected.
    if "login" in redirect_url.lower():
        print(
            "[ERROR] Login failed: server redirected back to the login page. "
            "Check your username and password."
        )
        return False

    # Follow the success redirect so session cookies are established.
    session.get(redirect_url, timeout=30)
    print("[OK] Login successful.")
    return True


def load_and_validate_excel(excel_path):
    print("\n[2/3] Reading and validating Excel data...")
    df = pd.read_excel(excel_path)

    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

    df = df[REQUIRED_COLUMNS].copy()

    # Load optional Holiday column (defaults to False if missing)
    if "Holiday" not in df.columns:
        df["Holiday"] = False
    df["Holiday"] = df["Holiday"].fillna(False).astype(bool)

    parsed_dates = pd.to_datetime(df["Date"], errors="coerce")
    invalid_dates = df[parsed_dates.isna()]
    if not invalid_dates.empty:
        raise ValueError(
            "Invalid Date values found. Ensure dates are valid and can be formatted as YYYY-MM-DD."
        )
    df["Date"] = parsed_dates.dt.strftime("%Y-%m-%d")
    df["_parsed_date"] = parsed_dates

    df["Regular Hours"] = pd.to_numeric(df["Regular Hours"], errors="coerce")
    df["OT Hours"] = pd.to_numeric(df["OT Hours"], errors="coerce").fillna(0.0)
    if df["Regular Hours"].isna().any():
        raise ValueError("Regular Hours contains non-numeric values.")

    df["Type"] = df["Type"].astype(str).str.strip().str.title()
    df["Activity Name"] = df["Activity Name"].apply(normalize_activity_name)

    # Validate weekday hour requirements
    _validate_weekday_hours(df)

    grouped = list(df.sort_values("Date").groupby("Date", sort=True))
    print(f"Found {len(grouped)} unique day(s) to process.")
    if len(grouped) == 0:
        print("[WARN] No rows found to process in the Excel file.")
    return df, grouped


def _validate_weekday_hours(df):
    """Validate that weekdays (Mon-Fri) have >= 8 hours or are marked as holidays."""
    for date_str, group in df.groupby("Date"):
        # Get the first row's parsed date to determine day of week
        parsed_date = group["_parsed_date"].iloc[0]
        weekday = parsed_date.weekday()  # 0=Mon, 4=Fri, 5=Sat, 6=Sun
        is_holiday = group["Holiday"].iloc[0]
        total_regular = group["Regular Hours"].sum()

        # Only validate Mon-Fri (weekday < 5)
        if weekday < 5 and not is_holiday:
            if total_regular < 8.0:
                raise ValueError(
                    f"Date {date_str} is a weekday with only {total_regular:.2f} regular hours. "
                    "Weekdays require >= 8 hours or must be marked as Holiday=True."
                )


def validate_mappings_against_excel(df, direct_activities, indirect_activities):
    print("\n[CHECK] Validating Excel activities against mappings...")
    direct_unknown = sorted(
        {
            name
            for _, row in df[df["Type"] == "Direct"].iterrows()
            for name in [row["Activity Name"]]
            if name not in direct_activities
        }
    )
    indirect_unknown = sorted(
        {
            name
            for _, row in df[df["Type"] == "Indirect"].iterrows()
            for name in [row["Activity Name"]]
            if name not in indirect_activities
        }
    )

    if direct_unknown:
        print("[WARN] Unknown direct activity names found in Excel:")
        for name in direct_unknown:
            print(f"  - {name}")

    if indirect_unknown:
        print("[WARN] Unknown indirect activity names found in Excel:")
        for name in indirect_unknown:
            print(f"  - {name}")

    if not direct_unknown and not indirect_unknown:
        print("[OK] All Excel activities are covered by the mappings.")
        return True

    print("[CHECK] Mapping validation found gaps. Please update mapping JSON files.")
    return False


def is_timesheet_locked(page_html, soup):
    status_matches = dict(
        re.findall(r'var\s+(DTS_STATUS_[A-Z]+)\s*=\s*"([^"]*)"', page_html)
    )
    approved_status = status_matches.get("DTS_STATUS_APPROVED")
    submitted_status = status_matches.get("DTS_STATUS_SUBMITTED")
    saved_status = status_matches.get("DTS_STATUS_SAVE")

    # The page JS treats an empty string as the active state marker.
    # approved == ''   -> approved/locked
    # submitted == ''  -> submitted/non-editable
    # save == ''       -> editable draft
    if approved_status == "":
        return True
    if submitted_status == "":
        return True
    if saved_status == "":
        return False

    # If the status variables are missing, fall back to the presence of the submit action.
    return soup.find("button", id="oss_dtsbundle_timesheet_submit") is None


def load_timesheet_page_for_date(session, target_date):
    base_response = session.get(TIMESHEET_URL, timeout=30)
    base_response.raise_for_status()

    base_soup = BeautifulSoup(base_response.text, "html.parser")
    token_input = base_soup.find("input", {"name": "oss_dtsbundle_timesheet[_token]"})
    if not token_input or not token_input.get("value"):
        raise ValueError("Failed to load base timesheet page token.")

    payload = {
        "_method": "PUT",
        "oss_dtsbundle_timesheet[date]": target_date,
        "oss_dtsbundle_timesheet[dailyHours]": "",
        "oss_dtsbundle_timesheet[_token]": token_input.get("value"),
    }
    headers = {
        "Referer": TIMESHEET_URL,
        "Content-Type": "application/x-www-form-urlencoded",
    }
    return session.post(UPDATE_URL, data=payload, headers=headers, timeout=30)


def build_daily_payload(
    group,
    csrf_token,
    target_date,
    project_id,
    workorder,
    direct_activities,
    indirect_activities,
):
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
            act_id = direct_activities.get(act_name)
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
            mapped = indirect_activities.get(act_name)
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


def submit_timesheets(
    session,
    grouped_data,
    project_id,
    workorder,
    delay,
    dry_run,
    direct_activities,
    indirect_activities,
):
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
        try:
            page_response = load_timesheet_page_for_date(session, target_date)
        except (requests.RequestException, ValueError) as exc:
            print(f"[ERROR] Failed to open timesheet page: {exc}")
            summary["failed"] += 1
            continue

        if page_response.status_code >= 400:
            print(
                f"[ERROR] Timesheet page returned HTTP {page_response.status_code}. "
                "Session may be expired or server is rejecting the request."
            )
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
            direct_activities,
            indirect_activities,
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
            "Referer": TIMESHEET_URL,
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

    direct_activities, indirect_activities = load_activity_mappings(
        args.direct_map, args.indirect_map
    )

    df, grouped_data = load_and_validate_excel(args.excel)
    if args.validate_mappings:
        is_valid = validate_mappings_against_excel(df, direct_activities, indirect_activities)
        return 0 if is_valid else 2

    password = args.password
    if not password:
        password = getpass.getpass(prompt="Enter your HR portal password: ")

    session = create_robust_session()
    try:
        if not login(session, args.username, password):
            return 1

        summary = submit_timesheets(
            session=session,
            grouped_data=grouped_data,
            project_id=args.project,
            workorder=args.workorder,
            delay=args.delay,
            dry_run=args.dry_run,
            direct_activities=direct_activities,
            indirect_activities=indirect_activities,
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