import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup
import getpass
import time

# --- Configurations ---
EXCEL_FILE = 'timesheet_data.xlsx'
LOGIN_URL = 'http://www.ntsp.nec.co.jp/projects/web/login_check' 
LOGIN_PAGE_URL = 'http://www.ntsp.nec.co.jp/projects/web/login' # The page with the login form
TIMESHEET_URL = 'http://www.ntsp.nec.co.jp/projects/web/dtstimesheet/'
UPDATE_URL = 'http://www.ntsp.nec.co.jp/projects/web/dtstimesheet/update'

USERNAME = 'dinopol.ij'

# --- Paste your DIRECT_ACTIVITIES and INDIRECT_ACTIVITIES dictionaries here ---
DIRECT_ACTIVITIES = {
    "Coding > [Tool] Source Code Modification": 42398,
    "Testing > UT Execution": 42383,
    # ... (add the rest from the previous step)
}

INDIRECT_ACTIVITIES = {
    "Progress meeting": {"category": 1, "activity": 7},
    "Company or group activities": {"category": 32, "activity": 92},
    # ... (add the rest from the previous step)
}

def create_robust_session():
    """Creates a session that automatically retries on 502/503 Server Errors."""
    session = requests.Session()
    
    # Setup automatic retries: 5 attempts, backoff factor spaces them out (1s, 2s, 4s, etc.)
    retries = Retry(total=5, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
    session.mount('http://', HTTPAdapter(max_retries=retries))
    session.mount('https://', HTTPAdapter(max_retries=retries))
    
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Origin': 'http://www.ntsp.nec.co.jp',
        'Upgrade-Insecure-Requests': '1'
    })
    return session

def login(session):
    print("\n[1/3] Initiating Login Sequence...")
    password = getpass.getpass(prompt='Enter your HR portal password: ')
    
    # Step 1: GET login page to scrape the CSRF token for the login form
    login_page = session.get(LOGIN_PAGE_URL)
    soup = BeautifulSoup(login_page.text, 'html.parser')
    
    # Try to find the Symfony login CSRF token (usually named _csrf_token)
    login_token_input = soup.find('input', {'name': '_csrf_token'})
    login_token = login_token_input.get('value') if login_token_input else ""
    
    payload = {
        '_username': USERNAME,
        '_password': password,
        '_csrf_token': login_token
    }
    
    # Step 2: POST credentials
    response = session.post(LOGIN_URL, data=payload)
    
    if "Invalid" in response.text or "Bad credentials" in response.text:
        print("❌ Login failed: Invalid credentials.")
        return False
        
    print("✅ Login successful!")
    return True

def process_timesheets(session):
    print("\n[2/3] Reading and formatting Excel Data...")
    df = pd.read_excel(EXCEL_FILE)
    
    # Force strict formatting: Dates to YYYY-MM-DD, replace blank OT hours with 0
    df['Date'] = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d')
    df['OT Hours'] = df['OT Hours'].fillna(0)
    
    # Group the rows by Date so we can process one day at a time
    grouped = df.groupby('Date')
    print(f"Found {len(grouped)} unique days to process.\n")
    
    print("-" * 40)
    print("[3/3] Submitting Timesheets...")
    
    for target_date, group in grouped:
        print(f"\n📅 Processing Date: {target_date} ({len(group)} activities)")
        
        # 1. Fetch the timesheet page for this date
        page_url = f"{TIMESHEET_URL}?dtsdate={target_date}"
        page_response = session.get(page_url)
        soup = BeautifulSoup(page_response.text, 'html.parser')
        
        # 2. Check for Approval/Lock flags
        # The JS renders DTS_STATUS_APPROVED = '' when it IS approved (disabling the buttons)
        # We also check if the submit button literally doesn't exist in the HTML as a fallback
        if "DTS_STATUS_APPROVED === ''" in page_response.text or soup.find('button', id='oss_dtsbundle_timesheet_submit') is None:
            print(f"⚠️ SKIPPING: Timesheet for {target_date} is already approved or locked.")
            continue
            
        # 3. Extract the Timesheet CSRF token
        token_input = soup.find('input', {'name': 'oss_dtsbundle_timesheet[_token]'})
        if not token_input:
            print(f"❌ Failed to find CSRF token for {target_date}. Session might be expired.")
            continue
            
        csrf_token = token_input.get('value')
        
        # Calculate Total Daily Hours for the payload
        daily_total_hours = sum(group['Regular Hours'])
        
        # 4. Initialize Payload Base
        payload = {
            "_method": "PUT",
            "oss_dtsbundle_timesheet[date]": target_date,
            "oss_dtsbundle_timesheet[dailyHours]": "{:.2f}".format(daily_total_hours),
            "oss_dtsbundle_timesheet[submit]": "",
            "oss_dtsbundle_timesheet[_token]": csrf_token
        }
        
        # 5. Dynamically attach Direct and Indirect Activities
        direct_idx = 1
        indirect_idx = 1
        
        for _, row in group.iterrows():
            act_type = str(row['Type']).strip().title()
            act_name = str(row['Activity Name']).strip()
            reg_hrs = "{:.2f}".format(row['Regular Hours'])
            ot_hrs = "{:.2f}".format(row['OT Hours'])
            
            if act_type == 'Direct':
                act_id = DIRECT_ACTIVITIES.get(act_name)
                if not act_id:
                    print(f"  [!] Warning: Direct activity '{act_name}' not found in mapping. Skipping row.")
                    continue
                    
                payload[f"oss_dtsbundle_timesheet[dtsDirectActivity][{direct_idx}][project]"] = '484'
                payload[f"oss_dtsbundle_timesheet[dtsDirectActivity][{direct_idx}][workorder]"] = 'S26-000-12V-00'
                payload[f"oss_dtsbundle_timesheet[dtsDirectActivity][{direct_idx}][projectActivity]"] = act_id
                payload[f"oss_dtsbundle_timesheet[dtsDirectActivity][{direct_idx}][regularHours]"] = reg_hrs
                payload[f"oss_dtsbundle_timesheet[dtsDirectActivity][{direct_idx}][OTHours]"] = ot_hrs
                direct_idx += 1
                
            elif act_type == 'Indirect':
                mapped = INDIRECT_ACTIVITIES.get(act_name)
                if not mapped:
                    print(f"  [!] Warning: Indirect activity '{act_name}' not found in mapping. Skipping row.")
                    continue
                    
                payload[f"oss_dtsbundle_timesheet[dtsIndirectActivity][{indirect_idx}][activityCategory]"] = mapped['category']
                payload[f"oss_dtsbundle_timesheet[dtsIndirectActivity][{indirect_idx}][activity]"] = mapped['activity']
                payload[f"oss_dtsbundle_timesheet[dtsIndirectActivity][{indirect_idx}][regularHours]"] = reg_hrs
                payload[f"oss_dtsbundle_timesheet[dtsIndirectActivity][{indirect_idx}][OTHours]"] = ot_hrs
                indirect_idx += 1
                
        # 6. Fire the Payload
        try:
            submit_headers = {'Referer': page_url}
            response = session.post(UPDATE_URL, data=payload, headers=submit_headers)
            
            if response.status_code in [200, 302]:
                print(f"✅ SUCCESS: Logged {daily_total_hours} hrs for {target_date}")
            else:
                print(f"❌ FAILED: Status {response.status_code} for {target_date}")
                
        except Exception as e:
            print(f"❌ ERROR: Request failed for {target_date}. Details: {e}")
            
        time.sleep(1.5) # Polite buffer between days

    print("-" * 40)
    print("🚀 Automation Complete!")

if __name__ == "__main__":
    http_session = create_robust_session()
    if login(http_session):
        process_timesheets(http_session)