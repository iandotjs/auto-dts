# Auto-DTS 🚀

A Python-based command-line utility to automate the submission of Daily Time Sheets (DTS) to the company HR portal. 

By bypassing brittle DOM-based UI automation (like UIPath) and interacting directly with the backend HTTP API, this tool reads timesheet data from a flat Excel file, handles CSRF token authentication automatically, and submits entries in seconds.

## ✨ Features
* **Headless API Integration:** Submits standard Form Data directly to the server, completely ignoring slow page load times and UI rendering.
* **Smart Excel Parsing:** Reads a simple, flat Excel structure and dynamically groups multiple tasks per day into a single payload.
* **Auto-Authentication:** Uses `requests.Session()` to handle login cookies and uses `BeautifulSoup` to scrape fresh CSRF security tokens on the fly.
* **Safe-Skip Logic:** Automatically detects `DTS_STATUS_APPROVED` flags and skips dates that have already been locked or approved by management to prevent crashes.
* **Robust Error Handling:** Built-in automatic retry adapter with backoff factors for handling sudden 502/503 internal server errors.
* **Secure Input:** Uses `getpass` to prompt for passwords securely in the terminal without hardcoding credentials.

## 🛠️ Prerequisites
* Python 3.8+
* An active company HR portal account

## 📦 Installation

1. Clone the repository:
   ```bash
   git clone [https://github.com/iandotjs/auto-dts.git](https://github.com/iandotjs/auto-dts.git)
   cd auto-dts
   