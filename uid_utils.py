import os
import sys
import ctypes
import time
import re
import hashlib
from datetime import datetime

# --- DYNAMIC GLOBAL PATH HOOKS ---
def _resolve_base_local_path():
    user_home = os.path.expanduser("~")
    default_path = os.path.join(user_home, "OneDrive", "Documents", "posters")
    candidate_paths = [
        os.environ.get("WILDPOSTING_POSTERS_ROOT"),
        os.environ.get("POSTERS_ROOT"),
        os.getcwd(),
        default_path,
    ]
    city_markers = ("philly", "dc", "boston")

    for candidate in candidate_paths:
        if not candidate:
            continue
        normalized = os.path.abspath(candidate)
        if not os.path.isdir(normalized):
            continue

        has_registry = os.path.exists(os.path.join(normalized, "Country_UID_Registry.xlsx"))
        has_city_dirs = all(os.path.isdir(os.path.join(normalized, city)) for city in city_markers)
        if has_registry or has_city_dirs or normalized == os.path.abspath(default_path):
            return normalized

    return os.path.abspath(default_path)


BASE_LOCAL_PATH = _resolve_base_local_path()
LOG_PATH = os.path.join(BASE_LOCAL_PATH, "UID_Automation_Log.txt")
REGISTRY_PATH = os.path.join(BASE_LOCAL_PATH, "Country_UID_Registry.xlsx")

def init_system_log():
    """Guarantees a clean text log document drops into the posters folder on line 1."""
    try:
        os.makedirs(BASE_LOCAL_PATH, exist_ok=True)
        with open(LOG_PATH, "w", encoding="utf-8") as f:
            f.write(f"============================================================\n")
            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] System Log Initialized.\n")
            f.write(f"Active Root Directory: {BASE_LOCAL_PATH}\n")
    except Exception as e:
        ctypes.windll.user32.MessageBoxW(0, f"Critical IO Error: Cannot write log file!\nError: {str(e)}", "Fatal Error", 0x10)

def log_message(message):
    """Appends live runtime progress steps to your text log ledger."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted_msg = f"[{timestamp}] {message}"
    print(formatted_msg)
    with open(LOG_PATH, "a", encoding="utf-8") as log_file:
        log_file.write(formatted_msg + "\n")

def show_popup(title, message, is_error=True):
    """Fires native Windows popup notification and response dialog boxes."""
    if is_error == "retry_cancel":
        response = ctypes.windll.user32.MessageBoxW(0, message, title, 0x10 | 0x05)
        return response == 4 # Returns True if user clicks Retry
    elif is_error:
        ctypes.windll.user32.MessageBoxW(0, message, title, 0x10 | 0x00)
    else:
        ctypes.windll.user32.MessageBoxW(0, message, title, 0x40 | 0x00)

def create_text_fingerprint(location_text):
    """Cleans street descriptions and text-hashes them into a solid 6-character string."""
    if not isinstance(location_text, str):
        return "UNKNOWN"
    clean = location_text.lower().strip()
    clean = re.sub(r'[^a-z0-9]', '', clean)
    return hashlib.md5(clean.encode('utf-8')).hexdigest()[:6].upper()

def extract_gps_from_url(url_string):
    """Regex engine that parses latitude/longitude coordinates out of pasted street view links."""
    if not isinstance(url_string, str) or not url_string:
        return None, None

    patterns = [
        r'(?:viewpoint=|@)(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)',
        r'[?&](?:ll|sll|query|destination)=\s*(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)',
        r'!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)',
    ]

    for pattern in patterns:
        match = re.search(pattern, url_string)
        if match:
            return match.group(1), match.group(2)

    return None, None

def ensure_file_is_free(file_path, display_name):
    """Loops and halts system processing if a targeted sheet is locked open by Excel."""
    if not os.path.exists(file_path):
        return True
    while True:
        try:
            f = open(file_path, 'a')
            f.close()
            return True
        except PermissionError:
            log_message(f"⚠️ Lock detected on file: {display_name}")
            user_choice = show_popup(
                "Excel File Locked!", 
                f"Close the {display_name} workbook idiot.\n\nClick 'Retry' once closed, or 'Cancel' to stop.",
                is_error="retry_cancel"
            )
            if user_choice:
                log_message(f"🔄 Retrying connection to {display_name}...")
                time.sleep(0.5)
            else:
                log_message("❌ Process aborted by user choice.")
                return False
