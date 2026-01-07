# =====================================================
# INAPROC PRODUCT NAME BOT
# FINAL VERSION (AUTO INSTALL REQUIREMENTS ADDED)
# =====================================================

BOT_VERSION = "1.1.1"

VERSION_URL = "https://raw.githubusercontent.com/Alnpan132/INAPROC-Check-Name-Bot/main/version.txt"
CH_PY_URL   = "https://raw.githubusercontent.com/Alnpan132/INAPROC-Check-Name-Bot/main/ch.PY"


import os, csv, sys, time, urllib.request, itertools, re, json, ctypes

# ================= EARLY AUTO INSTALL =================
def _early_ensure_requirements():
    marker = ".requirements_installed"
    if os.path.exists(marker):
        return

    required = ["selenium", "keyboard"]
    missing = []

    for pkg in required:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)

    if not missing:
        open(marker, "w").close()
        return

    print("\n[SETUP] Installing required packages...")
    for pkg in missing:
        print(f"Installing {pkg}...")
        os.system(f'"{sys.executable}" -m pip install {pkg}')

    open(marker, "w").close()
    print("[SETUP] Done. Restarting...\n")
    time.sleep(1)
    os.execv(sys.executable, [sys.executable] + sys.argv)

_early_ensure_requirements()

# ================= RUNTIME IMPORTS =================
import keyboard
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
# ================= CONSOLE TITLE =================
try:
    ctypes.windll.kernel32.SetConsoleTitleW("INAPROC PRODUCT NAME BOT")
except:
    pass

# ================= COLORS =================
RED="\033[91m"; GREEN="\033[92m"; YELLOW="\033[93m"; CYAN="\033[96m"; RESET="\033[0m"

# ================= VERSION CHECK =================
def check_for_update():
    try:
        print("Checking for update...", end="", flush=True)
        time.sleep(0.6)

        latest = urllib.request.urlopen(VERSION_URL, timeout=3).read().decode().strip()

        # clear the line
        print("\r" + " " * 40 + "\r", end="")

        if latest == BOT_VERSION:
            return

        print(f"{YELLOW}[UPDATE AVAILABLE]{RESET}")
        print(f"Current version : {BOT_VERSION}")
        print(f"Latest version  : {latest}")
        choice = input("Install update now? (Y/N) : ").strip().lower()

        if choice != "y":
            return

        print("\nDownloading update...")
        tmp = "ch.PY.new"
        data = urllib.request.urlopen(CH_PY_URL, timeout=5).read()

        with open(tmp, "wb") as f:
            f.write(data)

        os.replace(tmp, "ch.PY")

        print("Update installed successfully.")
        print("Restarting bot...\n")
        time.sleep(1)

        os.execv(sys.executable, [sys.executable] + sys.argv)

    except Exception:
        # silently ignore ALL update errors (404, offline, etc)
        print("\r" + " " * 40 + "\r", end="")
        return


# ================= CONFIG =================
DEFAULT_CHROME_PORT = 9555
CURRENT_CHROME_PORT = DEFAULT_CHROME_PORT

ERROR_DUP_TEXT="Nama Produk tidak boleh sama"
HEADER_SCAN_MAX_ROWS=10
MAX_FIX_TRY=20
CHECK_TIMEOUT=30
CSV_WRITE_RETRY=3

SCAN_DELAY = 0.5
ERROR_RETRY_DELAY = 0.7

SPINNER=itertools.cycle(["|","/","-","\\"])
browser=None

DATASOURCE_DIR = "DataSource"

# ================= HOTKEY FLAGS =================
PAUSE_SCAN=False
ABORT_TO_MENU=False
FORCE_SUMMARY=False

# ================= HOTKEY HANDLERS =================
def register_hotkeys():
    keyboard.add_hotkey("ctrl+down", toggle_pause)
    keyboard.add_hotkey("ctrl+left", abort_scan)
    keyboard.add_hotkey("ctrl+right", force_summary)

def toggle_pause():
    if not is_bot_foreground():
        return
    global PAUSE_SCAN
    if not PAUSE_SCAN:
        PAUSE_SCAN = True
        print(f"\n{YELLOW}[PAUSE REQUESTED]{RESET}")


def abort_scan():
    if not is_bot_foreground():
        return
    global ABORT_TO_MENU
    ABORT_TO_MENU = True

def force_summary():
    if not is_bot_foreground():
        return
    global FORCE_SUMMARY
    FORCE_SUMMARY = True

# ================= UTIL =================
def clear():
    os.system("cls")
# ================= MENU BOX UTIL =================
def box_border(width=34):
    return f"+{'-' * (width + 2)}+"

def box_line(text="", width=34):
    return f"| {text.ljust(width)} |"

def csv_files():
    if not os.path.isdir(DATASOURCE_DIR):
        return []
    return [
        f for f in os.listdir(DATASOURCE_DIR)
        if f.lower().endswith(".csv") and not f.startswith("_")
    ]


def select_csv_file():
    files=csv_files()
    if not files:
        input("No CSV found. Press ENTER...")
        return None

    clear()
    print(f"{CYAN}+------------------------------+{RESET}")
    print(f"{CYAN}|        SELECT CSV FILE       |{RESET}")
    print(f"{CYAN}+------------------------------+{RESET}")
    for i,f in enumerate(files,1):
        print(f"| [{i}] {f:<26}")
    print(f"{CYAN}+------------------------------+{RESET}")

    try:
        return os.path.join(DATASOURCE_DIR, files[int(input("Select number : "))-1])
    except:
        return None
def is_bot_foreground():
    try:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
        buff = ctypes.create_unicode_buffer(length + 1)
        ctypes.windll.user32.GetWindowTextW(hwnd, buff, length + 1)
        title = buff.value.lower()
        return "inaproc product name bot" in title
    except:
        return False


# ================= SAFE CSV =================

def read_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.reader(f))


def write_csv(path, rows):
    for _ in range(CSV_WRITE_RETRY):
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                csv.writer(f).writerows(rows)
            return True
        except PermissionError:
            print(f"{RED}[ERROR]{RESET} CSV file is locked. Please close Excel.")
            time.sleep(2)
    return False


def is_csv_locked(path):
    """
    Check if CSV file is currently locked (usually opened by Excel).
    This does NOT modify the file.
    """
    try:
        with open(path, "a"):
            pass
        return False
    except PermissionError:
        return True
    except:
        return False


# ================= COLUMN =================
def find_name_column(rows):
    for r in range(min(HEADER_SCAN_MAX_ROWS,len(rows))):
        for c,v in enumerate(rows[r]):
            if v.strip().lower() in ("nama produk","nama"):
                return r,c
    raise RuntimeError("Nama Produk column not found")

def find_status_column(rows,hr):
    header=rows[hr]
    for i,h in enumerate(header):
        if h.strip().upper()=="STATUS":
            return i
    header.append("STATUS")
    for r in rows[hr+1:]:
        r.append("")
    return len(header)-1

# ================= BRAIN CORE =================
def split_product_and_series(name):
    tokens=name.split()
    for i,t in enumerate(tokens):
        if re.search(r"[A-Za-z]",t) and re.search(r"\d",t):
            return " ".join(tokens[:i])," ".join(tokens[i:])
    return name,""

# ================= VARIANTS =================
def generate_variants(original):
    product,series=split_product_and_series(original)

    if not series:
        base=original.strip().title()
        words=base.split()
        variants=[base+".", base+" -"]

        if len(words)==2:
            swapped=f"{words[1]} {words[0]}"
            variants.append(swapped+".")
            variants.append(swapped+" -")
            variants.append(f"{words[0]}-{words[1]}")

        elif len(words)>=3:
            swapped=" ".join([words[1],words[0]]+words[2:])
            variants.append(swapped+".")
            variants.append(swapped+" -")

        return variants[:MAX_FIX_TRY]

    base=series.strip()
    has_paren=base.startswith("(") and base.endswith(")")
    plain=base[1:-1] if has_paren else base

    series_vars=[
        plain.replace("-"," "),
        plain.replace("-","_"),
        plain,
        plain.replace("-","").replace("_",""),
        re.sub(r"([A-Za-z])(\d)",r"\1-\2",plain)
    ]

    final=[]
    for s in series_vars:
        final.append(s)
        final.append(f"({s})")

    names=[]
    for s in final:
        full=f"{product} {s}".strip()
        if full!=original and full not in names:
            names.append(full)

    spec_variants=[]
    if any(x in plain.lower() for x in ["/",":","mm","cm","kg","hz"]):
        spec_variants.append(f"{product} {plain}")
        spec_variants.append(f"{product} - {plain}")

    all_variants=names+spec_variants
    all_variants+=[n+"." for n in all_variants]
    all_variants+=[n+" -" for n in all_variants]

    return all_variants[:MAX_FIX_TRY]

# ================= BROWSER =================

import socket

COMMON_CHROME_PORTS = [9222, 9333, 9444, 9555, 9666]

def scan_open_ports(ports, timeout=0.15):
    result = []
    for port in ports:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=timeout):
                result.append(port)
        except:
            pass
    return result


def reconnect_chrome_menu():
    global browser, CURRENT_CHROME_PORT

    WIDTH = 34

    while True:

        # ⚡ FAST PATH — only auto-return if already connected
        if browser:
            try:
                browser.title
                return  # already connected → back to main menu
            except:
                browser = None  # dead session → continue to menu

        clear()

        print(f"{CYAN}{box_border(WIDTH)}{RESET}")
        print(f"{CYAN}{box_line('CONNECTING TO CHROME', WIDTH)}{RESET}")
        print(f"{CYAN}{box_border(WIDTH)}{RESET}")

        print("| Current Chrome Port : ", end="")
        if browser is None:
            print(f"{RED}Not Connected{RESET}")
        else:
            print(f"{GREEN}{CURRENT_CHROME_PORT}{RESET}")

        print(f"{CYAN}{box_border(WIDTH)}{RESET}")

        open_ports = scan_open_ports(COMMON_CHROME_PORTS)

        if open_ports:
            print("| Common ports detected:")
            for i, p in enumerate(open_ports, 1):
                print(f"| {GREEN}[{i}]{RESET} {p}")
        else:
            print("| No common ports detected.")

        print(f"{CYAN}{box_border(WIDTH)}{RESET}")

        print(f"| {GREEN}[M]{RESET} Manual input port")
        print(f"| {RED}[0]{RESET} Return to Main Menu")
        print(f"{CYAN}{box_border(WIDTH)}{RESET}")

        choice = input("Select option : ").strip().lower()

        if choice.isdigit() and open_ports:
            idx = int(choice) - 1
            if 0 <= idx < len(open_ports):
                CURRENT_CHROME_PORT = open_ports[idx]
                try:
                    opt = Options()
                    opt.debugger_address = f"127.0.0.1:{CURRENT_CHROME_PORT}"
                    browser = webdriver.Chrome(options=opt)
                    return
                except:
                    print(f"\n{RED}[FAILED]{RESET} Cannot connect.")
                    input("Press ENTER...")
                    continue

        if choice == "m":
            p = input("Enter Chrome port : ").strip()
            if p.isdigit():
                CURRENT_CHROME_PORT = int(p)
                try:
                    opt = Options()
                    opt.debugger_address = f"127.0.0.1:{CURRENT_CHROME_PORT}"
                    browser = webdriver.Chrome(options=opt)
                    return
                except:
                    print(f"\n{RED}[FAILED]{RESET} Cannot connect.")
                    input("Press ENTER...")
                    continue

        if choice == "0":
            return


# ================= CHECK =================
def check_name(name, label):
    start = time.monotonic()

    field = WebDriverWait(browser, 15).until(
        EC.presence_of_element_located((By.XPATH, "//input[@name='name']"))
    )
    field.send_keys(Keys.CONTROL, "a", Keys.BACKSPACE)
    # --- ChromeDriver BMP safety ---
    safe_name = name.encode("utf-16", "surrogatepass").decode("utf-16", "ignore")
    field.send_keys(safe_name)

    while True:

        # 🔴 SAFE PAUSE CHECK (NEW FEATURE, NO SPEED CHANGE)
        if PAUSE_SCAN:
            return "PAUSED"

        if time.monotonic() - start > CHECK_TIMEOUT:
            return "ADA"

        sys.stdout.write(f"\r=> {label} {next(SPINNER)}")
        sys.stdout.flush()

        # ✅ RESTORED ORIGINAL SCAN SPEED
        time.sleep(SCAN_DELAY)

        try:
            browser.find_element(
                By.XPATH, f"//*[contains(text(),'{ERROR_DUP_TEXT}')]"
            )
            sys.stdout.write("\r" * 80)
            return "ADA"
        except:
            pass

        try:
            browser.find_element(
                By.XPATH, "//div[contains(@class,'Spinner')]"
            )
        except:
            sys.stdout.write("\r" * 80)
            return "TIDAK ADA"

# ================= PROCESS / SUMMARY / MAIN MENU =================
# (UNCHANGED — YOUR ORIGINAL LOGIC)

def process_csv(path, mode="FULL"):
    global PAUSE_SCAN, ABORT_TO_MENU, FORCE_SUMMARY
    PAUSE_SCAN = ABORT_TO_MENU = FORCE_SUMMARY = False

    # ===== CSV LOCK CHECK =====
    if is_csv_locked(path):
        clear()
        print(f"{RED}[ERROR]{RESET} CSV file is currently open or locked.")
        print("Please close the file in Excel and try again.")
        input("\nPress ENTER to return to menu...")
        return

    rows = read_csv(path)

    hr, ci = find_name_column(rows)
    si = find_status_column(rows, hr)

    targets = (
        [i for i in range(hr+1, len(rows)) if rows[i][si] != "TIDAK ADA"]
        if mode != "ERROR_ONLY"
        else [i for i in range(hr+1, len(rows)) if rows[i][si] == "ERROR"]
    )

    clear()
    print(f"{CYAN}{'RETRY ERROR LINES ONLY' if mode=='ERROR_ONLY' else 'FULL CSV SCAN'}{RESET}")

    valid = fixed = error = scanned = 0

    idx = 0
    while idx < len(targets):

        i = targets[idx]

        if ABORT_TO_MENU:
            return
        if FORCE_SUMMARY:
            break

        # ===== PAUSE MENU =====
        while PAUSE_SCAN:
            print(f"\n{YELLOW}[PAUSED]{RESET}")
            print("----------------------------------")
            print("[1] Resume scanning")
            print("[2] Return to main menu")
            print("----------------------------------")
            c = input("Select option : ").strip()
            if c == "1":
                PAUSE_SCAN = False
                print(f"{GREEN}[RESUMED]{RESET}")
            elif c == "2":
                return

        name = rows[i][ci].strip()

        print(f"\nScanning line {idx+1} / {len(targets)}")
        print(f"Product : {name}")
        scanned += 1   # ✅ CORRECT POSITION (after pause, before processing)

        # ===== CHECK ORIGINAL NAME =====
        result = check_name(name, "Checking")
        if result == "PAUSED":
            continue

        if result == "TIDAK ADA":
            rows[i][si] = "TIDAK ADA"
            valid += 1
            write_csv(path, rows)
            print(f"=> Result : {GREEN}TIDAK ADA{RESET}")
            idx += 1
            continue

        print(f"=> Result : {YELLOW}ADA{RESET}")
        print(f"{YELLOW}[FIX] Fixing name{RESET}")

        fixed_this_line = False

        # ===== TRY VARIANTS =====
        for v in generate_variants(name):
            if mode == "ERROR_ONLY":
                time.sleep(ERROR_RETRY_DELAY)

            result = check_name(v, "Trying")
            if result == "PAUSED":
                break

            if result == "TIDAK ADA":
                rows[i][ci] = v
                rows[i][si] = "TIDAK ADA"
                fixed += 1
                fixed_this_line = True
                write_csv(path, rows)
                print(f"{YELLOW}[{RESET}{GREEN}OK{RESET}{YELLOW}] FIXED -> {v}{RESET}")
                break

        if not fixed_this_line and not PAUSE_SCAN:
            rows[i][si] = "ERROR"
            error += 1
            write_csv(path, rows)
            print(f"{RED}[X] FAILED -> ERROR{RESET}")
            idx += 1
            continue

        if fixed_this_line:
            idx += 1

    summary_menu(path, valid, fixed, error, scanned)
    
def summary_menu(path, valid, fixed, error, total):
    filename = os.path.basename(path)
    m = re.search(r"(\d+)(?=\.csv$)", filename.lower())
    csv_number = m.group(1) if m else "UNKNOWN"

    while True:
        clear()
        print(f"{CYAN}+----------------------------------+{RESET}")
        title = f"SUMMARY {csv_number}"
        print(f"{CYAN}|{title.center(34)}|{RESET}")
        print(f"{CYAN}+----------------------------------+{RESET}")
        print(f"| TIDAK ADA : {GREEN}{valid:>5}{RESET}")
        print(f"| FIXED     : {YELLOW}{fixed:>5}{RESET}")
        print(f"| ERROR     : {RED}{error:>5}{RESET}")
        print(f"{CYAN}+----------------------------------+{RESET}")
        print(f"| TOTAL     : {total:>5}")
        print(f"{CYAN}+----------------------------------+{RESET}")
        print()
        print(f"{GREEN}[1]{RESET} Back to Main Menu")
        print(f"{GREEN}[2]{RESET} Open CSV File")
        print(f"{RED}[0]{RESET} Exit")

        c = input("Select option : ").strip()

        # ===== BACK TO MAIN MENU (MARK CSV AS DONE) =====
        if c == "1":
            try:
                folder, name = os.path.split(path)
                if not name.startswith("_"):
                    os.rename(path, os.path.join(folder, "_" + name))
            except:
                pass
            return

        # ===== OPEN CSV =====
        elif c == "2":
            try:
                os.startfile(path)
            except:
                print(f"{RED}[ERROR]{RESET} Unable to open CSV file.")
                input("Press ENTER...")

        # ===== EXIT PROGRAM =====
        elif c == "0":
            sys.exit(0)

# ================= MAIN MENU =================

def main_menu():
    global browser

    while True:
        clear()

        print(f"{CYAN}+----------------------------------+{RESET}")
        print(f"{CYAN}|  INAPROC PRODUCT NAME BOT         |{RESET}")
        print(f"{CYAN}+----------------------------------+{RESET}")

        print("| Current Chrome Port : ", end="")
        if browser is None:
            print(f"{RED}Not Connected{RESET}")
        else:
            print(f"{GREEN}{CURRENT_CHROME_PORT}{RESET}")

        print(f"| CSV Files           : {len(csv_files())}")

        print(f"{CYAN}+----------------------------------+{RESET}")

        print(f"| {GREEN}[1]{RESET} Full CSV Scan                 |")
        print(f"| {GREEN}[2]{RESET} Show CSV Files                |")
        print(f"| {GREEN}[3]{RESET} Reconnect Chrome              |")
        print(f"| {RED}[0]{RESET} Exit                          |")

        print(f"{CYAN}+----------------------------------+{RESET}")

        opt = input("Select option : ").strip()

        if opt == "1":
            path = select_csv_file()
            if not path:
                continue
            if browser is None:
                print(f"\n{RED}[ERROR]{RESET} Chrome is not connected.")
                print("Please use [Reconnect Chrome] first.")
                input("Press ENTER to return to menu...")
                continue
            process_csv(path, "FULL")

        elif opt == "2":
            clear()
            files = csv_files()
            if files:
                for f in files:
                    print(f)
            else:
                print("No CSV files found.")
            input("\nPress ENTER...")

        elif opt == "3":
            reconnect_chrome_menu()

        elif opt == "0":
            sys.exit(0)


# ================= ENTRY =================
if __name__ == "__main__":
    check_for_update()
    register_hotkeys()
    main_menu()

