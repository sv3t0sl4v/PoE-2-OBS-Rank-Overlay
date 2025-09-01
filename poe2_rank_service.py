import servicemanager
import win32serviceutil
import win32service
import win32event
import time
import random
import psutil
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Config
ACCOUNT_NAME = 'XXX#0000'
LEAGUE_URL = 'https://pathofexile2.com/ladder/HC%2520SSF%2520Rise%2520of%2520the%2520Abyssal'
OUTPUT_FILE = r'E:\scripts\rank.txt'
SHOW_RIP = 1  # 1 = show RIP section, 0 = hide it

def is_obs_running():
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'] and 'obs' in proc.info['name'].lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return False

def check_rank():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get(LEAGUE_URL)

        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr"))
        )

        rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
        alive_chars, dead_chars = [], []

        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) < 4:
                continue

            rank_text = cols[0].text.strip()
            account = cols[1].text.strip()
            character = cols[2].text.strip()
            char_class = cols[3].text.strip()

            try:
                rank = int(rank_text)
            except ValueError:
                continue

            entry = {
                "rank": rank,
                "account": account.lower(),
                "character": character.lower(),
                "character_raw": character,
                "class": char_class
            }

            if "(Dead)" in character or "(Retired)" in character:
                dead_chars.append(entry)
            else:
                alive_chars.append(entry)

        # Find my alive chars
        account_alive = [c for c in alive_chars if c["account"] == ACCOUNT_NAME.lower()]
        account_dead = [c for c in dead_chars if c["account"] == ACCOUNT_NAME.lower()]

        if not account_alive:
            output_text = "HC SSF Global Rank: N/A\nClass Rank: N/A"
        else:
            top_char = min(account_alive, key=lambda c: c["rank"])
            global_rank = f"#{top_char['rank']}"

            # Compute alive class ranks
            class_ranks = []
            for char in sorted(account_alive, key=lambda c: c["rank"]):
                same_class = sorted([c for c in alive_chars if c["class"] == char["class"]], key=lambda c: c["rank"])
                for i, c in enumerate(same_class):
                    if c["rank"] == char["rank"]:
                        class_ranks.append(f"#{i + 1} {char['class'].title()}")
                        break

            lines = []
            if len(class_ranks) <= 3:
                lines.append("Class Rank: " + ", ".join(class_ranks))
            else:
                first_line = "Class Rank: " + ", ".join(class_ranks[:2]) + ","
                lines.append(first_line)
                for i in range(2, len(class_ranks), 3):
                    lines.append(", ".join(class_ranks[i:i+3]))

            output_text = f"HC SSF Global Rank: {global_rank}\n" + "\n".join(lines)

        # Compute RIP class ranks (only if enabled)
        if SHOW_RIP and account_dead:
            rip_ranks = []
            for char in sorted(account_dead, key=lambda c: c["rank"]):
                same_class = sorted([c for c in dead_chars if c["class"] == char["class"]], key=lambda c: c["rank"])
                for i, c in enumerate(same_class):
                    if c["rank"] == char["rank"]:
                        rip_ranks.append(f"#{i + 1} {char['class'].title()}")
                        break

            if rip_ranks:
                output_text += "\nRIP: " + ", ".join(rip_ranks)

        # Write to file
        with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
            f.write(output_text)

        servicemanager.LogInfoMsg(f"Wrote to {OUTPUT_FILE}:\n{output_text}")

    finally:
        driver.quit()

class RankService(win32serviceutil.ServiceFramework):
    _svc_name_ = "PoE2RankChecker"
    _svc_display_name_ = "PoE2 Ladder Rank Checker Service"
    _svc_description_ = "Checks Path of Exile 2 ladder rank while OBS is running, updates file regularly."

    def __init__(self, args):
        super().__init__(args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.is_running = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.is_running = False
        win32event.SetEvent(self.stop_event)

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def main(self):
        while self.is_running:
            if is_obs_running():
                try:
                    check_rank()
                except Exception as e:
                    servicemanager.LogErrorMsg(f"Rank check error: {e}")
                wait_time = random.randint(60, 300)
                time.sleep(wait_time)
            else:
                time.sleep(30)

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(RankService)
