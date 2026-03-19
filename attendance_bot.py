import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import json
import os
import time
import datetime
import webbrowser
import subprocess
import sys
import re

# ── Optional imports ──────────────────────────────────────────────────────────
try:
    import speech_recognition as sr
    SR_AVAILABLE = True
except ImportError:
    SR_AVAILABLE = False

try:
    import pygame
    pygame.mixer.init()
    PYGAME_AVAILABLE = True
except ImportError:
    PYGAME_AVAILABLE = False

try:
    import keyboard
    KEYBOARD_AVAILABLE = True
except ImportError:
    KEYBOARD_AVAILABLE = False

try:
    import sounddevice as sd
    import soundfile as sf
    RECORD_AVAILABLE = True
except ImportError:
    RECORD_AVAILABLE = False

try:
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    GCAL_AVAILABLE = True
except ImportError:
    GCAL_AVAILABLE = False

# ── Constants ─────────────────────────────────────────────────────────────────
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "attendance_config.json")
AUDIO_FILE  = os.path.join(SCRIPT_DIR, "present_sir.wav")
TOKEN_FILE  = os.path.join(SCRIPT_DIR, "google_token.json")
CREDS_FILE  = os.path.join(SCRIPT_DIR, "credentials.json")
SCOPES      = ["https://www.googleapis.com/auth/calendar.readonly"]

CHROME_PROFILE = r"C:\Users\mmoee\AppData\Local\Google\Chrome\User Data"

C = {
    "bg":       "#0d0d0d",
    "surface":  "#161616",
    "surface2": "#1e1e1e",
    "border":   "#2a2a2a",
    "green":    "#00e676",
    "green_dk": "#00c853",
    "amber":    "#ffab00",
    "red":      "#ff5252",
    "muted":    "#555555",
    "text":     "#e0e0e0",
    "subtext":  "#888888",
}

MEET_PATTERNS = [
    r"https://meet\.google\.com/[a-z\-]+",
    r"https://zoom\.us/j/\S+",
    r"https://teams\.microsoft\.com/l/meetup-join/\S+",
]

DEFAULT_CONFIG = {
    "name": "", "roll_number": "", "hotkey": "ctrl+shift+p",
    "aliases": [], "meetings": [], "google_connected": False
}

NUM_WORDS = {
    "zero":"0","one":"1","two":"2","three":"3","four":"4","five":"5",
    "six":"6","seven":"7","eight":"8","nine":"9","ten":"10",
    "eleven":"11","twelve":"12","thirteen":"13","fourteen":"14","fifteen":"15",
    "sixteen":"16","seventeen":"17","eighteen":"18","nineteen":"19",
    "twenty":"20","thirty":"30","forty":"40","fifty":"50",
    "forty nine":"49","forty-nine":"49","twenty one":"21","twenty two":"22",
    "twenty three":"23","twenty four":"24","twenty five":"25",
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE) as f:
                return {**DEFAULT_CONFIG, **json.load(f)}
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()

def save_config(cfg):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)

def normalize_text(text):
    t = text.lower().strip()
    for w, d in NUM_WORDS.items():
        t = t.replace(w, d)
    return t

# ── Styled widgets ────────────────────────────────────────────────────────────
def styled_entry(parent, width=28, **kw):
    return tk.Entry(parent, width=width,
                    font=("Consolas", 10),
                    bg=C["surface2"], fg=C["text"],
                    insertbackground=C["green"],
                    relief="flat",
                    highlightthickness=1,
                    highlightbackground=C["border"],
                    highlightcolor=C["green"], **kw)

def styled_btn(parent, text, color=None, **kw):
    color = color or C["surface2"]
    fg = C["bg"] if color == C["green"] else C["text"]
    return tk.Button(parent, text=text,
                     font=("Consolas", 10),
                     bg=color, fg=fg,
                     activebackground=C["border"],
                     activeforeground=C["text"],
                     relief="flat", cursor="hand2",
                     padx=12, pady=6, bd=0, **kw)

def lbl(parent, text, size=10, color=None, **kw):
    return tk.Label(parent, text=text,
                    font=("Consolas", size),
                    fg=color or C["subtext"],
                    bg=C["bg"], **kw)

# ── Google Calendar ───────────────────────────────────────────────────────────
def get_google_service(log_fn=None):
    if not GCAL_AVAILABLE:
        if log_fn: log_fn("[!] Google API not installed.")
        return None
    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        except Exception:
            pass
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                creds = None
        if not creds:
            if not os.path.exists(CREDS_FILE):
                if log_fn: log_fn("[!] credentials.json not found.")
                return None
            try:
                flow  = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
                creds = flow.run_local_server(port=0)
            except Exception as e:
                if log_fn: log_fn(f"[x] Auth failed: {e}")
                return None
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())
    try:
        return build("calendar", "v3", credentials=creds)
    except Exception as e:
        if log_fn: log_fn(f"[x] Build failed: {e}")
        return None

def get_google_email(service):
    """Get the email address of the connected Google account."""
    try:
        profile = service.calendarList().get(calendarId="primary").execute()
        return profile.get("id", "")
    except Exception:
        try:
            oauth2 = build("oauth2", "v2",
                           credentials=service._http.credentials)
            info = oauth2.userinfo().get().execute()
            return info.get("email", "")
        except Exception:
            return ""

def find_chrome_profile(email, log_fn=None):
    """Scan all Chrome profiles and find the one logged in with given email."""
    if not email:
        return "Profile 3"
    base = r"C:\Users\mmoee\AppData\Local\Google\Chrome\User Data"
    if not os.path.exists(base):
        return "Profile 3"
    try:
        for folder in os.listdir(base):
            prefs_path = os.path.join(base, folder, "Preferences")
            if not os.path.exists(prefs_path):
                continue
            try:
                with open(prefs_path, encoding="utf-8", errors="ignore") as f:
                    prefs = json.load(f)
                accounts = prefs.get("account_info", [])
                for acc in accounts:
                    if acc.get("email","").lower() == email.lower():
                        if log_fn:
                            log_fn(f"[+] Found Chrome profile '{folder}' for {email}")
                        return folder
            except Exception:
                continue
    except Exception as e:
        if log_fn: log_fn(f"[!] Profile scan error: {e}")
    if log_fn: log_fn(f"[!] No Chrome profile found for {email} — using Profile 3")
    return "Profile 3"

def get_matching_chrome_profile(cfg, log_fn=None):
    """
    Compare the Gmail logged into the app with all Chrome profiles.
    Returns the Chrome profile folder name that matches.
    Priority:
      1. Manually selected profile in settings (chrome_profile)
      2. Auto-detected from google_email match
      3. Fallback to Profile 3
    """
    # 1. Use manually selected profile if set
    manual = cfg.get("chrome_profile", "")
    if manual:
        if log_fn: log_fn(f"[*] Using saved Chrome profile: '{manual}'")
        return manual

    # 2. Match from connected Google email
    email = cfg.get("google_email", "")
    if email:
        profile = find_chrome_profile(email, log_fn)
        if profile:
            cfg["chrome_profile"] = profile
            save_config(cfg)
            return profile

    # 3. Fallback
    if log_fn: log_fn("[!] Could not match profile — using Profile 3")
    return "Profile 3"

def extract_meet_link(text):
    if not text: return None
    for pat in MEET_PATTERNS:
        m = re.search(pat, text)
        if m: return m.group(0)
    return None

def fetch_calendar_meetings(service, log_fn=None, days_ahead=7):
    now   = datetime.datetime.utcnow()
    end   = now + datetime.timedelta(days=days_ahead)
    try:
        result = service.events().list(
            calendarId="primary",
            timeMin=now.isoformat() + "Z",
            timeMax=end.isoformat() + "Z",
            singleEvents=True,
            orderBy="startTime"
        ).execute()
    except Exception as e:
        if log_fn: log_fn(f"[x] Calendar fetch error: {e}")
        return []

    meetings = []
    import time as _time
    utc_offset = datetime.timedelta(
        seconds=-_time.timezone if not _time.daylight else -_time.altzone)

    for ev in result.get("items", []):
        start_str = ev.get("start", {}).get("dateTime", "")
        end_str   = ev.get("end",   {}).get("dateTime", "")
        if not start_str or "T" not in start_str:
            continue
        try:
            dt_s = datetime.datetime.fromisoformat(start_str.replace("Z", "+00:00"))
            dt_e = datetime.datetime.fromisoformat(end_str.replace("Z", "+00:00")) if end_str else None
            local_s = datetime.datetime(*dt_s.utctimetuple()[:6]) + utc_offset
            local_e = datetime.datetime(*dt_e.utctimetuple()[:6]) + utc_offset if dt_e else None
        except Exception:
            continue

        url = ev.get("hangoutLink")
        if not url:
            for ep in ev.get("conferenceData", {}).get("entryPoints", []):
                if ep.get("entryPointType") == "video":
                    url = ep.get("uri"); break
        if not url: url = extract_meet_link(ev.get("description", ""))
        if not url: url = extract_meet_link(ev.get("location", ""))
        if not url: continue

        platform = "Google Meet" if "meet.google.com" in url else \
                   "Zoom" if "zoom.us" in url else \
                   "Teams" if "teams.microsoft" in url else "Other"

        meetings.append({
            "day":        local_s.strftime("%A")[:3],
            "start_time": local_s.strftime("%H:%M"),
            "end_time":   local_e.strftime("%H:%M") if local_e else "",
            "platform":   platform,
            "url":        url,
            "title":      ev.get("summary", "Untitled"),
            "source":     "google",
            "date":       local_s.strftime("%Y-%m-%d"),
        })

    if log_fn: log_fn(f"[+] Fetched {len(meetings)} meeting(s) from Google Calendar.")
    return meetings

# ── Audio ─────────────────────────────────────────────────────────────────────
def get_vbcable_device():
    try:
        devices = sd.query_devices()
        for i, d in enumerate(devices):
            name = d["name"].lower()
            if "cable input" in name and d["max_output_channels"] > 0:
                return i
        for i, d in enumerate(devices):
            name = d["name"].lower()
            if "vb-audio" in name and d["max_output_channels"] > 0:
                return i
    except Exception:
        pass
    return None

def play_audio(log_fn=None):
    if not os.path.exists(AUDIO_FILE):
        if log_fn: log_fn("[!] No audio recorded. Click 'Record Audio' first.")
        return
    if RECORD_AVAILABLE:
        try:
            data, fs = sf.read(AUDIO_FILE)
            vb = get_vbcable_device()
            if vb is not None:
                if log_fn: log_fn(f"[+] Playing via VB-Cable (device {vb})...")
                sd.play(data, fs, device=vb)
                sd.wait()
            else:
                if log_fn: log_fn("[!] VB-Cable not found — playing via speakers.")
                sd.play(data, fs)
                sd.wait()
            if log_fn: log_fn("[+] 'Present Sir' played.")
            return
        except Exception as e:
            if log_fn: log_fn(f"[x] Audio error: {e}")
    if PYGAME_AVAILABLE:
        try:
            pygame.mixer.music.load(AUDIO_FILE)
            pygame.mixer.music.play()
            if log_fn: log_fn("[+] Playing via speakers (fallback)...")
        except Exception as e:
            if log_fn: log_fn(f"[x] Pygame error: {e}")

def record_audio(seconds=3, log_fn=None):
    if not RECORD_AVAILABLE:
        if log_fn: log_fn("[!] sounddevice not installed.")
        return False
    try:
        if log_fn: log_fn(f"[*] Recording {seconds}s... Say 'Present Sir'!")
        rec = sd.rec(int(seconds * 44100), samplerate=44100, channels=1)
        sd.wait()
        sf.write(AUDIO_FILE, rec, 44100)
        if log_fn: log_fn("[+] Audio saved.")
        return True
    except Exception as e:
        if log_fn: log_fn(f"[x] Record error: {e}")
        return False

# ── Voice detector ────────────────────────────────────────────────────────────
class VoiceDetector:
    def __init__(self, name, roll, log_fn, play_fn, aliases=None, input_device=None):
        self.name          = name.lower().strip()
        self.roll          = roll.lower().strip()
        self.aliases       = [a.lower().strip() for a in (aliases or []) if a.strip()]
        self.log           = log_fn
        self.play          = play_fn
        self.running       = False
        self.train_mode    = False
        self.heard_samples = []
        self.input_device  = input_device or ""

    def start(self, train_mode=False):
        self.train_mode = train_mode
        self.running    = True
        if SR_AVAILABLE:
            threading.Thread(target=self._loop_google, daemon=True).start()
            self.log("[*] Voice detection: Google Speech")
        else:
            self.log("[!] SpeechRecognition not installed. Run: pip install SpeechRecognition pyaudio")

    def stop(self):
        self.running = False

    def _is_match(self, text):
        t = normalize_text(text)
        for trigger in [self.name, normalize_text(self.roll)] + self.aliases:
            if trigger and trigger in t:
                return True
        return False

    def _handle(self, text):
        if not text.strip(): return
        self.log(f"[~] Heard: \"{text}\"")
        if self.train_mode:
            self.heard_samples.append(text.lower().strip())
            return
        if self._is_match(text):
            self.log("[!] Match! Responding...")
            threading.Thread(target=self.play, daemon=True).start()

    def _get_device_index(self):
        if not self.input_device or self.input_device == "Default (system mic)":
            return None
        try:
            names = sr.Microphone.list_microphone_names()
            if self.input_device in names:
                return names.index(self.input_device)
        except Exception:
            pass
        return None

    def _loop_google(self):
        r   = sr.Recognizer()
        r.dynamic_energy_threshold = True
        dev_idx = self._get_device_index()
        try:
            mic = sr.Microphone(device_index=dev_idx)
        except Exception:
            mic = sr.Microphone()
        with mic as src:
            r.adjust_for_ambient_noise(src, duration=1)
        self.log(f"[*] Listening on: {self.input_device or 'default mic'}")
        while self.running:
            try:
                with mic as src:
                    audio = r.listen(src, timeout=5, phrase_time_limit=5)
                text = r.recognize_google(audio).lower()
                self._handle(text)
            except sr.WaitTimeoutError:
                pass
            except sr.UnknownValueError:
                pass
            except sr.RequestError as e:
                self.log(f"[x] Speech API: {e}")
                time.sleep(3)
            except Exception as e:
                self.log(f"[x] {e}")
                time.sleep(1)

# ── Meeting scheduler ─────────────────────────────────────────────────────────
class MeetingScheduler:
    def __init__(self, log_fn, cfg=None):
        self.log     = log_fn
        self.running = False
        self.joined  = set()
        self.cfg     = cfg or {}

    def start(self, meetings):
        self.meetings = meetings
        self.running  = True
        threading.Thread(target=self._loop, daemon=True).start()
        self.log(f"[*] Scheduler watching {len(meetings)} meeting(s).")

    def stop(self):
        self.running = False

    def _loop(self):
        while self.running:
            now = datetime.datetime.now()
            for i, m in enumerate(self.meetings):
                self._check(i, m, now)
            time.sleep(20)

    def _check(self, i, m, now):
        url   = m.get("url", "").strip()
        start = self._parse_time(m.get("start_time", ""))
        end   = self._parse_time(m.get("end_time", ""))
        day   = m.get("day", "").lower()
        date  = m.get("date", "")
        if not url or not start: return

        today     = now.strftime("%A")[:3].lower()
        today_iso = now.strftime("%Y-%m-%d")

        if date:
            if date != today_iso: return
        elif day and day not in ("daily", "every day", ""):
            if day != today: return

        sched = now.replace(hour=start.hour, minute=start.minute,
                             second=0, microsecond=0)
        diff  = (now - sched).total_seconds()
        key   = f"join_{i}_{today_iso}"

        if 0 <= diff < 90 and key not in self.joined:
            self.joined.add(key)
            title = m.get("title", f"Meeting {i+1}")
            self.log(f"[+] Joining: {title}")
            threading.Thread(target=self._auto_join,
                              args=(url,), daemon=True).start()

        if end:
            end_dt   = now.replace(hour=end.hour, minute=end.minute,
                                    second=0, microsecond=0)
            end_diff = (now - end_dt).total_seconds()
            end_key  = f"end_{i}_{today_iso}"
            if 0 <= end_diff < 90 and end_key not in self.joined:
                self.joined.add(end_key)
                self.log("[*] End time reached.")

    def _auto_join(self, url):
        try:
            from selenium import webdriver
            from selenium.webdriver.common.by import By
            from selenium.webdriver.chrome.options import Options
            from selenium.webdriver.chrome.service import Service
            from selenium.webdriver.support.ui import WebDriverWait
            from selenium.webdriver.support import expected_conditions as EC
            from webdriver_manager.chrome import ChromeDriverManager

            profile_dir = get_matching_chrome_profile(
                self.cfg if hasattr(self, "cfg") else {},
                self.log
            )

            chrome_exe = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
            if not os.path.exists(chrome_exe):
                chrome_exe = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

            self.log(f"[*] Launching Chrome profile '{profile_dir}'...")
            subprocess.Popen([
                chrome_exe,
                "--remote-debugging-port=9222",
                f"--profile-directory={profile_dir}",
                "--use-fake-ui-for-media-stream",
                url
            ])

            # Connect — retry every 2 sec up to 8 attempts
            opts = Options()
            opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
            driver = None
            for _ in range(8):
                time.sleep(2)
                try:
                    driver = webdriver.Chrome(
                        service=Service(ChromeDriverManager().install()),
                        options=opts
                    )
                    self.log("[*] Connected!")
                    break
                except Exception:
                    pass

            if not driver:
                self.log("[x] Chrome not reachable.")
                webbrowser.open(url)
                return

            # Click Got it immediately
            try:
                for btn in driver.find_elements(By.XPATH, "//button"):
                    if "got it" in btn.text.lower():
                        driver.execute_script("arguments[0].click();", btn)
                        self.log("[*] Clicked 'Got it'!")
                        time.sleep(1)
                        break
            except Exception:
                pass

            # Fill name if guest
            try:
                inp = driver.find_element(By.XPATH, "//input[@placeholder='Your name']")
                if inp.is_displayed():
                    inp.clear()
                    inp.send_keys("Moeed Zahid")
                    time.sleep(0.5)
            except Exception:
                pass

            # Click Join — max 10 sec wait
            wait = WebDriverWait(driver, 10)
            joined = False
            for text in ["Join now", "Ask to join"]:
                try:
                    btn = wait.until(EC.element_to_be_clickable((
                        By.XPATH, f"//button[contains(., '{text}')]"
                    )))
                    driver.execute_script("arguments[0].click();", btn)
                    self.log(f"[+] Joined!")
                    joined = True
                    break
                except Exception:
                    continue

            if not joined:
                self.log("[!] Join button not found — click manually.")

        except ImportError:
            self.log("[!] Run: pip install selenium webdriver-manager")
            webbrowser.open(url)
        except Exception as e:
            self.log(f"[x] Join error: {e}")
            webbrowser.open(url)

    @staticmethod
    def _parse_time(s):
        try:
            return datetime.datetime.strptime(s.strip(), "%H:%M").time()
        except Exception:
            return None

# ── Main App ──────────────────────────────────────────────────────────────────
class AttendanceBot(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Attendance Bot v2")
        self.geometry("860x680")
        self.configure(bg=C["bg"])
        self.resizable(True, True)

        self.cfg          = load_config()
        self.bot_running  = False
        self.voice_det    = None
        self.scheduler    = None
        self.meeting_rows = []
        self.cal_meetings = []
        self.gcal_service = None

        self._build_ui()
        self._load_manual_meetings()
        self._check_deps()
        if KEYBOARD_AVAILABLE:
            self._setup_hotkey()

    def _build_ui(self):
        hdr = tk.Frame(self, bg=C["bg"])
        hdr.pack(fill="x", padx=20, pady=(14, 6))
        tk.Label(hdr, text="ATTENDANCE BOT",
                  font=("Consolas", 17, "bold"),
                  fg=C["green"], bg=C["bg"]).pack(side="left")
        self.status_lbl = tk.Label(hdr, text="● OFFLINE",
                                    font=("Consolas", 10),
                                    fg=C["muted"], bg=C["bg"])
        self.status_lbl.pack(side="right")

        style = ttk.Style()
        style.theme_use("default")
        style.configure("TNotebook",     background=C["bg"], borderwidth=0)
        style.configure("TNotebook.Tab", background=C["surface"],
                         foreground=C["muted"],
                         font=("Consolas", 10), padding=[14, 6])
        style.map("TNotebook.Tab",
                  background=[("selected", C["bg"])],
                  foreground=[("selected", C["green"])])
        style.configure("TFrame", background=C["bg"])

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=16, pady=4)

        self.tab_meetings = ttk.Frame(self.nb)
        self.tab_calendar = ttk.Frame(self.nb)
        self.tab_settings = ttk.Frame(self.nb)
        self.tab_log      = ttk.Frame(self.nb)

        self.nb.add(self.tab_meetings, text="  Manual  ")
        self.nb.add(self.tab_calendar, text="  Google Calendar  ")
        self.nb.add(self.tab_settings, text="  Settings  ")
        self.nb.add(self.tab_log,      text="  Log  ")

        self._build_manual_tab()
        self._build_calendar_tab()
        self._build_settings_tab()
        self._build_log_tab()

        bar = tk.Frame(self, bg=C["bg"])
        bar.pack(fill="x", padx=16, pady=(0, 14))

        self.toggle_btn = tk.Button(
            bar, text="▶  START BOT",
            font=("Consolas", 12, "bold"),
            bg=C["green"], fg=C["bg"],
            activebackground=C["green_dk"],
            relief="flat", cursor="hand2",
            padx=20, pady=10, bd=0,
            command=self._toggle_bot
        )
        self.toggle_btn.pack(side="left", padx=(0, 10))
        styled_btn(bar, "▶  Test Audio",
                    command=lambda: threading.Thread(
                        target=play_audio, args=(self._log,), daemon=True
                    ).start()).pack(side="left", padx=4)
        styled_btn(bar, "⬤  Record Audio",
                    command=self._record_dialog).pack(side="left", padx=4)
        self.hk_lbl = lbl(bar, f"Hotkey: {self.cfg.get('hotkey','ctrl+shift+p')}")
        self.hk_lbl.pack(side="right")

    # ── Manual tab ────────────────────────────────────────────────────────────
    def _build_manual_tab(self):
        f = self.tab_meetings
        hdr = tk.Frame(f, bg=C["bg"])
        hdr.pack(fill="x", padx=10, pady=(10, 2))
        for txt, w in [("Day",8),("Start",6),("End",6),("Platform",11),("URL",0)]:
            tk.Label(hdr, text=txt, font=("Consolas", 9), fg=C["muted"],
                      bg=C["bg"], width=w, anchor="w").pack(side="left", padx=3)

        canvas = tk.Canvas(f, bg=C["bg"], highlightthickness=0)
        sb = tk.Scrollbar(f, orient="vertical", command=canvas.yview)
        self.rows_frame = tk.Frame(canvas, bg=C["bg"])
        self.rows_frame.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=self.rows_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True, padx=10)
        sb.pack(side="right", fill="y")

        btn_row = tk.Frame(f, bg=C["bg"])
        btn_row.pack(fill="x", padx=10, pady=8)
        styled_btn(btn_row, "+ Add Row",
                    command=self._add_manual_row).pack(side="left", padx=(0,6))
        styled_btn(btn_row, "Save",
                    command=self._save_manual).pack(side="left")

    def _add_manual_row(self, data=None):
        row = tk.Frame(self.rows_frame, bg=C["surface"],
                        highlightthickness=1, highlightbackground=C["border"])
        row.pack(fill="x", pady=2, padx=2)
        ekw = dict(font=("Consolas",9), bg=C["surface2"], fg=C["text"],
                    relief="flat", insertbackground=C["green"])

        day_var = tk.StringVar(value=(data or {}).get("day","Mon"))
        om = tk.OptionMenu(row, day_var, "Mon","Tue","Wed","Thu","Fri","Sat","Sun","Daily")
        om.config(font=("Consolas",9), bg=C["surface2"], fg=C["subtext"],
                   relief="flat", width=5, highlightthickness=0, bd=0)
        om["menu"].config(bg=C["surface2"], fg=C["subtext"])
        om.pack(side="left", padx=(4,2))

        s_e = tk.Entry(row, width=7, **ekw)
        s_e.insert(0, (data or {}).get("start_time","08:00"))
        s_e.pack(side="left", padx=2)

        tk.Label(row, text="→", fg=C["muted"], bg=C["surface"],
                  font=("Consolas",9)).pack(side="left")

        e_e = tk.Entry(row, width=7, **ekw)
        e_e.insert(0, (data or {}).get("end_time","09:00"))
        e_e.pack(side="left", padx=2)

        plat_var = tk.StringVar(value=(data or {}).get("platform","Zoom"))
        pm = tk.OptionMenu(row, plat_var, "Zoom","Google Meet","Teams","Other")
        pm.config(font=("Consolas",9), bg=C["surface2"], fg=C["subtext"],
                   relief="flat", width=10, highlightthickness=0, bd=0)
        pm["menu"].config(bg=C["surface2"], fg=C["subtext"])
        pm.pack(side="left", padx=(4,2))

        url_e = tk.Entry(row, width=36, **ekw)
        url_e.insert(0, (data or {}).get("url",""))
        url_e.pack(side="left", padx=2, fill="x", expand=True)

        if (data or {}).get("source") == "google":
            tk.Label(row, text="CAL", font=("Consolas",8),
                      fg=C["green"], bg=C["surface"]).pack(side="left")

        tk.Button(row, text="✕", font=("Consolas",9),
                   bg=C["surface"], fg=C["red"],
                   relief="flat", cursor="hand2", bd=0,
                   command=lambda r=row: self._del_row(r)).pack(side="right", padx=4)

        self.meeting_rows.append({
            "frame": row, "day": day_var, "start": s_e,
            "end": e_e, "platform": plat_var, "url": url_e,
            "source": (data or {}).get("source","manual"),
            "title":  (data or {}).get("title",""),
            "date":   (data or {}).get("date",""),
        })

    def _del_row(self, frame):
        self.meeting_rows = [r for r in self.meeting_rows if r["frame"] is not frame]
        frame.destroy()

    def _load_manual_meetings(self):
        for m in self.cfg.get("meetings", []):
            self._add_manual_row(m)

    def _save_manual(self):
        meetings = []
        for r in self.meeting_rows:
            meetings.append({
                "day":        r["day"].get(),
                "start_time": r["start"].get(),
                "end_time":   r["end"].get(),
                "platform":   r["platform"].get(),
                "url":        r["url"].get(),
                "source":     r.get("source","manual"),
                "title":      r.get("title",""),
                "date":       r.get("date",""),
            })
        self.cfg["meetings"] = meetings
        save_config(self.cfg)
        self._log("[+] Meetings saved.")

    # ── Calendar tab ──────────────────────────────────────────────────────────
    def _build_calendar_tab(self):
        f = self.tab_calendar
        card = tk.Frame(f, bg=C["surface"],
                         highlightthickness=1,
                         highlightbackground=C["border"])
        card.pack(fill="x", padx=14, pady=(14,8))
        tk.Label(card, text="Google Calendar",
                  font=("Consolas",11,"bold"),
                  fg=C["text"], bg=C["surface"]).pack(anchor="w", padx=14, pady=(10,2))
        self.gcal_status = tk.Label(card,
            text="Not connected",
            font=("Consolas",9), fg=C["amber"], bg=C["surface"])
        self.gcal_status.pack(anchor="w", padx=14, pady=(0,6))
        btn_row = tk.Frame(card, bg=C["surface"])
        btn_row.pack(anchor="w", padx=14, pady=(0,10))
        styled_btn(btn_row, "Connect Google",
                    command=self._connect_google).pack(side="left", padx=(0,6))
        styled_btn(btn_row, "Sync Now",
                    command=self._sync_calendar).pack(side="left", padx=6)
        styled_btn(btn_row, "Import to Manual",
                    command=self._import_to_manual).pack(side="left", padx=6)
        styled_btn(btn_row, "Logout",
                    color=C["red"],
                    command=self._logout_google).pack(side="left", padx=6)

        tk.Label(f, text="Upcoming meetings:",
                  font=("Consolas",9), fg=C["muted"], bg=C["bg"]).pack(
            anchor="w", padx=14, pady=(8,2))
        self.cal_listbox = tk.Text(f, font=("Consolas",9),
            bg=C["surface"], fg=C["text"],
            relief="flat", state="disabled", height=16, wrap="none")
        self.cal_listbox.pack(fill="both", expand=True, padx=14, pady=(0,14))

    def _logout_google(self):
        if not messagebox.askyesno("Logout", "Disconnect Google account?\nYou'll need to sign in again."):
            return
        # Delete saved token
        if os.path.exists(TOKEN_FILE):
            os.remove(TOKEN_FILE)
        self.gcal_service = None
        self.cal_meetings = []
        self.cfg["google_connected"] = False
        save_config(self.cfg)
        self.gcal_status.config(text="Logged out", fg=C["amber"])
        self._refresh_cal_list()
        self._log("[+] Logged out. Click 'Connect Google' to sign in with a different account.")

    def _connect_google(self):
        if not GCAL_AVAILABLE:
            messagebox.showerror("Missing Package",
                "Run: pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib")
            return
        if not os.path.exists(CREDS_FILE):
            messagebox.showinfo("Setup Required",
                "Place credentials.json in the app folder.")
            return
        self._log("[*] Connecting to Google...")
        def do():
            svc = get_google_service(self._log)
            if svc:
                self.gcal_service = svc
                self.cfg["google_connected"] = True
                # Get and save email
                email = get_google_email(svc)
                if email:
                    self.cfg["google_email"] = email
                    self._log(f"[+] Signed in as: {email}")
                    # Find matching Chrome profile
                    profile = find_chrome_profile(email, self._log)
                    self.cfg["chrome_profile"] = profile
                    self._log(f"[+] Chrome profile set to: '{profile}'")
                save_config(self.cfg)
                status_txt = f"Connected: {email}" if email else "Connected"
                self.after(0, lambda: self.gcal_status.config(
                    text=status_txt, fg=C["green"]))
                self._log("[+] Google Calendar connected!")
                self._sync_calendar()
            else:
                self.after(0, lambda: self.gcal_status.config(
                    text="Failed", fg=C["red"]))
        threading.Thread(target=do, daemon=True).start()

    def _sync_calendar(self):
        if not self.gcal_service:
            self._log("[!] Not connected.")
            return
        self._log("[*] Syncing...")
        def do():
            self.cal_meetings = fetch_calendar_meetings(
                self.gcal_service, self._log)
            self.after(0, self._refresh_cal_list)
        threading.Thread(target=do, daemon=True).start()

    def _refresh_cal_list(self):
        self.cal_listbox.config(state="normal")
        self.cal_listbox.delete("1.0", "end")
        if not self.cal_meetings:
            self.cal_listbox.insert("end", "  No meetings found.\n")
        else:
            for m in self.cal_meetings:
                line  = f"  {m.get('date','')}  {m.get('start_time','')}-{m.get('end_time','')}  [{m.get('platform','')}]  {m.get('title','')[:30]}\n"
                line2 = f"    → {m.get('url','')[:60]}\n\n"
                self.cal_listbox.insert("end", line)
                self.cal_listbox.insert("end", line2)
        self.cal_listbox.config(state="disabled")

    def _import_to_manual(self):
        if not self.cal_meetings:
            messagebox.showinfo("No meetings", "Sync calendar first.")
            return
        existing = {r["url"].get() for r in self.meeting_rows}
        count = 0
        for m in self.cal_meetings:
            if m.get("url") and m["url"] not in existing:
                self._add_manual_row(m)
                count += 1
        self._save_manual()
        self.nb.select(self.tab_meetings)
        self._log(f"[+] Imported {count} meeting(s).")

    # ── Settings tab ──────────────────────────────────────────────────────────
    def _build_settings_tab(self):
        f = self.tab_settings
        rows = [
            ("Your Name",   "name",        "as teacher calls you"),
            ("Roll Number", "roll_number", "e.g. 49"),
            ("Aliases",     "aliases",     "comma separated e.g. mood,moody"),
            ("Hotkey",      "hotkey",      "e.g. ctrl+shift+p"),
        ]
        for i, (lbl_txt, key, hint) in enumerate(rows):
            tk.Label(f, text=lbl_txt, font=("Consolas",10),
                      fg=C["subtext"], bg=C["bg"]).grid(
                row=i, column=0, sticky="w", padx=16, pady=10)
            e = styled_entry(f, width=32)
            val = self.cfg.get(key,"")
            if isinstance(val, list): val = ",".join(val)
            e.insert(0, val)
            e.grid(row=i, column=1, sticky="w", padx=8, pady=10)
            tk.Label(f, text=hint, font=("Consolas",8),
                      fg=C["muted"], bg=C["bg"]).grid(
                row=i, column=2, sticky="w", padx=4)
            setattr(self, f"_{key}_entry", e)

        # Chrome profile dropdown
        tk.Label(f, text="Chrome Profile:", font=("Consolas",10),
                  fg=C["subtext"], bg=C["bg"]).grid(
            row=4, column=0, sticky="w", padx=16, pady=10)
        chrome_profiles = self._get_chrome_profiles()
        saved_profile = self.cfg.get("chrome_profile", "")
        # Find display name for saved profile
        default_display = saved_profile
        for display, folder in chrome_profiles:
            if folder == saved_profile:
                default_display = display
                break
        self._chrome_profile_var = tk.StringVar(value=default_display or (chrome_profiles[0][0] if chrome_profiles else "Default"))
        profile_cb = ttk.Combobox(f,
            textvariable=self._chrome_profile_var,
            values=[p[0] for p in chrome_profiles],
            state="readonly", font=("Consolas",9), width=40)
        profile_cb.grid(row=4, column=1, sticky="w", padx=8, pady=10)
        tk.Label(f, text="profile with uni Gmail",
                  font=("Consolas",8), fg=C["muted"], bg=C["bg"]).grid(
            row=4, column=2, sticky="w", padx=4)
        # Store mapping for lookup
        self._chrome_profiles_map = {display: folder for display, folder in chrome_profiles}

        styled_btn(f, "Save Settings",
                    command=self._save_settings).grid(
            row=5, column=0, columnspan=2, sticky="w", padx=16, pady=12)

        # Auto-train
        tf = tk.Frame(f, bg=C["surface"],
                       highlightthickness=1, highlightbackground=C["border"])
        tf.grid(row=6, column=0, columnspan=3, sticky="ew", padx=16, pady=(0,10))
        tk.Label(tf, text="Auto-Train — say your name to teach the bot your accent",
                  font=("Consolas",9), fg=C["amber"], bg=C["surface"]).pack(
            anchor="w", padx=12, pady=(8,4))
        tr = tk.Frame(tf, bg=C["surface"])
        tr.pack(anchor="w", padx=12, pady=(0,8))
        self.train_lbl = tk.Label(tr, text="",
                                   font=("Consolas",9),
                                   fg=C["green"], bg=C["surface"])

        def start_train():
            self.train_lbl.config(text="Listening 10 sec... say your name!")
            self.train_lbl.pack(side="left", padx=8)
            if self.voice_det:
                self.voice_det.train_mode    = True
                self.voice_det.heard_samples = []
                self.after(10000, finish_train)
            else:
                name = self._name_entry.get().strip()
                roll = self._roll_number_entry.get().strip()
                self._train_det = VoiceDetector(name, roll, self._log, lambda: None)
                self._train_det.start(train_mode=True)
                self.after(10000, finish_train)

        def finish_train():
            det = self.voice_det or getattr(self, "_train_det", None)
            if not det: return
            samples = list(det.heard_samples)
            det.train_mode = False
            if hasattr(self, "_train_det"):
                self._train_det.stop()
            if not samples:
                self.train_lbl.config(text="Nothing heard. Try again.")
                return
            from collections import Counter
            words = []
            for s in samples: words.extend(s.split())
            common = [w for w,_ in Counter(words).most_common(5)]
            existing = self._aliases_entry.get().strip()
            new_val = ",".join(common)
            if existing: new_val = existing + "," + new_val
            self._aliases_entry.delete(0, "end")
            self._aliases_entry.insert(0, new_val)
            self._save_settings()
            self.train_lbl.config(text=f"Done! Aliases: {', '.join(common)}")
            self._log(f"[+] Training done. Heard: {samples}")

        styled_btn(tr, "▶  Train (10 sec)", color=C["amber"],
                    command=start_train).pack(side="left")
        self.train_lbl.pack(side="left", padx=8)

        tk.Label(f, text="Engine: Google Speech (works on all PCs, no setup needed)",
                  font=("Consolas",8), fg=C["green"], bg=C["bg"]).grid(
            row=7, column=0, columnspan=3, sticky="w", padx=16, pady=(0,4))

        # Input device selector (Stereo Mix)
        tk.Label(f, text="Listen on (mic):", font=("Consolas",10),
                  fg=C["subtext"], bg=C["bg"]).grid(
            row=8, column=0, sticky="w", padx=16, pady=6)
        in_devs = self._get_input_devices()
        saved_in = self.cfg.get("input_device","")
        default_in = saved_in if saved_in in in_devs else (in_devs[0] if in_devs else "")
        self._input_device_var = tk.StringVar(value=default_in)
        cb = ttk.Combobox(f, textvariable=self._input_device_var,
                           values=in_devs, state="readonly",
                           font=("Consolas",9), width=30)
        cb.grid(row=8, column=1, sticky="w", padx=8, pady=6)
        tk.Label(f, text="set to Stereo Mix to hear teacher",
                  font=("Consolas",8), fg=C["muted"], bg=C["bg"]).grid(
            row=8, column=2, sticky="w", padx=4)

        tk.Label(f,
                  text="Tip: Select the Chrome profile with your uni Gmail above.",
                  font=("Consolas",8), fg=C["amber"], bg=C["bg"],
                  wraplength=500).grid(
            row=9, column=0, columnspan=3, sticky="w", padx=16, pady=(0,4))

        self.deps_lbl = tk.Label(f, text="", font=("Consolas",8),
                                  fg=C["amber"], bg=C["bg"], justify="left",
                                  wraplength=500)
        self.deps_lbl.grid(row=10, column=0, columnspan=3, sticky="w", padx=16)
        styled_btn(f, "Install Missing Packages",
                    command=self._install_deps).grid(
            row=11, column=0, columnspan=2, sticky="w", padx=16, pady=6)

    def _get_input_devices(self):
        devs = ["Default (system mic)"]
        if SR_AVAILABLE:
            try:
                names = sr.Microphone.list_microphone_names()
                # Put Stereo Mix first
                names.sort(key=lambda x: 0 if any(k in x.lower()
                            for k in ["stereo mix","wave out","loopback"]) else 1)
                devs.extend(names)
            except Exception:
                pass
        return devs

    def _get_chrome_profiles(self):
        profiles = []
        base = r"C:\Users\mmoee\AppData\Local\Google\Chrome\User Data"
        if not os.path.exists(base):
            return [("Default", "Default")]
        try:
            for folder in sorted(os.listdir(base)):
                prefs_path = os.path.join(base, folder, "Preferences")
                if not os.path.exists(prefs_path):
                    continue
                try:
                    with open(prefs_path, encoding="utf-8", errors="ignore") as pf:
                        prefs = json.load(pf)
                    accounts = prefs.get("account_info", [])
                    email = accounts[0].get("email", "") if accounts else ""
                    display = f"{folder} — {email}" if email else folder
                    profiles.append((display, folder))
                except Exception:
                    profiles.append((folder, folder))
        except Exception:
            pass
        return profiles or [("Default", "Default")]

    def _save_settings(self):
        self.cfg["name"]         = self._name_entry.get().strip()
        self.cfg["roll_number"]  = self._roll_number_entry.get().strip()
        self.cfg["hotkey"]       = self._hotkey_entry.get().strip()
        self.cfg["input_device"] = self._input_device_var.get() if hasattr(self, "_input_device_var") else ""
        raw = self._aliases_entry.get().strip()
        self.cfg["aliases"] = [a.strip() for a in raw.split(",") if a.strip()]
        # Save selected Chrome profile folder name
        if hasattr(self, "_chrome_profile_var") and hasattr(self, "_chrome_profiles_map"):
            selected_display = self._chrome_profile_var.get()
            folder = self._chrome_profiles_map.get(selected_display, "Default")
            self.cfg["chrome_profile"] = folder
            self._log(f"[+] Chrome profile: '{folder}'")
        save_config(self.cfg)
        self._log(f"[+] Settings saved. Name='{self.cfg['name']}' Roll='{self.cfg['roll_number']}'")
        self.hk_lbl.config(text=f"Hotkey: {self.cfg['hotkey']}")
        if KEYBOARD_AVAILABLE:
            self._setup_hotkey()

    # ── Log tab ───────────────────────────────────────────────────────────────
    def _build_log_tab(self):
        f = self.tab_log
        self.log_box = scrolledtext.ScrolledText(
            f, font=("Consolas",9),
            bg="#080808", fg=C["green"],
            relief="flat", state="disabled",
            insertbackground=C["green"])
        self.log_box.pack(fill="both", expand=True, padx=8, pady=8)
        styled_btn(f, "Clear", command=self._clear_log).pack(
            anchor="e", padx=8, pady=(0,8))

    # ── Bot toggle ────────────────────────────────────────────────────────────
    def _toggle_bot(self):
        if self.bot_running: self._stop_bot()
        else: self._start_bot()

    def _start_bot(self):
        self._save_manual()
        name = self.cfg.get("name","")
        roll = self.cfg.get("roll_number","")
        if not name and not roll:
            messagebox.showwarning("Missing Info",
                "Enter your name or roll number in Settings.")
            return
        if not os.path.exists(AUDIO_FILE):
            if not messagebox.askyesno("No Audio",
                "No audio recorded.\nContinue anyway?"):
                return

        all_meetings = list(self.cfg.get("meetings",[]))
        existing_urls = {m.get("url") for m in all_meetings}
        for m in self.cal_meetings:
            if m.get("url") and m["url"] not in existing_urls:
                all_meetings.append(m)

        self.bot_running = True
        self.toggle_btn.config(text="■  STOP BOT", bg=C["red"], fg="#ffffff")
        self.status_lbl.config(text="● ACTIVE", fg=C["green"])

        self.voice_det = VoiceDetector(
            name, roll, self._log,
            lambda: play_audio(self._log),
            aliases=self.cfg.get("aliases",[]),
            input_device=self.cfg.get("input_device","")
        )
        self.voice_det.start()

        self.scheduler = MeetingScheduler(self._log, cfg=self.cfg)
        self.scheduler.start(all_meetings)

        self._log("="*52)
        self._log(f"[+] Bot ACTIVE | Name: '{name}' | Roll: '{roll}'")
        self._log(f"[+] Aliases: {self.cfg.get('aliases',[])}")
        self._log(f"[+] Listening on: {self.cfg.get('input_device','default mic')}")
        self._log('[+] Engine: Google Speech')
        self._log(f"[+] Meetings: {len(all_meetings)}")
        self._log(f"[+] Hotkey: {self.cfg.get('hotkey')}")
        self._log("="*52)

    def _stop_bot(self):
        self.bot_running = False
        self.toggle_btn.config(text="▶  START BOT", bg=C["green"], fg=C["bg"])
        self.status_lbl.config(text="● OFFLINE", fg=C["muted"])
        if self.voice_det: self.voice_det.stop()
        if self.scheduler: self.scheduler.stop()
        self._log("[*] Bot stopped.")

    def _setup_hotkey(self):
        try: keyboard.unhook_all()
        except Exception: pass
        hk = self.cfg.get("hotkey","ctrl+shift+p")
        try:
            keyboard.add_hotkey(hk, lambda: threading.Thread(
                target=play_audio, args=(self._log,), daemon=True).start())
            self._log(f"[+] Hotkey: {hk}")
        except Exception as e:
            self._log(f"[!] Hotkey error: {e}")

    def _record_dialog(self):
        win = tk.Toplevel(self)
        win.title("Record Audio")
        win.geometry("320x200")
        win.configure(bg=C["bg"])
        win.resizable(False, False)
        tk.Label(win, text="Press Record then say:\n\"Present Sir\"",
                  font=("Consolas",10), fg=C["text"],
                  bg=C["bg"], justify="center").pack(pady=16)
        dur = tk.IntVar(value=3)
        tk.Scale(win, from_=2, to=8, orient="horizontal",
                  variable=dur, label="Duration (sec)",
                  font=("Consolas",8),
                  bg=C["surface2"], fg=C["subtext"],
                  troughcolor=C["border"],
                  highlightthickness=0).pack(fill="x", padx=20)
        st = tk.Label(win, text="", font=("Consolas",9),
                       fg=C["amber"], bg=C["bg"])
        st.pack()
        def do():
            st.config(text="Recording...", fg=C["red"])
            win.update()
            ok = record_audio(dur.get(), self._log)
            st.config(text="Saved!" if ok else "Error",
                       fg=C["green"] if ok else C["red"])
        tk.Button(win, text="⬤  Record",
                   font=("Consolas",10,"bold"),
                   bg=C["red"], fg="#fff",
                   relief="flat", cursor="hand2",
                   padx=14, pady=6, bd=0,
                   command=lambda: threading.Thread(
                       target=do, daemon=True).start()
                   ).pack(pady=10)

    def _check_deps(self):
        missing = []
        if not SR_AVAILABLE:      missing.append("SpeechRecognition pyaudio")
        if not PYGAME_AVAILABLE:  missing.append("pygame")
        if not KEYBOARD_AVAILABLE: missing.append("keyboard")
        if not RECORD_AVAILABLE:  missing.append("sounddevice soundfile")
        if not GCAL_AVAILABLE:    missing.append("google-api-python-client google-auth-httplib2 google-auth-oauthlib")
        
        if missing:
            self.deps_lbl.config(
                text="Missing: " + "  |  ".join(missing))
        else:
            self.deps_lbl.config(text="All packages installed.", fg=C["green"])

    def _install_deps(self):
        pkgs = ("SpeechRecognition pyaudio pygame keyboard sounddevice soundfile "
                "google-api-python-client google-auth-httplib2 google-auth-oauthlib "
                "selenium webdriver-manager")
        self._log("[*] Installing packages...")
        def run():
            r = subprocess.run(
                [sys.executable, "-m", "pip", "install"] + pkgs.split(),
                capture_output=True, text=True)
            if r.returncode == 0:
                self._log("[+] Done! Restart the app.")
            else:
                self._log(f"[x] {r.stderr[-300:]}")
        threading.Thread(target=run, daemon=True).start()

    def _log(self, msg):
        def _do():
            self.log_box.config(state="normal")
            ts = datetime.datetime.now().strftime("%H:%M:%S")
            self.log_box.insert("end", f"[{ts}] {msg}\n")
            self.log_box.see("end")
            self.log_box.config(state="disabled")
        self.after(0, _do)

    def _clear_log(self):
        self.log_box.config(state="normal")
        self.log_box.delete("1.0","end")
        self.log_box.config(state="disabled")


if __name__ == "__main__":
    app = AttendanceBot()
    app.mainloop()
