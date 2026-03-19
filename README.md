# 🎓 Attendance Bot

Automatically joins your university meetings and says "Present Sir" when your name or roll number is called.

**Works with:** Google Meet • Zoom • Microsoft Teams

---

## ⚡ Quick Start

```
1. Install Python 3.10+
2. Run setup.bat
3. Follow the setup guide below
4. Run: python attendance_bot.py
```

---

## 📋 Requirements

- Windows 10/11
- Python 3.10 or higher → https://python.org/downloads
- Google Chrome installed
- Internet connection

---

## 🔧 Installation

### Step 1 — Install Python packages

Double-click `setup.bat` OR run in CMD:

```
pip install sounddevice soundfile SpeechRecognition pyaudio keyboard selenium webdriver-manager google-api-python-client google-auth-httplib2 google-auth-oauthlib pygame
```

### Step 2 — Get credentials.json (for Google Calendar)

You need your own Google API credentials. Follow these steps:

1. Go to https://console.cloud.google.com/
2. Sign in with any Gmail (personal Gmail works fine)
3. Click **Select a project** → **New Project**
4. Name it anything (e.g. `attendance-bot`) → **Create**
5. Left sidebar → **APIs & Services** → **Library**
6. Search `Google Calendar API` → click it → **Enable**
7. Left sidebar → **APIs & Services** → **Credentials**
8. Click **+ Create Credentials** → **OAuth client ID**
9. If asked to configure consent screen:
   - Click **Configure Consent Screen**
   - Choose **External** → **Create**
   - Fill App name (anything), your Gmail → **Save and Continue** (skip rest)
   - Go to **Test users** → **+ Add Users** → add your university Gmail → **Save**
   - Go back to **Credentials**
10. Click **+ Create Credentials** → **OAuth client ID**
11. Application type: **Desktop app** → Name: anything → **Create**
12. Click **Download JSON**
13. Rename the downloaded file to exactly: `credentials.json`
14. **Place it in the same folder as `attendance_bot.py`**

---

## ⚙️ First Time Setup (do in this order)

### Step 1 — Settings Tab

Open the app and go to **Settings** tab:

| Field | What to enter |
|---|---|
| Your Name | Exactly how teacher calls you (e.g. `Your Name`) |
| Roll Number | Your roll number (e.g. `Rollno`) |
| Aliases | Leave empty for now (auto-filled after training) |
| Hotkey | Keep as `ctrl+shift+p` or change to your preference |
| Chrome Profile | Select the profile with your university Gmail |

Click **Save Settings**.

### Step 2 — Record Your Audio

1. Click **⬤ Record Audio** button at the bottom
2. When it says "Recording..." — say **"Present Sir"** clearly
3. Wait for "Saved!" confirmation
4. Click **▶ Test Audio** to make sure it plays correctly

### Step 3 — Train for Your Accent (IMPORTANT)

Google Speech Recognition may not recognize your name correctly due to accent differences. Train it:

1. Go to **Settings** tab
2. Click **▶ Train (10 sec)**
3. During those 10 seconds — say your **name** and **roll number** multiple times
   - Example: say "Jack... Jack Jack... 01... Jack... one"
4. After 10 seconds it shows: `Done! Aliases: Jack, Ack, 01...`
5. These aliases are now saved — bot will respond to all of them
6. Click **Save Settings**

> **Tip:** Do training 2-3 times and say your name in different ways — how teacher might say it, fast, slow, with surname.

### Step 4 — VB-Cable Setup (REQUIRED for audio in meeting)

Without VB-Cable, only YOU hear "Present Sir" — teacher won't hear it.

1. Download VB-Cable: https://download.vb-audio.com/Download_CABLE/VBCABLE_Driver_Pack43.zip
2. Extract the zip
3. Right-click `VBCABLE_Setup_x64.exe` → **Run as Administrator**
4. Click **Install Driver** → **OK**
5. **Restart your PC**
6. After restart, open Google Meet/Zoom → Settings → Microphone
7. Select **CABLE Output (VB-Audio Virtual Cable)**

> ⚠️ You must set VB-Cable as your microphone in the meeting app EVERY TIME before class.

### Step 5 — Connect Google Calendar

1. Place `credentials.json` in the app folder
2. Go to **Google Calendar** tab
3. Click **Connect Google**
4. Browser opens → sign in with your **university Gmail**
5. It says "Google hasn't verified this app" → click **Advanced** → **Go to attendance-bot (unsafe)**
6. Click **Allow**
7. Browser says "Authentication successful" — go back to app
8. Status shows **Connected** in green
9. Click **Sync Now** to fetch your meetings
10. Click **Import to Meetings** to add them to the schedule

### Step 6 — Add Meetings

**From Google Calendar (automatic):**
- After connecting Google → Sync Now → Import to Meetings

**Manually:**
1. Go to **Meetings** tab
2. Click **+ Add**
3. Fill in: Day, Start time (HH:MM), End time, Platform, Meeting URL
4. Click **Save**

> Meeting URL formats:
> - Google Meet: `https://meet.google.com/xxx-xxxx-xxx`
> - Zoom: `https://zoom.us/j/XXXXXXXXX`
> - Teams: `https://teams.microsoft.com/l/meetup-join/...`

---

## 🚀 Running the Bot

1. **Close Chrome completely** before starting the bot
2. Run `python attendance_bot.py`
3. Make sure settings are saved
4. Click **▶ START BOT**
5. Bot will:
   - Auto-open Chrome with your university profile at meeting time
   - Auto-click "Join now" or "Ask to join"
   - Listen for your name/roll number
   - Play "Present Sir" through VB-Cable into the meeting

---

## ⌨️ Hotkey

Press `Ctrl+Shift+P` (or your custom hotkey) at any time to instantly play "Present Sir" — useful as a backup if voice detection misses your name.

---

## 📁 Files in this folder

| File | Purpose |
|---|---|
| `attendance_bot.py` | Main application |
| `credentials.json` | Your Google API credentials (you create this) |
| `attendance_config.json` | Your saved settings (auto-created) |
| `google_token.json` | Google login token (auto-created) |
| `present_sir.wav` | Your recorded audio (auto-created) |
| `setup.bat` | Install packages |

---

## ❌ Troubleshooting

| Problem | Fix |
|---|---|
| `credentials.json not found` | Place it in the app folder (same as attendance_bot.py) |
| Bot doesn't hear my name | Do the accent training (Step 3) |
| Teacher can't hear "Present Sir" | Set VB-Cable as mic in meeting app |
| Chrome opens wrong profile | Settings tab → Chrome Profile → select correct one |
| Auto-join not working | Close Chrome before starting bot |
| `pyaudio` install fails | Run: `pip install pipwin` then `pipwin install pyaudio` |
| Google Calendar shows 0 meetings | Make sure meetings have a Google Meet/Zoom link |
| `python` not recognized | Add Python to PATH during installation |

---

## 🔒 Privacy & Security

- `credentials.json` — **Never share publicly or upload to GitHub**
- `google_token.json` — **Never share** (contains your login session)
- `attendance_config.json` — Contains your name/roll (don't share publicly)
- `present_sir.wav` — Your voice recording (keep private)

Add this `.gitignore` file to your repo to prevent accidental uploads:

```
credentials.json
google_token.json
attendance_config.json
present_sir.wav
*.wav
__pycache__/
*.pyc
```

---

## 👥 For Classmates

Each person needs their own `credentials.json`. Steps:
1. Follow **Step 2** above to create your own
2. Sign in with **your own** university Gmail when connecting
3. Everything else works the same

---

## 📞 Support

If something isn't working, check the **Log tab** in the app — it shows exactly what the bot is doing and any errors.
