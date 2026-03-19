@echo off
echo ============================================
echo   Attendance Bot v2 - Setup
echo ============================================
echo.
python --version
if errorlevel 1 (
    echo Python not found! Get it from python.org
    pause
    exit
)
echo.
echo Installing all dependencies...
pip install SpeechRecognition pyaudio pygame keyboard sounddevice soundfile google-api-python-client google-auth-httplib2 google-auth-oauthlib
echo.
echo ============================================
echo   Done! Run: python attendance_bot.py
echo ============================================
pause
