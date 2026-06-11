@echo off
setlocal

cd /d "%~dp0"

set "STREAMLIT_CMD=C:\Users\jumsu\AppData\Local\Programs\Python\Python313\Scripts\streamlit.exe"

if not exist "%STREAMLIT_CMD%" (
  echo streamlit.exe not found:
  echo %STREAMLIT_CMD%
  pause
  exit /b 1
)

echo Starting server...
"%STREAMLIT_CMD%" run uidemo_streamlit.py --server.port 8501

pause
