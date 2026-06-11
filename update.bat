@echo off
REM ============================================================
REM  Update the live prospect map with a new Excel file.
REM
REM  HOW TO USE:
REM   1. Drop your new .xlsx into THIS folder
REM      (you can leave the old one; the newest file is used).
REM   2. Double-click this update.bat.
REM   3. Wait for it to say "Done", then give Vercel ~1 minute.
REM ============================================================
cd /d "%~dp0"

echo.
echo === Step 1/2: Reading the newest Excel and geocoding addresses ===
echo (New addresses take ~1 sec each; ones seen before are instant.)
echo.
python build_data.py
if errorlevel 1 (
  echo.
  echo *** BUILD FAILED - see the message above. Nothing was published. ***
  pause
  exit /b 1
)

echo.
echo === Step 2/2: Publishing to the live site ===
git add -A
git commit -m "Update prospect data (%DATE% %TIME%)"
git push

echo.
echo ============================================================
echo  Done. Vercel will redeploy your link in ~30-60 seconds.
echo  Refresh the site to see the new data.
echo ============================================================
pause
