@echo off
REM Quick manual deployment to Netlify
REM Run this to immediately update the Netlify site

cd /d "%~dp0"
echo Generating static snapshot...
python scripts\generate_static_pvs.py
if errorlevel 1 (
    echo ERROR: Failed to generate snapshot
    pause
    exit /b 1
)

echo.
echo Deploying to Netlify...
rem Deploy from project root using pre-built static folder
netlify deploy --prod --dir=netlify_static --no-build
if errorlevel 1 (
    echo ERROR: Netlify deployment failed
    pause
    exit /b 1
)

echo.
echo SUCCESS: Dashboard deployed to Netlify!
pause
