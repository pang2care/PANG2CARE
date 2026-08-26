@echo off
cd /d "%~dp0"

echo ===================================
echo   PANG2CARE - Git Save
echo ===================================
echo.
echo Changed files:
git status -s
echo.

set "msg=%*"
if not "%msg%"=="" goto havemsg
set /p msg=Enter commit message (press Enter for default): 
if "%msg%"=="" set "msg=update"

:havemsg
git add -A
git commit -m "%msg%"
git push

echo.
echo ===================================
echo   Done! Site updates in 1-2 minutes.
echo   https://pang2care.github.io/PANG2CARE/
echo ===================================
pause
