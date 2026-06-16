@echo off
setlocal

cd /d "%~dp0" || (
  echo [ERROR] Failed to change directory to "%~dp0"
  pause
  exit /b 1
)

where python >nul 2>nul || (
  echo [ERROR] python was not found in PATH.
  pause
  exit /b 1
)

call :find_script_dir "uk200_exe_accdb.py" || goto :fail

call :run "uk200_exe_accdb.py" || goto :fail
call :run "uk201.py" || goto :fail
call :run "uk203.py" || goto :fail
call :run "uk204.py" || goto :fail
call :run "uk205.py" || goto :fail
call :run "uk206.py" || goto :fail
call :run "uk2061.py" || goto :fail
call :run "uk207.py" || goto :fail
call :run "uk208.py" || goto :fail
call :run "uk209.py" || goto :fail

echo.
echo [DONE] UK200 pipeline finished.
pause
exit /b 0

:find_script_dir
set "SCRIPT_DIR="
for /d %%D in ("%~dp0*") do if exist "%%~fD\%~1" if not defined SCRIPT_DIR set "SCRIPT_DIR=%%~fD\"
if not defined SCRIPT_DIR (
  echo [ERROR] Script not found: %~1
  exit /b 1
)
exit /b 0

:run
set "CURRENT_SCRIPT=%~1"
echo.
echo [RUN] %CURRENT_SCRIPT%
python "%SCRIPT_DIR%%CURRENT_SCRIPT%"
if errorlevel 1 (
  echo [ERROR] Failed: %CURRENT_SCRIPT%
  exit /b 1
)
exit /b 0

:fail
echo.
echo [STOPPED] UK200 pipeline aborted.
pause
exit /b 1
