@echo off
rem Starts Planner. Keep this file next to outlook_todoist_planner.py and the planner folder.
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
  py outlook_todoist_planner.py %*
) else (
  python outlook_todoist_planner.py %*
)
if errorlevel 1 pause
