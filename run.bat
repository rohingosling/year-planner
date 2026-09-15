@echo off
REM Run the Year Planner generator
REM Usage: run.bat [options]
REM   --year YEAR       Planner year (default: current year)
REM   --filename NAME   Output filename or path (default: year-planner-<year>)
REM   --pdf             Also produce a PDF (Windows + Microsoft Word only)
REM   --version         Show version and exit
REM   --help            Show help and exit

cls
.venv\Scripts\python.exe -X utf8 src/main.py %*
