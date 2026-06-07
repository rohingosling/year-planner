@echo off
REM Build the standalone Year Planner executable (dist\year-planner.exe)
REM Stamps version + build date, runs PyInstaller via year-planner.spec, then smoke-tests the exe.
REM Re-bundles config\config.yaml, so edit the YAML first, then run this to bake the changes into the exe.
REM Usage: build.bat

cls
.venv\Scripts\python.exe scripts\build.py
