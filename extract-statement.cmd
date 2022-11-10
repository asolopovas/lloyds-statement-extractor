@echo off
SET script_path=%~dp0
python.exe %script_path%extract-statement.py %*
