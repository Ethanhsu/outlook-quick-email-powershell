@echo off
cd /d "%~dp0"
powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File "quick-email.ps1"