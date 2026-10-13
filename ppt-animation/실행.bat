@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0make_ppt_animation.ps1" %*
pause
