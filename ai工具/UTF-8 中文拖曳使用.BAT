@echo off
chcp 65001 >nul
"%~dp0exiftool(-k -lang zh_tw).exe" %*
pause
