@echo.
@setlocal
@set _appdir=c:\apps\WindowsLayoutSnapshot\
@set _startup="%appdata%\Microsoft\Windows\Start Menu\Programs\Startup"

xcopy /s /v /y /w /i x64\Debug\*.*  %_appdir%
  
@echo.
@echo.  Creating %_startup%\WindowsSnap.url
@echo [InternetShortcut] > %_startup%\WindowsSnap.url
@echo URL=%_appdir%\WindowsSnap.exe >> %_startup%\WindowsSnap.url

@echo.
@echo. pausing 6s...
@ping -n 6 localhost >nul

@start %_appdir%
