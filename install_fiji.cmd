@echo off

setlocal

rem Check if the AppData folder exists
if not exist "%APPDATA%" (
    echo The AppData folder could not be found.
    exit /b 1
)

rem Create the Fiji folder
set FIJI_HOME=%APPDATA%\Fiji.app
if not exist "%FIJI_HOME%" (
    mkdir "%FIJI_HOME%"
)

rem Download the Fiji installer
set FIJI_INSTALLER_URL=https://downloads.imagej.net/fiji/latest/fiji-win64.zip
set FIJI_INSTALLER_FILE=%TEMP%\fiji.zip
powershell -Command "(New-Object System.Net.WebClient).DownloadFile('%FIJI_INSTALLER_URL%', '%FIJI_INSTALLER_FILE%')"

rem Extract the Fiji installer to the AppData folder
powershell -Command "Expand-Archive '%FIJI_INSTALLER_FILE%' -DestinationPath '%FIJI_HOME%'"

rem Remove the Fiji installer
del "%FIJI_INSTALLER_FILE%"

rem Add ImageJ to the Start menu
set START_MENU_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Fiji
if not exist "%START_MENU_FOLDER%" (
    mkdir "%START_MENU_FOLDER%"
)
set SHORTCUT_FILE=%START_MENU_FOLDER%\ImageJ.lnk
set TARGET_FILE=%FIJI_HOME%\Fiji.app\ImageJ-win64.exe
set ICON_FILE=%FIJI_HOME%\Fiji.app\ImageJ-win64.exe

echo Set objShell = WScript.CreateObject("WScript.Shell") > "%TEMP%\CreateShortcut.vbs"
echo Set objShortcut = objShell.CreateShortcut("%SHORTCUT_FILE%") >> "%TEMP%\CreateShortcut.vbs"
echo objShortcut.TargetPath = "%TARGET_FILE%" >> "%TEMP%\CreateShortcut.vbs"
echo objShortcut.IconLocation = "%ICON_FILE%" >> "%TEMP%\CreateShortcut.vbs"
echo objShortcut.Save >> "%TEMP%\CreateShortcut.vbs"
cscript /nologo "%TEMP%\CreateShortcut.vbs"

rem Run Fiji
"%FIJI_HOME%\Fiji.app\ImageJ-win64.exe"