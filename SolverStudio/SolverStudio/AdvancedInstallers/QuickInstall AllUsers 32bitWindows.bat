@ECHO OFF

ECHO SolverStudio Quick Installer
ECHO ============================
ECHO %0
ECHO.

set _batchFileName=C:\Program Files\SolverStudio\SolverStudio\AdvancedInstallers\QuickInstall AllUsers 32bitWindows.bat
if "%~f0" == "%_batchFileName%" goto checkVSTO

ECHO.
ECHO This batch file has name:
ECHO "%~f0"
ECHO but this installer was expecting the name:
ECHO "%_batchFileName%"
ECHO.
ECHO Please copy the SolverStudio folder into C:\Program Files, and try again.
ECHO.
pause
GOTO:EOF

:checkVSTO

set _vstoFile=C:\Program Files\SolverStudio\SolverStudio\SolverStudio.vsto

IF EXIST "%_vstoFile%" GOTO doInstall

ECHO.
ECHO Unable to find the SolverStudio.vsto file:
ECHO "%_vstoFile%"
ECHO.
ECHO Please copy the SolverStudio folder into C:\Program Files, and try again.
ECHO.
pause
GOTO:EOF

:doInstall
REM Iterate thru all sub-directories in "Application Files" (there should only be one)
REM looking for a dir such as SolverStudio_00_06_12_00, being the folder we will copy from
set AppFilesDir="%~dp0..\Application Files\SolverStudio*"
FOR /D %%A in (""%AppFilesDir%"") DO SET _src=%%~fA\
set _dest=%~dp0..\

@ECHO ON
copy "%_src%SolverStudio.dll.manifest" "%_dest%"
copy "%_src%SolverStudio.dll.deploy" "%_dest%SolverStudio.dll"
copy "%_src%IronPython.Modules.dll.deploy" "%_dest%IronPython.Modules.dll"
copy "%_src%IronPython.dll.deploy" "%_dest%IronPython.dll"
copy "%_src%Microsoft.Dynamic.dll.deploy" "%_dest%Microsoft.Dynamic.dll"
copy "%_src%Microsoft.Scripting.dll.deploy" "%_dest%Microsoft.Scripting.dll"
copy "%_src%Microsoft.Scripting.Metadata.dll.deploy" "%_dest%Microsoft.Scripting.Metadata.dll"
copy "%_src%Microsoft.Office.Tools.Common.v4.0.Utilities.dll.deploy" "%_dest%Microsoft.Office.Tools.Common.v4.0.Utilities.dll"
copy "%_src%ScintillaNET.dll.deploy" "%_dest%ScintillaNET.dll"

rem Add Registry Settings
rem The 32,32 bit registry settings will also work 64,64 bit
regedit "%~dp0SupportFiles\Install_HKLM_OS32_Office32.reg"

rem Update the ClickOnce cache: http://ddkonline.blogspot.co.nz/2011/07/fix-for-vsto-clickonce-application.html
rem This is suitable for client PCs (& development machines)
rundll32 dfshim CleanOnlineAppCache

pause
GOTO:EOF 

