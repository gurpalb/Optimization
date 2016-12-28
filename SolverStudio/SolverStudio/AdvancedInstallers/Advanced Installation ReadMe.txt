Current User Install
====================
If you run the SolverStudio "setup.exe" or "SolverStudio.vsto" file, SolverStudio will be installed for the current user. No Administrator permissions are needed. Running setup.exe installs all the required components (if they are not already installed); running SolverStudio.vsto skips this (potentially important!) step.

Uninstall Current User Install
==============================
An installation for the current user only can be uninstalled using the Control Panel, or using SolverStudio's About... Uninstall button.

All User Installs
=================
The "InstallAllUsers.xlsx" and the QuickInstall files allow SolverStudio to be easily installed for "All Users".

QuickInstall AllUsers 32bitWindows.bat
======================================
If you have a 32 bit version of Windows, you can use this QuickInstall batch file to install SolverStudio for all users if:

(1) Your SolverStudio files have been copied to:
      C:/Program Files/SolverStudio
    so that, for example, the setup.exe file is located at
      C:/Program Files/SolverStudio/SolverStudio/setup.exe

(2) You already have all the .Net v4 files and Office VSTO files installed (which is true for most modern systems, or if you have previously run SolverStudio's setup.exe program.)

(3) You have uninstalled any existing (current-user) SolverStudio installation

(4) You have patched Office 2007 to allow all-user installs; see http://SolverStudio.org

Note that the QuickInstall batch files copy files from:
  C:\Program Files\SolverStudio\SolverStudio\Application Files\SolverStudio_00_05_20_00
to
  C:\Program Files\SolverStudio\SolverStudio 
removing any ".deploy" extension from DLL's.  

QuickInstall AllUsers 64bitWindows 32bitOffice.bat
==================================================
If you have a 64 bit version of Windows with a 32 bit version of Office, you can use this QuickInstall batch file to install SolverStudio for all users if:

(1) Your SolverStudio files have been copied to:
      C:/Program Files (x86)/SolverStudio
    so that, for example, the setup.exe file is located at
      C:/Program Files (x86)/SolverStudio/SolverStudio/setup.exe

(2) You already have all the .Net v4 files and Office VSTO files installed (which is true for most modern systems, or if you have previously run SolverStudio's setup.exe program.)

(3) You have uninstalled any existing (current-user) SolverStudio installation

(4) You have patched Office 2007 to allow all-user installs; see http://SolverStudio.org

Note that the QuickInstall batch files copy files from:
  C:\Program Files (x86)\SolverStudio\SolverStudio\Application Files\SolverStudio_00_05_20_00
to
  C:\Program Files (x86)\SolverStudio\SolverStudio 
removing any ".deploy" extension from DLL's.  

QuickInstall AllUsers 64bitWindows 64bitOffice.txt
==================================================
If you are installing SolverStudio for all user on a 64 bit Windows with 64 bit Office, then you can use the following (32 bit) file:
QuickInstall AllUsers 32bitWindows.bat

InstallAllUsers.xlsx
====================
This should be used to install SolverStudio for all users if the quick install files don't apply (or don't work) for you, or you don't know if you have 32 or 64 versions of Windows or Office.

Be sure to first copy the SolverStudio folder into either:
  C:\Program Files (x86)\SolverStudio\
if you are running Office 32 bit on WIndows 64 bit, and
  C:\Program Files\SolverStudio\
otherwise. (See above for more detailed comments on this location.)

UninstallAllUsers.bat
=====================
This will uninstall an all-users installation on both 32 and 64 bit Windows systems.

SolverStudio About... Uninstall
===============================
If SolverStudio has installed successful, its About... box has an un-install button that will uninstall both current-user and all-user installations. It will also brute-force remove registry settings if the VSTO installer cannot be found to remove a current-user installation. If this brute-force registry removal occurs, then you should run the "RefreshClickOnceCache.bat" file if needed to fix any resulting issues.

Extraneous Files
================
The other files in this folder are provided to support Office 2007 and debugging of SolverStudio (i.e. VSTO) installation problems. Running the "RefreshClickOnceCache.bat" file can sometimes help fix minor glitches.



