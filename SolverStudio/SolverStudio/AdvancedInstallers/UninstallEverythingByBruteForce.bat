rem Remove Registry Settings
rem OS32_Office32 settings cover the OS64_Office64 case as well
regedit -s .\SupportFiles\Delete_HKLM_OS32_Office32.reg
regedit -s .\SupportFiles\Delete_HKLM_OS64_Office32.reg
regedit -s .\SupportFiles\Delete_HKCU_OS32_Office32.reg
regedit -s .\SupportFiles\Delete_HKCU_OS64_Office32.reg

rem Remove Registry enttries created by SolverStudio for user settings
regedit -s .\SupportFiles\Delete_SolverStudioSettings.reg

rem Update the ClickOnce cache: http://ddkonline.blogspot.co.nz/2011/07/fix-for-vsto-clickonce-application.html
rem This is suitable for client PCs (& development machines)
rundll32 dfshim CleanOnlineAppCache

