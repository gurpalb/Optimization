
This SupportFiles folder is added to the system PATH environment variable for the SolverStudio process when it starts up, and so any files in here are accessible to SolverStudio via the PATH.

This folder is always added to the PATH. 

This SupportFiles folder includes the files

SciLexer.dll
SciLexer64.dll
 - the DLL's needed for SolverStudio's Scintilla.Net code editor. Scintilla chooses which DLL to open at run time.
 
gdxdclib.dll, gdxdclib64.dll, gmszlib1.dll, gmszlib164.dll
 - the DLL's used by GAMS (which chooses the right versions to match the bitness)
 
The GAMS files are distributed with the kind support of GAMS; http://www.gams.com, and can be found by downloading and installing a 64 bit version of GAMS. The files are found in, for example, C:\GAMS\win64\24.4\

Note:
In previous versions (before 20150811), one of SupportFiles/SupportFiles32 or SupportFiles/SuppoertFiles64 was added, depending on the bit-ness of the SolverStudio process. (Note that this must match the process bitness, not the operating system bitness.) This was done for GAMS, but as of GAMS 24.4, is no longer needed (and so is no longer done in the code).


