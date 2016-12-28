SolverStudio Directories

The SolverStudio directories contain the following.

AdvancedInstallers
Installation tools for doing more than a standard VSTO installation. You can use these to install for "all users" on a machine.

IronPython:
License files etc for IronPython; not used by SolverStudio
Lib directory used by IronPython within SolverStudio. Some of the IronPython Lib files differ from the CPython versions, and so this folder is usually required (and is included in the download). However, if it is missing, SolverStudio will use the CPython/Lib folder instead.


PuLP:
The PuLP files as used both when running PuLP natively and under CPython; this is a full copy of PuLP including Solvers etc, with the following changes:
1/ The PuLP/src/pulp/solverdir has all copies of cbc removed, and instead PuLP finds CBC from the PATH which contains the Solvers and the appropriate Solvers/XXbit (to match the operating system) path.
2/ A VERSION file is added in the PuLP directory which is read by SolverStudio.
3/ The following code change is made in solver.py, under class PYGLPK(LpSolver):
    try:
        #import the model into the global scope
        global glpk
        import glpk.glpkpi as glpk
    except: # AJM 2013; cannot use "except ImportError" in IronPython

CPython:
init.py - language setup file
RunCPython.py - file used to run CPython
Notes: 
1/ SolverStudio prefers to use an installed version of CPython, which it looks for by searching for python.exe in the entries in the environment variable 'path' and, if this fails, in a folder "python*" within the program files folder(s).
2/ When the user starts their Python file with an import SolverStudio command, this adds SolverStudio's PuLP directory to the Python path
3/ SolverStudio will look in the CPython folder for a Python installation if no other version can be found. If you wish to use this, then place in this folder a full Python 2.6.2 setup (or later) as installed by the Portable Python installer. 
4/ Some older versions of PuLP require lib/pyparsing.py - a file that should be added to the Lib directory as it is used by PuLP (if it is available) to read AMPL files. Newer PuLP versions will work even if this file is deleted, but AMPL data reading won't be possible.

GLPK
init.py - language setup file for GLPK that adds GLPK-specific menu items
RunGLPSol.py - file used to solve GMPL models
gmpl.pdf - GMPL documentation accessed from within SolverStudio using the GMPL menu added by init.py
glpsol.exe - the GNU interpretor for GMPL models and associated solver
Other GLPK license files etc

SolverStudio Readme
Files ReadMe.txt - this file
ChangeLog.txt

Application Files
SolverStudioTools.xlam - the SolverStudio data item editor Excel add-in. This is loaded into Excel by SolverStudio when first used during a session, and unloaded when Excel quits.
SolverStudio_XX_XX_XX_XX - the compiled SolverStudio add-in files

AMPL
init.py - language setup file for AMPL that adds AMPL-specific menu items
RunAMPL.py - the AMPL runner file used when solving AMPL files
InstallAMPLCML.py - used by init.py to install the free limited version of AMPL
(Optional - installed using a SolverStudio menu) amplcml - the installed copy of the limited free version of AMPL
Note: SolverStudio looks first for an installed version of AMPL, which it looks for by searching for ampl.exe in the entries in the environment variable 'path'. Only if this is not found, does SolverStudio look for ampl.exe in the SolverStudio folder ampl/amplcml/

AMPLNEOS
These files support AMPL solving on the NEOS server; no local copy of AMPL is needed for this

GAMS
init.py - language setup file for GAMS that adds GAMS-specific menu items
RunAMPL.py - the GAMS runner file used when solvimng GAMS models
NB: Various GAMS DLL's - used by SolverStudio to load and save GAMS GDX files are located in SupportFiles

Note: SolverStudio looks for an installed version of GAMS, which it looks for by searching for gams.exe in the entries in the environment variable 'path' and, if this fails, in a folder "gams*" within the program files folder(s).

GAMSNEOS
These files support GAMS solving on the NEOS server; no local copy of GAMS is needed. 

GurobiPython
init.py - language setup file
RunGurobiPython.py - the Gurobi runner file used when solving Gurobi files
Please see http://solverstudio.org for more details.

Solvers
Directory containing some ready to run solvers. SolverStudio uses either the 32bit or the 64bit folder by adding it to the path. ReadMe files for the solvers are also included. New solvers can be added here, and will be available in most (all?) languages including AMPL

SupportFiles
The DLL's needs for the ScintillaNet editor (both 32 and 64 bit, with ScintillaNet loading the right one), and 32 bit and 64 bit versions of the GAMS DLL's. As of 20150813, GAMs loads the right DLL for the process bitness. (We no longer have 32 anbd 64 bit folders with different dll's.)

A Mason
20150813
