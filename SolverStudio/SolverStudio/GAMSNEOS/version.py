# SolverStudio GAMSNEOS Processor
# versionString = "v1.0, 2013.02.21"
# versionString = "v1.1, 2013.07.01" # Moved NEOS warning into the first solve on NEOS
# versionString = "v1.2, 2014.03.08" # No longer copy GAMS DLL around; use new NEOS file locations# 
# versionString = "v1.21, 2014.11.04" # Bug fix: Correctly catch NEOS timeout exceptions
# 20141105 Convert any "]]>" into "]] >" in the model so it can be embedded in a CData object (for which ]]> is special)
# versionString = "v1.22, 2014.11.05" # "]]>" fix
# versionString = "v1.3, 2015.01.27" # Changed NEOS code to match RunAMPLNEOS.py with better error reporting etc
# versionString = "v1.4, 2015.06.15" # Better handling of errors when result .zip file contains the error msg reporting a missing file
versionString = "v1.41, 2016.05.20" #  # Changed interval between NEOS requests from 1 to 5 to avoid NEOs rejecting too many requests (20150520)
