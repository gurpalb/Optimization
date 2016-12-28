# SolverStudio COOPR Processor
#versionString = "v1.0, 2013.02.21"
#versionString = "v1.1, 2013.06.06"
#versionString = "v1.2, 2013.06.10"
#versionString = "v1.3, 2013.06.25" # Fixed missing quotes in 2D tables
#versionString = "v1.31, 2013.06.26" # Changed previous fix to use repr(), not just add quotes
#versionString = "v1.4, 2013.06.26" # Changed previous fix to use repr(), not just add quotes, and to write integers without decimal pts
# versionString = "v1.5, 2014.04.10" # Changed parser of Pyomo output to read current Pyomo version output (but note that this does not quote indices containing a comma
# versionString = "v1.6, 2014.10.09" # Removed trailing space on Pyomo "Sheet" output file, removed printing of var names, and specified --json output; open file explicitly as JSON
# versionString = "v1.7, 2015.01.21" # Switched from COOPR to Pyomo as name; updated code to handle Pyomo fix #4525
#versionString = "v1.8, 2015.09.17" # Removed brackets around indices
versionString = "v1.9, 2015.10.29" # Convert solver name to lower case before passing it on the command line; this fixes any solvers stored by previous versions in upper case.
