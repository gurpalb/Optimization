# SolverStudio AMPL NEOS Processor
# versionString = "v0.2, 2013.02.21"
# versionString = "v0.21, 2013.06.03" # Close the 'sheet' file
# versionString = "v0.22, 2013.07.01" # Don't try to write missing values (indicated by a ".")
# versionString = "v0.23, 2014.03.08" # Allow indexed sets, only set solver if not specified in model; fix for new NEOS file locations
#versionString = "v0.3, 2014.06.20" # Remove :g15 formatting, ensuring we get . in decimal numbers for all locales
#versionString = "v0.31, 2014.11.05" # Patched to allow NEOS timeouts, and convert "]]>" into "]] >" to make the model safe for inclusion in a CData (for which [[> is special)
#versionString = "v0.32, 2015.01.27" # Clearer reporting of timeouts wehn user has cancelled the thread, and reading in of string values
versionString = "v0.33, 2016.05.20" # Less frequent checking of NEOS status to avoid NEOS errors.
