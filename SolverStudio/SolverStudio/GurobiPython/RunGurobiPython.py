# This file handles running Gurobi code within OpenSolver Studio
#
# It firstly writes out a file "SolverStudio.py" that is a module ready to be imported
# by the Gurobi file. This contains all data items defined on the sheet.
#
# The user file, which has been written out as "model.py", is then run using Gurobi. This
# imports the SolverStudio.py module, giving access to the data items.
#
# The user file modifies the data items, quits, and an atexit handler in SolverStudio.py
# writes out the modified variables to a file "GurobiResults.py"
#
# This file then loads "GurobiResults.py" as a module, getting access to the new data items values,
# and then writes these values  back to the sheet
#
# 2012.02.15: Now reports as "loaded" only those variables that actually have new values
# 2012.02.22: Writes out indexed parameters (dictionaries) which have default values as 

# The main run code must be in a subroutine to avoid clashes with SolverStudio globals (which can cause COM exceptions in, eg, "for (name,value) in SolverStudio.SingletonDataItems.iteritems()" if value exists as a data item)
def DoRun():
	print "## Executing "+__file__
	print "## Building Gurobi input file, module 'SolverStudio.py', for %s data items"%len(SolverStudio.DataItems)

	# We will write out a results file called ...
	resultsfilename = "GurobiResults.py"

	# Create a 'data' file, that contains if statements to read in the appropriate data
	filename = "SolverStudio.py"
	g = open(filename,"w")
	g.write("# Python data file created by RunGurobiPython.py in OpenSolver Studio\n\n")

	#create a docstring
	g.write('''""" OpenSolver Studio Spreadsheet Data Module
This module provides access to the data items on the current Excel sheet.
The items are as follows:
''');
	for (key,value) in SolverStudio.DataItems.iteritems():
		s = repr(value)
		if len(s)>30:
			s = s[0:30]+"..."
		g.write("   "+key+" = "+s+"\n")
	g.write('"""\n\n');

	g.write('''# The following supports defaultdict which allows default values for dictionaries. 
# We ignore errors at this stage as we may not need to use a defaultdict
try:
   import itertools
   import collections
   def constant_factory(value):
      return itertools.repeat(value).next
except:
   pass
''')

	dataName = "GurobiData.py"
	d = open(dataName,"w")
	d.write("# Python data file created by RunGurobiPython.py in SolverStudio\n\n")
	# Write the data items. (We could use isinstance(value,str) to get the specific types, 
	# or we could iterate thru the dicts of each type
	print "## Writing data items"
	# Create data items in a module
	for (name,value) in SolverStudio.DataItems.iteritems():
		 if not isinstance(value,dict):
			 d.write(name+" = "+repr(value)+"\n")
		 elif value.BadIndexOption == value.BadIndexOptions.ThrowError:
			 # This is a dictionary with no default value: create a standard dictionary
			 d.write(name+" = "+repr(value)+"\n")
		 elif value.BadIndexOption == value.BadIndexOptions.ReturnUserValue:
			 # This dictionary has a default value
			 d.write(name+" = collections.defaultdict(constant_factory("+repr(value.BadIndexValue)+"),"+repr(value)+")\n")
		 elif value.BadIndexOption == value.BadIndexOptions.ReturnPythonNone:
			 # This dictionary has a default value of None
			 print "Error: Returning None for an unrecognised index is not yet implemented."
	d.close()

	g.write('\nfrom GurobiData import *\n')

	# Create a method to write the data to an output file ready to read back into the sheet
	# We first get a copy of global _main__ dictionary
	g.write('''

# Create a dummy variable that we will use to test for: from SolverStudio import *
SolverStudio_GlobalTester = 0

# Get a copy of the main global dictionary; doing this in an atexit call is too late as the variables have been deleted
import __main__
g = __main__.__dict__

def WriteData():
   """This writes out all the data items when the Python program finishes"""
   f=open("'''+resultsfilename+'''","w")
   f.write("# '''+resultsfilename+''' - output file generated when running a Gurobi file from OpenSolver Studio\\n")
   if not "SolverStudio_GlobalTester" in g:
      # The user has gone "import SolverStudio", and so the variables are SolverStudio.
''')
	# Write lines like...
	#      f.write("Test2 = "+repr(Test2)+"\n")
	for (name,value) in SolverStudio.DataItems.iteritems():
		 if isinstance(value,dict) and value.BadIndexOption == value.BadIndexOptions.ReturnUserValue:
			 # This will have been created as a defaultdict; cast it into a simple dict
			 g.write('      f.write("'+name+' = "+repr(dict('+name+'))+"\\n")\n')
		 else:
			 g.write('      f.write("'+name+' = "+repr('+name+')+"\\n")\n')
	g.write('''   else:
      # The user has gone: "from SolverStudio import *", so we get all our variables from the globals
      # That is the only way to get new values for strings, ints etc
''')
	# Write lines like...
	#       f.write("Test2 = "+repr(g["Test2"])+"\n")
	for (name,value) in SolverStudio.DataItems.iteritems():
		 if isinstance(value,dict) and value.BadIndexOption == value.BadIndexOptions.ReturnUserValue:
			 # This will have been created as a defaultdict; cast it into a simple dict
			 g.write('      f.write("'+name+' = "+repr(dict(g["'+name+'"]))+"\\n\")\n')
		 else:
			 g.write('      f.write("'+name+' = "+repr(g["'+name+'"])+"\\n\")\n')
	g.write("   f.close()\n\n")

	# Create an at-exit handler to write the data items to an output
	g.write("import atexit; atexit.register(WriteData)\n\n")

	# # Add the SolverStudio PuLP to the path (at the end of the path so it is a last resort)
	# g.write("# Add SolverStudio's PuLP to the end of the Python path so that the user can import pulp\n")
	# g.write("import sys\n")
	# g.write("sys.path.append(r\""+SolverStudio.GetPuLPSourcePath(mustExist = False)+"\")\n")

	g.close()

	# Ensure the file "Sheet" that Gurobi writes to is empty
	# (We don't just delete it, as that require import.os to use os.remove)
	g = open(resultsfilename,"w")
	g.close()

	#import os
	#command = "\"" + SolverStudio.GetGurobiPath() + "\" \"" + SolverStudio.ModelFileName + "\""
	#succeeded = not os.system(command)

	exitCode = SolverStudio.RunExecutable(SolverStudio.GetGurobiPath(),'"'+ SolverStudio.ModelFileName + '"')

	if exitCode==0:
		print "## Gurobi run completed successfully."
		# Read a file such as: 
		#flow :=
		#A 1   0
		#B 5   0
		#;
		itemsLoaded = ""
		# import resultsfilename as results
		import GurobiResults as results
		g=results.__dict__
		for (name,dataItem) in SolverStudio.DataItems.iteritems():
			if name in g:
				# Process singletons first
				if isinstance(dataItem,float) or isinstance(dataItem,int) or isinstance(dataItem,str) or dataItem==None: 
					if isinstance(g[name],float) or  isinstance(g[name],int) or isinstance(g[name],str) or g[name]==None:
						if globals()[name]!=g[name]:
							globals()[name]=g[name] # This is a singleton, so needs to be written directly
							itemsLoaded = itemsLoaded+name+" "
					else:
						raise TypeError("Unable to assign the new value '"+repr(g[name])+"' to the single-cell data item '"+name+"'.")
				# Process indexed items, copying them over value-by-value (only where values differ because writing to a sheet is slow, but reads are cached and so fast)
				elif isinstance(dataItem,dict):
					if isinstance(g[name],dict):
						changed = False
						for index in g[name].iterkeys(): # Iterate thru the new keys so we spot and report (as an error) any added keys
							if dataItem[index]!=g[name][index]:
								dataItem[index]=g[name][index]
								changed = True
						if changed:
							itemsLoaded = itemsLoaded+name+" "
					else:
						raise TypeError("Unable to assign the new value '"+repr(g[name])+"' (which is not a dictionary) to the indexed data item '"+name+"'.")
				else:
					#print name," is a ",type(dataItem)
					pass
		if itemsLoaded=="":
			print "## No new values were loaded into the sheet."
		else:
			print "## New values loaded for data items:", itemsLoaded
	else:
		print "## Gurobi did not complete (Error Code %d); no solution is available." % exitCode
	print "## Done"

DoRun()
