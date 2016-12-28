import os
import sys
import clr
import shutil
import time
import json
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import Form, Label, ToolStripMenuItem, MessageBox, MessageBoxButtons, DialogResult

# Converts singletons (double and strings) to strings (outputting floats containing integers with no decimal point)
def fmtSngl(s):
   # Ensure integers such as 1.0 (showing in Excel as 1, but stored as floats) are output as integers
   # Integers are much more natural for indexing etc than 1.0, 2.0, 3.0 that happens otherwise
   if type(s) is float:
      # We could, but to match IronPython and CPython (but not AMPL), do not, convert integer-valued float to integer
      # if s.is_integer():
      #   return repr(int(s))
      #else:
         return repr(s)
   elif s == None: 
      return "nothing" # The closest Julia has to None/null
   elif type(s) is bool: 
      if s == True:
         return "true"
      else:
         return "false"
   # It must be a string; output it with double quotes (escaping any internal " and \)
   return json.dumps(s)

# Get string represenation of tuples to strings (with no surrounding brackets, but with a trailing , to handle tuples with 1 value)
def fmtTuple(s):
   r = "("
   for i in s:
      r=r+fmtSngl(i)+","
   return r+")"

# Write s as a text string suitable for Julia
def writeItem(f,s):
   if type(s) is tuple:
      #f.write( '('+','.join([repr(t) for t in s])+')' )
      f.write( fmtTuple(s) )
   else:
      f.write(fmtSngl(s))

# Code to write SolverStudio.jl, which will contain code to write JuliaResults.py		
def WriteDataFile(datafilename, resultsfilename):
	print "## Building Julia input file, 'SolverStudio.jl', for %s data items"%len(SolverStudio.DataItems)
	if SolverStudio.DataItems.Count>0: 
		sys.stdout.write("## Writing data items... ")
	writeAsModule = False # Warning; may not work as True unless we can have modules in modules?
	targetVersion3 = False
	targetVersion4 = False
	targetBothVersions = True
	g = open(datafilename, "w")
	g.write("# SolverStudio Data Items File for Julia 2\n# "+repr(SolverStudio.DataItems.Count)+" data items\n\n")
	if (targetBothVersions): g.write("using Compat # For Julia 0.3.7 and 0.4 compatibility\n\n")
	if writeAsModule:
		g.write("module SolverStudio\n\n")
		if SolverStudio.DataItems.Count>0: 
			g.write("export "+ (",".join(map(str,SolverStudio.DataItems.keys()))) + "\n\n")
	g.write('''# SSJuliaTools contains functions to write data itesm to the output file\ninclude("SSJuliaTools.jl")\n\n''')
	# Write out the current workbook fullname and path (tweaked from a suggestion by Leonardo). We roughly mimic the VBA syntax\
	g.write('ActiveWorkbook_Path='+json.dumps(ActiveWorkbook.Path)+'\n')
	g.write('ActiveWorkbook_FullName='+json.dumps(ActiveWorkbook.FullName)+'\n')
	g.write('ActiveWorkbook_Name='+json.dumps(ActiveWorkbook.Name)+'\n')
	g.write('ActiveSheet_Name='+json.dumps(ActiveSheet.Name)+'\n')
	# Write out the data items at the top of the file
	for (name,value) in SolverStudio.DataItems.iteritems():
		# Writing out a dictionary; note that Julia 0.3 and 0.4 syntax are different and incompatible
		if isinstance(value,dict):
			if (targetVersion3): g.write(name + " = {") # Deprecated in Julia 0.4; use Dict{Any,Any{(4.0=>'i',5.0=>'j',3.0=>'h',2.0=>'f',1.0=>'g'), not {4.0=>'i',5.0=>'j',3.0=>'h',2.0=>'f',1.0=>'g'}
			if (targetVersion4): g.write(name + " = Dict{Any,Any}(") # Julia 0.4 syntax; not compatible with 0.3!
			if (targetBothVersions): g.write(name + " = @compat Dict{Any,Any}(") # Julia 0.4 syntax, compatible with 0.3 via Compat module
			for (i, v) in value.iteritems():
				writeItem(g,i) # Writing out a tuple in a (firstIndex,secondIndex) format
				g.write(" => ")
				writeItem(g,v) # Writing out a tuple in a (firstIndex,secondIndex) format
				g.write(", ")
			# Finished writing a dictionary
			if (targetVersion3): g.write("}\n") # Deprecated in Julia 0.4
			if (targetVersion4 or targetBothVersions): g.write(")\n") # Not compatible with v0.3
		# Writing out a list
		elif isinstance(value,list):
			g.write(name + " = [")
			for it in list(value):
				writeItem(g,it)
				g.write(", ")
			g.write("]\n")
		else:
			g.write(name + " = ")
			writeItem(g,value)
			g.write("\n")
	#
	# Now write the supporting code
	g.write("\n\n# SolverStudio Julia Support Code.............................................\n\n")
	# g.write("module SolverStudioSupport\n") # We do NOT put this in a module as running interactively, we get a warning about replacing the module
	# g.write("export GenerateSolverStudioResultsFile\n\n")
	#for name in SolverStudio.DataItems.keys()[0:-1]:
	#	g.write(name + ", ")
	#g.write(SolverStudio.DataItems.keys()[-1] + "\n\n")
	
	# Removed by AJM on 20160408 as we use the SSJuliaTools file instead.
	# g.write("function SS_WriteSolverStudioVariable(io,var)\n")
	# g.write('''	write(io,string(repr(var), "\\r\\n"))''' + "\n")
	# g.write("end\n\n")
	# g.write("function SS_WriteSolverStudioVariable(io,var::Bool)\n")
	# g.write("\tif var\n")
	# g.write("\t\t" + '''	write(io,"True\\r\\n")''' + "\n")
	# g.write("\telse\n")
	# g.write("\t\t" + '''	write(io,"False\\r\\n")''' + "\n")
	# g.write("\tend\n")
	# g.write("end\n\n")
	
	# g.write("try\n")
	# g.write("\tusing JuMP\n")
	# g.write("\tfunction SS_WriteSolverStudioVariable(io, var::JuMP.JuMPContainer)\n")
	# g.write("\t\t" + '''write(io, "{")''' + "\n")
	# g.write("\t\tfor it in var\n")
	# g.write("\t\t\tif ndims(var) == 1\n")
	# g.write("\t\t\t\t" + '''write(io, string(repr(it[1]), " : ", repr(it[2]), ", "))''' + "\n")
	# g.write("\t\t\telse\n")
	# g.write("\t\t\t\tindices = it[1:end-1]\n")
	# g.write("\t\t\t\tvalue = it[end]\n")
	# g.write("\t\t\t\t" + '''write(io, string("(",join(map(repr,indices),","),") : ", repr(value), ", "))''' + "\n")
	# g.write("\t\t\tend\n")
	# g.write("\t\tend\n")
	# g.write("\t\t" + '''write(io, "}\\r\\n")''' + "\n")
	# g.write("\tend\n")
	# g.write("end\n\n")

	# g.write("try\n")
	# g.write("\tusing JuMP\n")
	# g.write("\tfunction SS_WriteSolverStudioVariable(io, var::JuMP.JuMPArray)\n")
	# g.write("\t\t" + '''write(io, "{")''' + "\n")
	# g.write("\t\tfor it in keys(var)\n")
	# g.write("\t\t\tif length(it)==1\n")
	# g.write("\t\t\t\t" + '''write(io, string(repr(it[1]), " : ", repr(var[it...]), ", "))''' + "\n")
	# g.write("\t\t\telse\n")
	# g.write("\t\t\t\t" + '''write(io, string(repr(it), " : ", repr(var[it...]), ", "))''' + "\n")
	# g.write("\t\t\tend\n")
	# g.write("\t\tend\n")
	# g.write("\t\t" + '''write(io, "}\\r\\n")''' + "\n")
	# g.write("\tend\n")
	# g.write("end\n\n")

	# g.write("function SS_WriteSolverStudioVariable(io, var::Dict)\n")
	# g.write('''	write(io, "{ ")''' + "\n")
	# g.write("	for (index, value) in var\n")
	# g.write("\t\t" + '''write(io, string(repr(index), " : ", value == nothing ? "None" : repr(value), ", "))''' + "\n")
	# g.write("	end\n")
	# g.write('''	write(io, "}\\r\\n")''' + "\n")
	# g.write("end\n\n")
	# g.write("function SS_WriteSolverStudioVariable(io, var::Array)\n")
	# g.write('''	write(io, "[ ")''' + "\n")
	# g.write("	for value in var\n")
	# g.write("\t\t" + '''write(io, string(repr(value), ", "))''' + "\n")
	# g.write("	end\n")
	# g.write('''	write(io, "]\\r\\n")''' + "\n")
	# g.write("end\n\n")
	
	g.write("function SS_GenerateSolverStudioResultsFile()\n") # This name must match that given in init.py
	g.write('''	file = open("''' + resultsfilename + '''", "w")''' + "\n")
	g.write("	try\n")
	g.write('''		write(file,"# Julia Result File written by SolverStudio\\r\\n")''' + "\n") # We must always write something into the file; a file lenght of 0 could confuse the code
	g.write('''		write(file,"NaN=None # Allow julia NaN to be read as None\\r\\n")''' + "\n")
	for name in SolverStudio.DataItems.iterkeys():
		g.write('''		write(file,"''' + name + ''' = ")''' + "\n")
		if isinstance(SolverStudio.DataItems[name],dict):
			g.write("		SS_WriteIndexedSolverStudioVariable(file, Main." + name + ")\n")
		else:
			g.write("		SS_WriteSolverStudioVariable(file, Main." + name + ")\n")
	g.write("	finally\n")
	g.write("		close(file)\n")
	g.write("	end\n")
	g.write("end\n")
	# g.write("end # end of SolverStudioSupport module\n")
	g.write("\natexit(SS_GenerateSolverStudioResultsFile)")
	if writeAsModule:
		g.write("\n\nend")
	g.close()
	print "SolverStudio.jl written"

def DoRun():
	datafilename = "SolverStudio.jl"
	resultsfilename = "JuliaResults.py"

	#print "## Executing "+__file__ 
	WriteDataFile(datafilename, resultsfilename)
	shutil.copy2(str(SolverStudio.LanguageDirectory()) + "\SSJuliaTools.jl", str(SolverStudio.WorkingDirectory()))
	exeFile = SolverStudio.GetJuliaPath(mustExist = True);
	
	# Delete any existing results file so we can be sure we do not open an old file
	#g = open(resultsfilename,"w")
	#g.close()
	if os.path.exists(resultsfilename):
		os.remove(resultsfilename)

	#if not os.path.isfile(exeFile) :
	#	sys.exit("ERROR: Julia could not be found")

	# print "os.getcwd()=",os.getcwd() 
	runInteractively=(SolverStudio.GetRegistrySetting("Julia","RunInConsole",0) == 1)
	if runInteractively:
		interactiveExe = SolverStudio.InteractiveExes.GetInteractiveExe("Julia")
		if not interactiveExe.IsRunning:
			print "## Starting Julia ("+interactiveExe.Executable+") console"
			interactiveExe.Start()
			# print "Working Dir =",interactiveExe.WorkingDirectory;
		print "## Running file",SolverStudio.ModelFileName,"in console"
		commandToSend = "include(\""+SolverStudio.ModelFileName.encode("string_escape")+"\"); "   # execute the model file; note that we convert, eg, c:\temp\.. into c:\\temp\\..
		commandToSend = commandToSend + "SS_GenerateSolverStudioResultsFile(); print(\"\");"   # Write the results (as would be down by atexit if not running interactively). The print clears any output by clearing ans
		interactiveExe.SendCommandAndWaitForCompletion(commandToSend)
	else:
		print "## Running Julia ("+exeFile+")"
		print "## with file:",SolverStudio.ModelFileName+"\n"
		exitCode = SolverStudio.RunExecutable(exeFile,'"'+SolverStudio.ModelFileName+'"', killChildProcesses = True)
		if not (exitCode==0):
			print "## Julia did not complete (Error Code %d); no solution is available." % exitCode
			return

	print "## Julia run completed. Reading results..."
	# Making sure JuliaResults.py is completely written before continuing
	# Wait for the file to exist (which should always be immediately). If it doesn't exist after 1s, assume it never will 
	# This can happen if the user does not have "using SolverStudio" for example.
	for i in xrange(40):
		if os.path.exists(resultsfilename):
			break
		time.sleep(5.0/1000.0)
	else:
		print "## No results were produced, and so no results were loaded into the sheet."
		return

	#	 time.sleep(5.0/1000.0)
	#while os.stat(resultsfilename).st_size==0:
	#	time.sleep(5.0/1000.0) # Wait 5 ms
	#	pass
	# Now wait until the file is able to be opened for read/write, meaning Julia must have finished creating it
	for i in xrange(50):
		 try:
			  with open(resultsfilename, 'r+b') as _:
					break
		 except IOError:
			time.sleep(5.0/1000.0)
	else: # execute this if the for loop finishes (without a break) after 10,000 loops
		 #raise IOError('Could not access {} after {} attempts'.format(resultsfilename,10000))
		 raise IOError('Could not access the Julia results file "'+resultsfilename+'" because it never became available for reading.')

	import JuliaResults as results
	g=results.__dict__

	itemsLoaded = ""
	itemsChanged = False
	for (name,dataItem) in SolverStudio.DataItems.iteritems():
		if name in g:
			if isinstance(dataItem,float) or isinstance(dataItem,int) or isinstance(dataItem,str) or dataItem==None: 
				if isinstance(g[name],float) or  isinstance(g[name],int) or isinstance(g[name],str) or g[name]==None:
					if globals()[name]==g[name]:
						itemsLoaded = itemsLoaded+name+" "
					else:
						globals()[name]=g[name] # This is a singleton, so needs to be written directly
						itemsChanged=True
						itemsLoaded = itemsLoaded+name+"* "
			elif isinstance(dataItem,dict):
				if isinstance(g[name],dict):
					changed = False
					for index in g[name].iterkeys():
						if dataItem[index]!=g[name][index]:
							dataItem[index]=g[name][index]
							changed = True
					if changed:
						 itemsLoaded = itemsLoaded+name+"* "
						 itemsChanged = True
					else:
						itemsLoaded = itemsLoaded+name+" "
	if itemsLoaded=="":
		print "## No results were loaded into the sheet."
	elif not itemsChanged:
		print "## Results loaded for data items:", itemsLoaded
	else :
		print "## Results loaded for data items:", itemsLoaded
		print "##   (*=data item values changed on sheet)"

DoRun()
print "## Done"
