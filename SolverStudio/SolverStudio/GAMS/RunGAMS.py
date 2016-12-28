# Updated 20140518 to remove writing to c:/temp/t of a test output file!
# Updated 20140730 to redirect GAMS output to STDOUT using LO=3

import os
import re
from shutil import copy2

def ScanModelFileForDataItems(modelFileName):
	unusedKeys = ["positive","equations","model","solve","scalar"]#,"display"]
	Sets= list()
	Params = list()
	Tables = list()
	AllVars = list()
	Variables={}
	dispBool = False
	#singlelineKeyWords = { "set" : Sets, 'parameter' : Params, 'table' : Tables }
	keyWords = { "sets" : Sets, 'parameters' : Params, 'tables' : Params, "set" : Sets, 'parameter' : Params, 'table' : Params,"variables":AllVars }
	lineCount = 0
	f = open(SolverStudio.ModelFileName,"r")
	items = None
	for line in f:
		line = line.strip() # Remove white space, including the newline
		# Replace anything that can follow the name with a 'space' so the string splits correctly.
		l = line.replace("("," ").replace(";"," ; ").replace(","," ") #.replace("<"," ").replace(">"," ").replace("{"," ").replace(":"," ").replace(","," ")
		lineCount += 1
		tokens = l.split()
		if len(tokens)==0: continue
		if tokens[0].lower() in keyWords.keys():
			items = keyWords[tokens[0].lower()]
			tokens = tokens[1:]
		elif tokens[0].lower() in unusedKeys:
			items = None
		if len(tokens)==0: continue
		if items <> None:
			if tokens[0]<>";" and not tokens[0] in items:
				# print "##  param ",tokens[1]
				if tokens[0] in SolverStudio.DataItems:
					items.append(tokens[0])	
		if tokens[0]=="display":
			dispBool = True
			for i in tokens[1:]:
				if i == ";": break
				level = 0
				if len(i)>2:
					if i[-2]==".": 
						if i[-1]=="m": level = 1 
						i=i[:-2]			
				if i in Variables:
					raise Exception("\nUnable to write two instances of the variable " + i + " to the sheet. If you are trying to display values and marginals create data items var_l and var_m.")
				if i in SolverStudio.DataItems:
						Variables[i]=level
				else:
					print "## Unable to write " + i + " to sheet. Create a data item first"

				# elif i in AllVars:
					# if i in SolverStudio.DataItems:
						# Variables[i]=level
					# else:
						# print "## Unable to write " + i + " to sheet. Create a data item first"
				# else:
					# print "## " + i + " is not a variable. Did not write to sheet"
		#The command ends with a ; so search for this
		if ";" in tokens:
			items = None
	f.close()
	if not dispBool:
		for i in AllVars:
			if i in SolverStudio.DataItems:
				Variables[i]=0
	print "Sets=",Sets
	print "Params=",Params
	print "Variables=",AllVars
	return Sets, Params,Variables

def FixInput(model,Sets,Params):
	sList = ""
	for i in Sets+Params:
		sList = sList + i + ", "
	r=None
	l=re.compile(r'^\$\s*GDXIN\s*"SheetData.gdx".*\n*^\$\s*LOAD.*',re.IGNORECASE | re.MULTILINE)
	if not re.search(l,model):
		r1=re.compile(r'^\$\s*GDXIN\s*"SheetData.gdx".*',re.IGNORECASE | re.MULTILINE)
		r2=re.compile(r'^\s*\$\s*include\s*"SheetData.dat".*',re.IGNORECASE | re.MULTILINE)
		if re.search(r1,model): r=r1
		if re.search(r2,model): r = r2
		if r:
			model=r.sub('\n$GDXIN "SheetData.gdx"\n***SolverStudio: Above Line Modified\n$LOAD '+sList[:-2]+"\n***SolverStudio: Above Line Modified\n$GDXIN\n***SolverStudio: Above Line Modified\n",model)
		else:
			r=re.compile(r'^\s*solve ',re.IGNORECASE|re.MULTILINE)
			model=r.sub('\n$GDXIN "SheetData.gdx"\n***SolverStudio: Above Line Modified\n$LOAD '+sList[:-2]+"\n***SolverStudio: Above Line Modified\n$GDXIN\n***SolverStudio: Above Line Modified\nsolve ",model)
	return model
	
def FixOutput(model,Variables):
	r=re.compile(r'^display\s*',re.IGNORECASE | re.MULTILINE)
	sVars = ""
	if re.search(r,model):
		model = r.sub("execute_unload 'results.gdx' ",model)
	else:
		for i in Variables:
			sVars=sVars+i+", "
		model = model + "\nexecute_unload 'results.gdx', " + sVars[:-2] + ";\n***SolverStudio: Line Modified"
	return model

def DoRun():
	if SolverStudio.GetGAMSPath(False) == None:
		raise Exception("\nGAMS does not appear to be installed on this computer.")
			
	print "## Scanning model for sets and parameters"
	Sets, Params,Variables = ScanModelFileForDataItems(SolverStudio.ModelFileName)

	#++++++++++++++++++++++
	with open(SolverStudio.ModelFileName,"r") as inFile:
			model = inFile.read()
			
	if SolverStudio.ActiveModelSettings.GetValueOrDefault("FixMinorErrors",True):
		model = FixInput(model,Sets,Params)
		
	model = FixOutput(model,Variables)

	with open(SolverStudio.ModelFileName,"w") as outFile:
			outFile.write(model)
	#+++++++++++++++++++++
	#with open("C:\\Temp\\t","w") as test:
	#		test.write(model)

	dataFilename = "SheetData.gdx"
	print "## Building GAMS input file for %s sets and %s parameters"%(len(Sets),len(Params))
	PGX = SolverStudio.OpenGDX(dataFilename,False)
	for i in Sets:
		SolverStudio.WriteSetToGDX(PGX,i,SolverStudio.DataItems)
	for i in Params:
		SolverStudio.WriteParamToGDX(PGX,i,SolverStudio.DataItems)
	SolverStudio.CloseGDX(PGX)

	# Ensure the file "results.gdx" that AMPL writes its results into is empty
	# (We don't just delete it, as that require import.os to use os.remove)
	resultsfilename = "results.gdx"
	g = open(resultsfilename,"w")
	g.close()

	print "## Running GAMS ("+SolverStudio.GetGAMSPath()+")"
	print "## GAMS model file:",SolverStudio.ModelFileName
	print "## GAMS data file:",dataFilename
	exitCode = SolverStudio.RunExecutable(SolverStudio.GetGAMSPath()," \""+SolverStudio.ModelFileName+"\"  lo=3")
	  # LO=3 forces GAMS to output to STDOUT instead of the console

	if exitCode==0:
		print "\n## GAMS completed successfully.\n## Reading results..."
		PGX = SolverStudio.OpenGDX(resultsfilename)
		itemsLoaded = list()
		for i in Variables:
			try:
				SolverStudio.ReadSymbolFromGDX(PGX,i,Variables[i],SolverStudio.DataItems)
				itemsLoaded.append(i)
			except:
				print "## Unable to load " + i
		SolverStudio.CloseGDX(PGX)
		if len(itemsLoaded)==0:
			print "## No results were loaded into the sheet."
		else:
			print "## Results loaded for data items:", itemsLoaded
	else:
		print "## GAMS did not complete (Error Code %d); no solution is available." % exitCode
	print "## Done"
	
DoRun()
