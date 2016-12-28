# Updated 2012.07.20 to interpret quoted indices with spaces (by using csv.reader)
# Updated 20120814 to solve both GMPL and AMPL models
# This allows us to parse AMPL model files to distinguish between 
# singletons that can be params and single-member sets 
# Updated 2013.06.06 to implement a Pyomo work-around for errors when passing tables indexed by parameters into Pyomo.
# Updated 20140410 to read indexed data such as "flow[A B,1]" where indices are unquoted (and fail with index pairs like 'A,B',C)
# 20140620 Removed any :g15 formatting as this is locale specific, and so fails in non-English countries
# 20141007: Removed trailing space after "Sheet" in the Pyomo output file; there are reports of Pyomo failing on 64 bit systems that this may fix. 
# 20150917 Updated data file format to remove the brackets around the indices (which works with 2D indices, but not 3D or higher)

#language = "GMPL"
language = "Pyomo"

import csv
import re
import os
import sys
import clr
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import Form, Label, ToolStripMenuItem, MessageBox, MessageBoxButtons, DialogResult

# Converts singles to strings (outputting floats containing integers with no decimal point)
def fmtSngl(s):
   # Ensure integers such as 1.0 (showing in Excel as 1, but stored as floats) are output as integers
   # Integers are much more natural for indexing etc than 1.0, 2.0, 3.0 that happens otherwise
   if type(s) is float:
      if s.is_integer():
         return repr(int(s))
   return repr(s)

# Converts tuples to strings (with no surrounding brackets, but with commas)
def fmtTuple_CommaSeparated(s):
   r = fmtSngl(s[0])
   for i in s[1:]:
      r=r+", "+fmtSngl(i)
   return r

def WriteParamIndex(f,s):
   # Write an index for Pyomo, eg 22 or 1 2 "Auckland". Note: No brackets, no commas; these confuse Pyomo
   if type(s) is tuple:
      #f.write('['+','.join([fmt(t) for t in s])+']')
      # f.write('['+fmt(s)+']') # Fails
      for i in s:
         f.write(fmtSngl(i))
         f.write(" ")
   else:
      f.write(fmtSngl(s))

def WriteItem(f,s):
   # Write an item, eg an element of a set. Tuples are written as (1,2,"A"). Note that these do not work as indices in Pyomo.
   if type(s) is tuple:
      f.write( "("+fmtTuple_CommaSeparated(s)+")" )
   else:
      f.write(fmtSngl(s))

def WriteSnglParam(f,name,value):
   f.write("param "+name+" := "+fmtSngl(value)+";\n\n")

# set NUTR := A B1 B2 C ;
def WriteSet(f,name,value):
   f.write("set "+name+" := ")
   j = 0
   for i in value:
      if j>20:
         j=0
         f.write("\n")
      f.write(" ")
      WriteItem(f,i)
      j = j+1
   f.write(";\n\n")
  
# set NUTR := A B1 B2 C ;
def WriteSingletonSet(f,name,value):
   f.write("set "+name+" := ")
   j = 0
   WriteItem(f,value)
   f.write(";\n\n")
  
# param demand := 
#   2 6 
#   3 7 
#   4 4 
#   7 8; 
# param sc  default 99.99 (tr) :

# As of 7/6/2013, PYOMO will not read in lists of data with tuples as indices
# Hence, we write them as 2D tables instead
# def Write2DTable(f,name,value):
	# if int(value.BadIndexOption) == 2: # 2 = ReturnUserValue; TODO - refer to the enum correctly
		# f.write("param "+name+" default ");
		# WriteItem(f,value.BadIndexValue);
		# f.write(" :")
	# else:
		# f.write("param "+name+" :")
	# indexSet0 = set()
	# indexSet1 = set()
	# for (index,val) in value.iteritems():
		# indexSet0.add(index[0])
		# indexSet1.add(index[1])
	# for i in indexSet1:
		# f.write(" "+fmt(i))
	# f.write(" :=\n")
	# for index0 in indexSet0:
		# f.write(" ")
		# f.write(fmt(index0))
		# f.write(" ")
		# for index1 in indexSet1:
			# index = (index0,index1)
			# val = value.get(index)
			# if val != None:
				# f.write(fmt(val))
			# else:
				# f.write(".")
			# f.write(" ")
		# f.write("\n")
	# f.write(";\n\n")

def WriteIndexedParam(f,name,value):
	if int(value.BadIndexOption) == 2: # 2 = ReturnUserValue; TODO - refer to the enum correctly
		f.write("param "+name+" default ");
		WriteItem(f,value.BadIndexValue);
		f.write(" := ")
	else:
		f.write("param "+name+" := ")
	for (index,val) in value.iteritems():
			f.write("\n  ")
			WriteParamIndex(f,index)
			f.write(" ")
			if val != None:   # Skip empty cells (which can still get a value from a 'default' option above
				WriteItem(f,val)
			else:
				f.write(".")
	f.write("\n;\n\n")

# 2014: This is a work around for a Pyomo bug; we handle double-indexed tables directly as a special case; 
# higher order indices will fail until Pyomo is fixed
# 2015: We now write triple indices and higher with no brackets around the index; this works fine.
# def WriteIndexedParamPyomo(f,name,value):
	# (index,val)=value.items()[0]
	# if type(index) is not tuple:
		# WriteIndexedParam(f,name,value)
	# elif len(index)==1:
		# WriteIndexedParam(f,name,value)
	# # elif len(index)==2: # 20150921 No longer handle 2D tables as a special case
	# #	Write2DTable(f,name,value)
	# else:
		# WriteIndexedParam(f,name,value)	# 20150917 This now writes without brackets, and so works ok with Pyomo 4.1
		# # 2014 This may fail unless the user uses the latest Pyomo, and this Pyomo has been fixed!
		# #raise Exception("As of 2013.06.07, Pyomo cannot read in tables produced by SolverStudio with more than 2 values in an index tuple. The table '"+name+"' has "+str(len(index))+" values in its index.")

def ScanModelFileForDataItems(modelFileName):
   Sets= list();
   Params = list();
   lineCount = 0
   f = open(SolverStudio.ModelFileName,"r")
   for line in f:
      line = line.strip() # Remove white space, including the newline
      # Replace anything that can follow the name with a 'space' so the string splits correctly.
      l = line.replace("("," ").replace(";"," ").replace("<"," ").replace(">"," ").replace("{"," ").replace(":"," ").replace(","," ")
      lineCount += 1
      tokens = l.split()
      if len(tokens)<3: continue
      if tokens[2].lower()=="param":
         if "." not in tokens[0] or tokens[1] != "=":
            f.close()
            raise Exception("Unable to find the name for the "+tokens[0]+" defined by '"+line+"' on line "+str(lineCount)+" of the Pyomo model.")
         name = tokens[0][tokens[0].index(".")+1:]
         if not name in Params:
            # print "##  param ",tokens[1]
            Params.append(name)   
      elif tokens[2].lower()=="set":
         if "." not in tokens[0] or tokens[1] != "=":
            f.close()
            raise Exception("Unable to find the name for the "+tokens[0]+" defined by '"+line+"' on line "+str(lineCount)+" of the Pyomo model.")
         name = tokens[0][tokens[0].index(".")+1:]
         if not name in Sets:
            # print "##  set ",tokens[1]
            Sets.append(name)
   f.close()	
   return Sets, Params

# The main run code must be in a subroutine to avoid clashes with SolverStudio globals (which can cause COM exceptions in, eg, "for (name,value) in SolverStudio.SingletonDataItems.iteritems()" if value exists as a data item)
def DoRun():
	# Create a 'data' file, that contains the appropriate data
	dataFilename = SolverStudio.WorkingDirectory() + "\\SheetData.dat"
	g = open(dataFilename,"w")
	g.write("# "+language+" data file created by SolverStudio\n\n")

	print "## Scanning model for sets and parameters"
	Sets, Params = ScanModelFileForDataItems(SolverStudio.ModelFileName)
	#print "Sets=",Sets
	#print "Params=",Params

	print "## Building input file for %s sets and %s parameters"%(len(Sets),len(Params))

	print "## Writing simple parameters...",
	# Create constants
	count = 0
	for (name,value) in SolverStudio.SingletonDataItems.iteritems():
		 if name in Params:
			 count += 1
			 WriteSnglParam(g, name,value)
	print count,"items written."
			 
	print "## Writing sets...", #  Sets may be Python singletons or lists
	# Create sets
	count = 0
	for (name,value) in SolverStudio.SingletonDataItems.iteritems():
		 if name in Sets:
			 count += 1
			 WriteSingletonSet(g, name,value)
	for (name,value) in SolverStudio.ListDataItems.iteritems():
		 if name in Sets:
			 count += 1
			 WriteSet(g, name,value)
	print count,"items written."
		
	print "## Writing indexed parameters...",
	# Create indexed params
	count = 0
	for (name,value) in SolverStudio.DictionaryDataItems.iteritems():
		 if name in Params:
			 count += 1
			 WriteIndexedParam(g, name,value)
	print count,"items written."

	g.close()

	# Ensure the file "Sheet" that AMPL/GMPL writes its results into is empty
	# (We don't just delete it, as that require import.os to use os.remove)
	resultsfilename = "Sheet"
	g = open(resultsfilename,"w")
	g.close()

	print "## Running "+language+"..."
	print "## "+language+" model file:",SolverStudio.ModelFileName
	print "## "+language+" data file:",dataFilename+"\n"

	#try:
		# solver = "H:\\Desktop\\SolverStudio\\IronPython\\Scripts\\pyomo"
		# command = solver + " --solver=lpsolve --save-results=" + SolverStudio.WorkingDirectory() + "\\Sheet " + SolverStudio.ModelFileName + " " + dataFilename
		# command = command.replace("\\","\\\\")
		# errBool = os.system(command)
		# print errBool
		# succeeded = (errBool == 0)
		
		# import subprocess
		# pyomoPath = SolverStudio.GetPYOMOPath()
		# if pyomoPath != "":
			# command = pyomoPath + " --stream-solver --log -q --solver=lpsolve" + " --output=C:\\Temp\\pyomoOut.txt --save-results=\"" + SolverStudio.WorkingDirectory() + "\\Sheet \" \"" +  SolverStudio.ModelFileName + "\" \"" + dataFilename + "\""
			# subprocess.call(command)
		# print command
		#
		#try:
		#	sOptions=""
		#	options = SolverStudio.DataItems["PyomoOptions"]
		#	if "SingleIndexedRangeDictionary" in str(type(options)):
		#		try:
		#			hList=["","-","--"]
		#			for i in options:
		#				if i[0]=="":continue
		#				if i[0]!="-":
		#					if len(i)==1: h=1
		#					elif i[1]!="-":	h=2
		#				elif len(i)>2 and i[1]!="-": h=1
		#				else:
		#					h=0
		#				if options[i]==None:
		#					sOptions = sOptions + hList[h] + i.lower().strip() + " "
		#				else:
		#					sOptions = sOptions + hList[h] + i.lower().strip() + "=" + options[i].strip() + " "
		#		except:
		#			print "## There was an error with PyomoOptions."
		#	else:
		#		print "## PyomoOptions must contain two columns. RHS indexed by LHS. Leave RHS blank is no option needed."
		#except:
		#	pass
		#
	# sOptions="solve " We will need to add this "sub-command" when the Pyomo version increases, and this becomes compulsory. Without it we get a Warning at the moment
	sOptions="solve "
	r = None
	try:
		r = Application.ActiveSheet.Range("PyomoOptions")
	except:
		pass
	if r != None:
		if r.Columns.Count == 2:
			try:
				hList=["","-","--"]
				for row in range(r.Rows.Count):
					try:
						key = "(unknown)"
						key = r.Cells(row+1,1).Value2
						if key is None: continue
						key = str(key).strip().lower()  # convert all keys to lowercase for Pyomo; is this correct?
						if key=="": continue
						if key[0]!="-":
							if len(key)==1: h=1
							elif key[1]!="-":	h=2
						elif len(key)>2 and key[1]!="-": h=1
						else:
							h=0
						value = r.Cells(row+1,2).Value2
						if value!=None: value = str(value).strip()
						# 20152910: Force the solver name into lower case (TODO: do the same for other argunemts?)
						if value!=None and key=="solver": value = value.lower()
						if value==None or value=="":
							sOptions = sOptions + hList[h] + key.strip() + " "
						else:
							sOptions = sOptions + hList[h] + key.strip() + "=" + value + " "
					except Exception as inst:
						raise Exception( "##Error processing PyomoOption '"+key+"': "+str(inst))
			except:
				raise Exception("## There was an unknown error with PyomoOptions.")
		else:
			raise Exception("## PyomoOptions must be a table with two columns. Each row should contain an option such as 'solver' in column 1, and any associated value such as 'CBC' in column 2.")

	if not "solver=" in sOptions:
		# In previous versions before 2015.10.29, the solver was stored in upper case. Here, we force lower case to keep Pyomo happy.
		sOptions = sOptions + " --solver="+SolverStudio.ActiveModelSettings.GetValueOrDefault("COOPRSolver","cbc").lower() + " "

	sOptions=sOptions + " --json --save-results=\"" + SolverStudio.WorkingDirectory() + "\\Sheet\" \"" +  SolverStudio.ModelFileName + "\" \"" + dataFilename

	cPath = SolverStudio.GetPYOMOPath()
	if cPath == None:
		print "Unable to find pyomo.exe\nIf you have installed Pyomo, please add the full path to pyomo.exe (normally found in your Python directory inside Scripts) to the system path, or set the SolverStudio_PYOMO_PATH environment variable.\n"
		succeeded = False
	else:
		print "## Running: "+cPath+" "+sOptions
		exitCode = SolverStudio.RunExecutable(cPath, sOptions , addExeDirectoryToPath=True)
		# For some reason, in versions before Dec 2014, exitcode > 0 even though PYOMO succeeds, and so we ignore return code
		# if exitCode != 0: print "## Warning: Pyomo returned error code %d; this is common in pre-2015 versions" % exitCode
		import json
		try:
			with open("Sheet","r") as data_file:
				#sDict=data_file.read()
				myDict=json.load(data_file)
				#print myDict
				#print myDict["Solver"][0]["Error rc"]
			succeeded = myDict["Solver"][0]["Error rc"] = "0"
		except IOError as e: # Unable to open output file
			 print "## Error: No output file produced by Pyomo:",e
			 succeeded = False
		except ValueError as e: # Unable to read json
			 print "## Error: The Pyomo output file did not contain any valid data."
			 succeeded = False
		except Exception as e:
			 print "## Error reading Pyomo output:", e
			 succeeded = False
	#except:
	#	print "## Unable to call pyomo.exe"
	#	succeeded = False

	if succeeded:
		print "## "+language+" run completed."
		# with open("Sheet","r") as data_file:
			# sDict=data_file.read()
		Infinity = float('inf')
		# myDict = eval(sDict)
		if myDict["Solver"][0]["Termination condition"] == "infeasible":
			print "## Model is infeasible"
			succeeded = False
		else:
			# Read data such as:
			#"Variable": {
			#          "flow[A B,1]": {
			#              "Id": 1, 
			#              "Value": 500
			#          }, 
			#          "flow[A B,5]": {
			#              "Id": 6, 
			#              "Value": 700
			#          }, 
			# but see comment below on Pyomo bug fix for quoting of index elements
			returnVars = myDict["Solution"][1]["Variable"]
			if returnVars == "No values":
				print "## Pyomo returned 'No Values'"
				succeeded = False
		
	if succeeded:
		itemsLoaded = ""
		itemsCleared = []
		itemsChanged = False
		for varName in returnVars:
			# print varName
			changed = False
			value = float(returnVars[varName]["Value"])
			name = varName[:varName.index("[")]
			if not name in SolverStudio.DataItems:
				continue	# Skip this item, as it is not one we know about
			if not name in itemsCleared:
				try:
					for i in SolverStudio.DataItems[name]:
							SolverStudio.DataItems[name][i]=0
					itemsCleared.append(name)
				except: continue
			dataItem = SolverStudio.DataItems[name]
			items = varName[varName.index("[")+1:-1]

			# Read indices in things like "flow[A B,1]", as a tuple of ("A B",1) 
			# NB: As of 20140410, Pyomo does not quote indices, and so an index pair 'A,B',C is written (badly) as A,B,C
			# These used to be quoted correctly, hence the quote handling code below
			# This has been fixed in Pyomo 4; see https://software.sandia.gov/trac/pyomo/changeset/9624
			# So now, ' has been escaped to become \', and '...' added around any string containing a comma. We now use (below) a csv.reader() that implements escaping. Note: We have requested a further change to escape \, so \ becomes \\
			# if ord(items[0]) == 34: #Begins with "
				# items = items.replace("'","''").replace('"',"'")
			# if '"' in items: #Contains double quotes; replace all these by single quotes
			#	items = items.replace(',"',",'").replace('",',"',").replace('"',"'").replace("'","''").replace("'',''","','")
			#	if items[0:2]=="''": items=items[1:]
			#	if items[-2:]=="''": items=items[:-1]
			# AJM 20140410: Deleted the following line (perhaps because Pyomo has changed its output?)
			#elif not ord(items[0]) in (range(48,58)+[39,46]): #Single element
			#	items = "'" + items + "'"

			# index = csv.reader([items], delimiter=",", skipinitialspace=True, quotechar="'").next()
			index = csv.reader([items],delimiter=",", skipinitialspace=True, quotechar="'", doublequote=False, escapechar='\\').next() # New code for Pyomo 4 and earlier (but not too early!) versions
			
			#print "items=",items
			#print "index=",index
			
			# Convert any values into floats
			for j in range(len(index)):
				try:
					index[j]=float(index[j])
				except ValueError:
					pass

			if len(index)==1:
				index2 = index[0]  # The index is a single value, not a tuple
			else:
				index2 = tuple(index)# The index is a tuple of two or more items
			if dataItem[index2]!=value:
				dataItem[index2]=value
				changed = True
			if changed and not name in itemsLoaded:
				itemsLoaded = itemsLoaded+name+" "#"* "
				itemsChanged = True
			##elif not name in itemsLoaded:
		##		itemsLoaded = itemsLoaded+name+" "
		if itemsLoaded=="":
			print "## No results were loaded into the sheet."
		elif not itemsChanged:
			print "## Results loaded for data items:", itemsLoaded
		else :
			print "## Results loaded for data items:", itemsLoaded
		##	print "##   (*=data item values changed on sheet)"
	else:
		print "## "+language+" did not return a valid solution. The sheet has not been changed."
	print "## Done"

DoRun()