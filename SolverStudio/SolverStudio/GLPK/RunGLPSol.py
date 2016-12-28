# Updated 2012.07.20 to interpret quoted indices with spaces (by using csv.reader)
# Updated 20120814 to solve both GMPL and AMPL models
# This allows us to parse AMPL model files to distinguish between 
# singletons that can be params and single-member sets 
#
# v1.41, 2013.06.03; Closed the results file
# v1.5: Removed all "{:g15}".format(i) as this is locale dependent, and fails with , replacing . in some countries.

language = "GMPL"
#language = "AMPL"

import csv
import re
import os
import sys
import clr

# Converts singles to strings (outputting floats containing integers with no decimal point)
def fmtSngl(s):
   # Ensure integers such as 1.0 (showing in Excel as 1, but stored as floats) are output as integers
   # Integers are much more natural for indexing etc than 1.0, 2.0, 3.0 that happens otherwise
   if type(s) is float:
      if s.is_integer():
         return repr(int(s))
   return repr(s)

# Converts tuples to strings (with no surrounding brackets)
def fmtTuple(s):
   r = fmtSngl(s[0])
   for i in s[1:]:
      r=r+", "+fmtSngl(i)
   return r

def fmt(s):
   if type(s) is tuple:
      #f.write('['+','.join([repr(t) for t in s])+']')
      return fmtTuple(s)
   else:
      return fmtSngl(s)

# Return s as a string suitable for indexing, which means tuples are listed items in []
def WriteIndex(f,s):
   if type(s) is tuple:
      #f.write('['+','.join([repr(t) for t in s])+']')
      f.write('['+fmt(s)+']')
   else:
      f.write(fmt(s))

# Return s as a string suitable for an item in a set, which means tuples are listed items in ()
def WriteItem(f,s):
   if type(s) is tuple:
      #f.write( '('+','.join([repr(t) for t in s])+')' )
      f.write( fmt(s) )
   else:
      f.write( fmt(s))

def WriteParam(f,name,value):
   f.write("param "+name+" := "+fmt(value)+";\n\n")

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
  
# set NUTR := A ;
def WriteSingletonSet(f,name,value):
   f.write("set "+name+" := ")
   j = 0
   WriteItem(f,value)
   f.write(";\n\n")

def WriteIndexedSet(f,name,value):
   for (index,val) in value.iteritems():
      f.write("set ")
      f.write(name)
      f.write("[")
      f.write(fmt(index))
      f.write("] := ")
      if val != None and val != ():   # Write empty rows (which end up as an empty tuple, hence "()" in code) as empty sets (by writing nothing)
         f.write(fmt(val)) # We write values separated by commas...
      f.write(";\n")
   f.write("\n")
  
# param demand := 
#   2 6 
#   3 7 
#   4 4 
#   7 8; 
# param sc  default 99.99 (tr) :

def WriteIndexedParam(f,name,value):
   if int(value.BadIndexOption) == 2: # 2 = ReturnUserValue; TODO - refer to the enum correctly
      f.write("param "+name+" default ");
      WriteItem(f,value.BadIndexValue);
      f.write(" := ")
   else:
      f.write("param "+name+" := ")
   for (index,val) in value.iteritems():
      if val != None:   # Skip empty cells (which can still get a value from a 'default' option above
         f.write("\n  ")
         WriteIndex(f,index)
         f.write(" ")
         WriteItem(f,val)
   f.write("\n;\n\n")
   
def ScanModelFileForDataItems(modelFileName):
   Sets= list();
   Params = list();
   lineCount = 0
   f = open(SolverStudio.ModelFileName,"r")
   for line in f:
      line = line.strip() # Remove white space, including the newline
      # Replace anything that can follow the name with a 'space' so the string splits correctly.
      l = line.replace("="," ").replace(";"," ").replace("<"," ").replace(">"," ").replace("{"," ").replace(":"," ").replace(","," ")
      lineCount += 1
      tokens = l.split()
      if len(tokens)==0: continue
      if tokens[0]=="#": continue # skip comments
      if tokens[0]=="data":
          raise Exception("When using GMPL in SolverStudio, you should not include a 'data SheetData.dat' statement. Please delete this and try again.\n\n(The error occurred in the string '"+repr(line)+"' on line "+repr(lineCount)+" of file "+repr(modelFileName)+".)")
      if tokens[0].lower()=="param":
         if len(tokens)<2:
            f.close()
            raise Exception("Unable to find the name for the "+tokens[0]+" defined by '"+line+"' on line "+str(lineCount)+" of the GMPL model.")
         if not tokens[1] in Params:
            # print "##  param ",tokens[1]
            Params.append(tokens[1])   
      elif tokens[0].lower()=="set":
         if len(tokens)<2:
            f.close()
            raise Exception("Unable to find the name for the "+tokens[0]+" defined by '"+line+"' on line "+str(lineCount)+" of the GMPL model.")
         if not tokens[1] in Sets:
            # print "##  set ",tokens[1]
            Sets.append(tokens[1]) 
   f.close()
   return Sets, Params

# The main run code must be in a subroutine to avoid clashes with SolverStudio globals (which can cause COM exceptions in, eg, "for (name,value) in SolverStudio.SingletonDataItems.iteritems()" if value exists as a data item)
def DoRun():

	# Create a 'data' file, that contains the appropriate data
	dataFilename = "SheetData.dat"
	g = open(dataFilename,"w")
	g.write("# "+language+" data file created by SolverStudio\n\n")
	g.write("data;\n\n")

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
			 WriteParam(g, name,value)
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
	for (name,value) in SolverStudio.DictionaryDataItems.iteritems():
		if name in Sets:
			count += 1
			WriteIndexedSet(g, name,value)
	print count,"items written."
		
	print "## Writing indexed parameters...",
	# Create indexed params
	count = 0
	for (name,value) in SolverStudio.DictionaryDataItems.iteritems():
		 if name in Params:
			 count += 1
			 WriteIndexedParam(g, name,value)
	print count,"items written."

	g.write("end;\n") # Required for GMPL (and optional in AMPL?)

	g.close()

	# Ensure the file "Sheet" that AMPL/GMPL writes its results into is empty
	# (We don't just delete it, as that require import.os to use os.remove)
	resultsfilename = "Sheet"
	g = open(resultsfilename,"w")
	g.close()

	print "## Running "+language+"..."
	print "## "+language+" model file:",SolverStudio.ModelFileName
	print "## "+language+" data file:",dataFilename+"\n"

	sOptions=""
	r = None
	try:
		 r = Application.Range("GMPLOptions")
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
							  key = str(key).strip().lower()  # convert all keys to lowercase 
							  if key=="": continue
							  if len(key)==1: h=1
							  elif key[0]=="-": 
									if key[1]!="-": h=1
									else: h=0
							  else: h=2
							  value =  r.Cells(row+1,2).Value2
						
							  #try: value, rest = (str(r.Cells(row+1,2).Value2)).split(".",1)
							  #except:  value =  str(r.Cells(row+1,2).Value2)
						 
							  if value!=None and type(value) == float: 
									if value >= 1:
										 if value % int(value) == 0:
											  value = int(value)
									value = str(value).strip()
							  if value==None or value=="" or value=="None":
									sOptions = sOptions + hList[h] + key.strip() + " "
							  else:
									sOptions = sOptions + hList[h] + key.strip() + " " + str(value) + " "
						 except Exception as inst:
							  raise Exception( "##Error processing GMPLOptions '"+key+"': "+str(inst))
			  except:
					raise Exception("## There was an unknown error with GMPLOptions.")
		 else:
			  raise Exception("## GMPLOptions must be a table with two columns. Each row should contain an option such as 'timlm' in column 1, and any associated value such as '60' in column 2.")
	print "## GMPLOptions: "+sOptions


	if language=="AMPL":
		# Run AMPL, ensuring we add the AMPL directory to the path which allows the student version to find the solver such as CPLEX in this folder
		exitCode = SolverStudio.RunExecutable(SolverStudio.GetAMPLPath(),'"'+SolverStudio.ModelFileName+'"', addExeDirectoryToPath=True)
	else:
		exitCode = SolverStudio.RunExecutable(SolverStudio.GetGLPSolPath(),"--model \""+SolverStudio.ModelFileName+"\""+" --data \""+dataFilename+"\""+" "+sOptions)
		
	print

	if exitCode == 0:
		print "## "+language+" run completed."
		g = open(resultsfilename,"r")

		# Read a file such as AMPL might produce:
		#flow :=
		#A 1   0
		#B 5   0
		#;
		#age = 22
		# Amount [*] :=     <<<<< Note the [*] which we work around
		#    BEEF  60
		# CHICKEN   0
		#     GEL  40
		#  MUTTON   0
		#    RICE   0
		#   WHEAT   0
		# ;
		# TODO: We have switched to csv.reader; we can improve the code as a result
		lineCount = 0
		itemsLoaded = ""
		itemsChanged = False
		while True :
			line = g.readline()
			if not line: break
			lineCount += 1
			line = line.strip()
			if line=="": continue # skip blank lines
			# Parse the line into its space delimited components, recognising quoted items containing spaces
			items = csv.reader([line], delimiter=' ', skipinitialspace=True, quotechar="'").next()
			if len(items)==0: continue
			name = items[0]
			# print "'"+name+"'="
			if name=="" or name==":":
					raise Exception("An unexpected entry has occured when writing the new values to the sheet; the data item name is missing in the temporary file.\n\nPlease check the SolverStudio examples to see how results are written back to the spreadsheet.\n\n(The error occurred in the string '"+repr(line)+"' on line "+repr(lineCount)+" of file "+repr(resultsfilename)+".)")
			try:      
				dataItem = SolverStudio.DataItems[name]
			except:
				raise Exception("A value was written to the sheet for data item '"+repr(name)+"', but this data item does not exist.\n\n(The error occurred in the string "+repr(line)+" on line "+repr(lineCount)+" of file "+repr(resultsfilename)+".)")
			changed = False
			if len(items)==3 and items[1]=="=":
				# Read a line like: Age = 22 or 'Boys Age' = 22
				try:
					value = float(items[-1])
				except:
					raise Exception("When processing the value written to the sheet for data item '"+repr(name)+"', an error occurred when converting the last item '"+repr(items[-1])+"' to a number.\n\n(The error occurred in string '"+repr(line)+"' in line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")
				try:
					if dataItem!=value:
						globals()[name]=value # Update the python variable of this name (without creating a new variable, which would not be seen by SolverStudio)
						changed = True
				except:
					raise Exception("When writing new values for data item "+repr(name)+" to the sheet, an error occurred when assigning the value "+repr(items[-1])+" to "+repr(name)+".\n\n(The error occurred in string '"+repr(line)+"' in line "+repr(lineCount)+" of file "+repr(resultsfilename)+".)")
			else:
				# Read lines such as: 
				#flow :=    **OR, if we have one index**   flow [*] :=
				#'A B' 1   0
				#B 5   0
				#;
				# For GMPL we do not check for only a :=, but also allow =
				if items[-1] != ":=" and items[-1] != "=":
					raise Exception("When writing new values for data item "+repr(name)+" to the sheet, an error occurred in line "+repr(line)+"; '"+repr(items[-1])+"' was expected to be ':='. (Line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")
				while True:
					line = g.readline()
					lineCount += 1
					if not line: 
					  raise Exception("When writing new values for data item "+repr(name)+" to the sheet, an unexpected end of file occurred. (Line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")
					line = line.replace("''","'")
					line = line.replace(" ';",';')
					items = csv.reader([line], delimiter=' ', skipinitialspace=True, quotechar="'").next()
					if len(items)==0: continue
					if items[0]==";": break
					if len(items)<2:
						raise Exception("When writing new values for data item "+repr(name)+" to the sheet, the line "+repr(line)+" did not contain an index and a value. (Line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")
					try:
						value = float(items[-1])
					except:
						raise Exception("When writing new values for data item "+repr(name)+" to the sheet, an error occurred in line "+repr(line)+" when converting the last item '"+repr(items[-1])+"' to a number. (Line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")
					index = items[0:-1]
					# Convert any values into floats
					for j in range(len(index)):
						try:
							index[j]=float(index[j])
						except ValueError:
							pass
					#print tuple(index),value
					#try:
					if len(index)==1:
						index2 = index[0]  # The index is a single value, not a tuple
					else:
						index2 = tuple(index)# The index is a tuple of two or more items
					if dataItem[index2]!=value:
						dataItem[index2]=value
						changed = True
					#except:
					#   raise Exception("When writing new values for data item "+repr(name)+" to the sheet, an error occurred in line "+repr(line)+" when assigning the value "+repr(items[-1])+" to index "+repr(index2)+". (Line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")            
			if changed:
				itemsLoaded = itemsLoaded+name+"* "
				itemsChanged = True
			else:
				itemsLoaded = itemsLoaded+name+" "
		g.close()
		if itemsLoaded=="":
			print "## No results were loaded into the sheet."
		elif not itemsChanged:
			print "## Results loaded for data items:", itemsLoaded
		else :
			print "## Results loaded for data items:", itemsLoaded
			print "##   (*=data item values changed on sheet)"
	else:
		print "## "+language+" did not complete (Error Code %d); no solution is available. The sheet has not been changed." % exitCode
	print "## Done"

DoRun()
