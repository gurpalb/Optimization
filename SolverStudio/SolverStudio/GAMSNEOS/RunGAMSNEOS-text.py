# This unused file implements GAMS data transfer using text files instead of using
# GDX files. It does not require the propietary GAMS dll's

import csv
import os
import sys
import xmlrpclib
import time
import re
import base64
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import  Form, Label, ToolStripMenuItem, MessageBox, MessageBoxButtons, Button, TextBox, FormBorderStyle, DialogResult, OpenFileDialog #Application,
from System.Drawing import Point

# Return s as a string suitable for indexing, which means tuples are listed items in []
def WriteIndex(f,s):
   if type(s) is tuple:
      k = 0
      for i in s:
         k=k+1
         if isinstance(i,float):
             # We cannot have doubles as indices as we cannot have a '.'
             f.write( repr(int(i+0.5)) )
         elif isinstance(i,int):
             f.write( repr(i) )
         else:
             f.write( repr(i) )
         if k<>len(s):
             f.write( "." )
         #f.write('['+','.join([repr(t) for t in s])+']')
         #f.write(repr(s)[1:-1].replace(",","."))
   else:
      if isinstance(s,float):
          # We cannot have doubles as indices as we cannot have a '.'
          f.write( repr(int(s+0.5)) )
      elif isinstance(s,int):
          f.write( repr(s) )
      else:
          f.write( repr(s) )

# Return s as a string suitable for an item in a set, which means tuples are listed items in ()
def WriteItem(f,s):
   if type(s) is tuple:
      #f.write( '('+','.join([repr(t) for t in s])+')' )
      f.write( repr(s).replace(",",".") )
   else:
      if isinstance(s,float):
          # We cannot have doubles as indices as we cannot have a '.'
          f.write( repr(int(s+0.5)) )
      elif isinstance(s,int):
          f.write( repr(s) )
      else:
          f.write( repr(s) )

def WriteParam(f,name,value):
   f.write("param "+name+" := "+repr(value)+";\n\n")

# set NUTR := A B1 B2 C ;
def WriteSet(f,name,value):
   f.write("Set "+name+" / \n")
   k = 0
   for i in value:
      k=k+1
      f.write(" ")
      WriteItem(f,i)
      if k <> len(value):
         f.write(",\n")
   f.write(" / ;\n\n")

# param demand := 
#   2 6 
#   3 7 
#   4 4 
#   7 8; 
# param sc  default 99.99 (tr) :

def WriteIndexedParam(f,name,value):
   f.write("Parameter "+name+" /\n")
   for (index,val) in value.iteritems():
      f.write(" ")
      WriteIndex(f,index)
      f.write(" ")
      WriteItem(f,val)
      f.write("\n")
   f.write(" / ;\n\n")

def WriteTable(f,name,value):
   f.write("Parameter "+name+" /\n")
   for (index,val) in value.iteritems():
      f.write(" ")
      WriteIndex(f,index)
      f.write(" ")
      WriteItem(f,val)
      f.write("\n")
   f.write(" / ;\n\n")

# Read lines like:
# Sets
# i  canning plants
# j  markets / a b / ;
# ---- and ---- #
# Set i  canning plants / Seattle, San-Diego / ;

def ScanModelFileForDataItems(modelFileName):
   unusedKeys = ["variables","positive","equations","model","solve","scalar","display"]
   Sets= list()
   Params = list()
   Tables = list()
   #singlelineKeyWords = { "set" : Sets, 'parameter' : Params, 'table' : Tables }
   keyWords = { "sets" : Sets, 'parameters' : Params, 'tables' : Params, "set" : Sets, 'parameter' : Params, 'table' : Params }
   lineCount = 0
   f = open(SolverStudio.ModelFileName,"r")
   items = None
   for line in f:
      line = line.strip() # Remove white space, including the newline
      # Replace anything that can follow the name with a 'space' so the string splits correctly.
      l = line.replace("("," ").replace(";"," ; ") #.replace("<"," ").replace(">"," ").replace("{"," ").replace(":"," ").replace(","," ")
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
            items.append(tokens[0])
      #The command ends with a ; so search for this
      if ";" in tokens:
          items = None
		
   f.close()
   return Sets, Params
	
def parseNEOSQueue(queue, jobNumber):
   lines = queue.splitlines()
   running = False
   queued = False
   numRunning = 0
   numQueued = 0
   position = -1
   for line2 in lines:
      line = line2.strip()
      if line == "":
         pass
      elif line == "Running:":
         running = True
         queued = False
      elif line == "Queued:":
         queued = True
         running = False
      elif line[0:4] != "job#":
          pass
      elif running:
          numRunning += 1
      elif queued:
          numQueued += 1
          items = line.split(None,2) # split up job# 311706 nco LOQO AMPL submitted 
          #print items[1]
          try:
             job = int(items[1])
             if job == jobNumber:
                position = numQueued
          except:
             pass
          
   return numRunning, numQueued, position

def BuildXML(model,dataFileName):
   with open(dataFileName,'r') as content_file:
      data=content_file.read()
   #remove $include "SheetData.dat"
   r=re.compile('^\$include\s*"SheetData.dat"\s*;',re.IGNORECASE | re.MULTILINE)
   model = r.sub(data,model)
   #remove GAMS local output calls
   model = model[0:model.index(";",model.index("solve"))]+"\n**SolverStudio: Line(s) modified"
   if len(model) <10000:   
	   typeStart="<![CDATA["
	   typeEnd="]]>"
   else: #compress large files
      with open(SolverStudio.ModelFileName,"w") as temp:
         temp.write(model)
      model = encodeGzip(SolverStudio.ModelFileName,"model.gms.gz")
      typeStart = "<base64>"
      typeEnd = "</base64>"
   oldDirectory = SolverStudio.ChangeToLanguageDirectory()
   with open('NEOSTemplate.xml', 'r') as content_file:
       xml= content_file.read()
   SolverStudio.SetCurrentDirectory(oldDirectory)
   if SolverStudio.ActiveModelSettings.GetValueOrDefault("QueueLength",False):
		priority="short"
   else:
      priority="long"
   solver =  SolverStudio.ActiveModelSettings.GetValueOrDefault("GAMSNeosSolver",'Cbc')
   category =  SolverStudio.ActiveModelSettings.GetValueOrDefault("GAMSNeosCategory",'milp')
   xml=xml%(category,solver,priority,typeStart,model,typeEnd,"")
   with open('NEOSJob.xml','w') as content_file:
       content_file.write(xml)
   return
   
def BuildXMLGDX1(model):  
   if os.path.exists('GAMSdata.gdx'):
      os.remove('GAMSdata.gdx')
   listItems=Sets+Params
   strList=""
   for i in listItems:
		strList = strList + i + ","
   strList=strList[0:-1]
   model = model[0:model.index(";",model.index("solve"))]+";"  
   r=re.compile('solve')
   model=r.sub('execute_unload "GAMSdata.gdx" '+strList+';\n*SOLVE',model)
   with open(SolverStudio.ModelFileName,"w") as dest_file:
      dest_file.write(model)         
   exitCode = SolverStudio.RunExecutable(SolverStudio.GetGAMSPath()," \""+SolverStudio.ModelFileName+"\"")
   if exitCode != 0: raise Exception("Error running GAMS: Exit Code %d" % exitCode)
   r=re.compile('\*SOLVE')
   model=r.sub('solve',model)
   r=re.compile('^\s*\$\s*include\s*"SheetData.dat"\s*;',re.IGNORECASE | re.MULTILINE)
   model=r.sub("\n$GDXIN in.gdx \n$LOAD " + strList ,model)
   r=re.compile('^execute_unload.*',re.IGNORECASE | re.MULTILINE)
   model=r.sub("",model)
   with open(SolverStudio.ModelFileName,"w") as dest_file:
      dest_file.write(model)
   oldDirectory = SolverStudio.ChangeToLanguageDirectory()
   with open('NEOSTemplate.xml', 'r') as content_file:
     xml= content_file.read()
   SolverStudio.SetCurrentDirectory(oldDirectory)
   # with open('GAMSdata.gdx',"r") as data_file:
      # data = base64.encodestring(data_file.read());
   data = encodeGzip('GAMSdata.gdx','GAMSdata.gdx.gz')
   if SolverStudio.ActiveModelSettings.GetValueOrDefault("ShortQueueLength",False):
		priority="short"
   else:
      priority="long"
   solver =  SolverStudio.ActiveModelSettings.GetValueOrDefault("NeosSolver",'Cbc')
   category =  SolverStudio.ActiveModelSettings.GetValueOrDefault("NeosCategory",'milp')
   xml=xml%(category,solver,priority,"<![CDATA[",model,"]]>",data)
   with open('NEOSJob.xml','w') as content_file:
      content_file.write(xml)
   return
   
def encodeGzip(in_file,out_gz):
	import gzip
	with open(in_file,'rb') as orig_file:
		with gzip.open(out_gz,'wb') as zipped_file:
			zipped_file.writelines(orig_file)	
	with open(out_gz,'rb') as zipped_file:
		encodegzip = base64.encodestring(zipped_file.read());
	return encodegzip	
	
def runGAMSonNEOS(resultsfilename):
	import urllib2
	f=open("NEOSJob.xml")
	xml=f.read()
	NEOS_HOST="neos-server.org"
	NEOS_PORT=3332
	neos=xmlrpclib.ServerProxy("http://%s:%d" % (NEOS_HOST, NEOS_PORT)) # AJM: .server is deprecated
	queue = neos.printQueue()
	numRunning, numQueued, position = parseNEOSQueue(queue, -1)
	print "NEOS Status: %d running, %d queued jobs" % (numRunning, numQueued)
	jobNumber=0
	status = "NotStarted"
	completed = False
	try:
		(jobNumber,password) = neos.submitJob(xml)
		if (jobNumber == 0):
			sys.stdout.write("NEOS job submission failed\nError=%s\n" % password)
			return False
		sys.stdout.write("Submitted JobNumber=%d, password=%s\n" % (jobNumber,password))
		status="Waiting for NEOS status..."
		status = neos.getJobStatus(jobNumber, password)
		if status=="Waiting":
			print "NEOS Status: Waiting in queue..."
			counter = 0
			while status=="Waiting":
				time.sleep(1)
				counter -= 1
				status = neos.getJobStatus(jobNumber, password)
				posn = -2
				numRunning = -2
				currentJob = ""
				if status == "Waiting" and counter<=0:
					counter = 1
					queue = neos.printQueue()
					newNumRunning, numQueued, newPosn = parseNEOSQueue(queue, jobNumber)
					if posn != newPosn or numRunning != newNumRunning:
						posn = newPosn
						numRunning = newNumRunning
						print "Waiting:%3d/%2d in queue (%2d jobs running)" % (posn, numQueued, numRunning)
		if status == "Running":
			print "NEOS Status: Running..."
			offset=0
			while status == "Running":
				time.sleep(1)
				(msg,offset) = neos.getIntermediateResults(jobNumber,password,offset)
				sys.stdout.write(msg.data)
				status = neos.getJobStatus(jobNumber, password)
		if status=="Done":
			msg = neos.getFinalResults(jobNumber, password).data
			sys.stdout.write("NEOS Status: Job Done\n")#Job Output:\n")
			sys.stdout.write(msg)
			succeeded=True
			with open(resultsfilename,'w') as content_file:
				content_file.write(msg)
				completed = True
			print "Downloading GDX result file"
			try:
				url = "http://www.neos-server.org/neos/jobs/"+str(jobNumber)+"-"+str(password)+"-out.gdx"
				page = urllib2.urlopen(url)
				gdxData = page.read()
				page.close()
				if os.path.exists('outputGDXdata.gdx'): os.remove('outputGDXdata.gdx')
				with open('outputGDXdata.gdx',"wb") as dest_file:
					dest_file.write(gdxData)
				print "Download complete"
			except:
				print "Unable to download GDX result file"
		else:
			print "Error: NEOS reported %s" % status
	finally:
		if jobNumber != 0 and status!="Done" and status!="NotStarted":
			neos.killJob(jobNumber, password, '')
			print "Killed NEOS job %d" % jobNumber
	return completed

def GetResults(resultsfilename):  
   f=open(resultsfilename,"r")
   g=open("GAMSResults.temp","w")
   doPrint=False
   while True:
      line=f.readline()
      if not line: break
      line2 = line.replace("("," ").replace("'"," ").replace(","," ").replace(")"," ")
      line2 = line2.strip()
      if line2=="": continue # skip blank lines
      items = line2.split()
      if "VAR"in items: doPrint = True
      if doPrint and "****" in items: break
      if doPrint:
         g.write(line)
   f.close()
   g.close()
   return

def FixResults():
   f=open("GAMSResults.temp","r")
   g=open("Sheet","w")
   good=True
   lineCount = 0
   while good:
      line=f.readline()
      lineCount = lineCount+1
      if not line: break
      line2 = line.replace("("," ").replace("'"," ").replace(","," ").replace(")"," ")
      line2 = line2.strip()
      items = line2.split()
      for i in range(0,len(items)):
          if items[i] == "VAR":
              varName=items[i+1]
      line=f.readline()
      lineCount = lineCount+1
      if not line: break
      while True:
         line=f.readline()
         lineCount = lineCount+1
         if not line:
            good=False
            break
         #line2 = line.replace("("," ").replace("'"," ").replace(","," ").replace(")"," ")
         #line2 = line2.strip()
         if line=="": continue # skip blank lines
         items = line.split()
         if items[0]=="LOWER": break
         try:
			   index=line.index("      ",line.index("."))
         except:
			   raise Exception("There was an error reading line " + str(lineCount) + " of the file 'GAMSResults.temp'.")
         if len(items)>5:
            items2=list()
            first=""
            for i in range(0,len(items)-4):
               first=first+items[i]
            items2.append(first)
            items2=items2+items[-4:]
            items = items2
         items[0]=line[0:index]
         i=items[0].index(".")
         items[0]="'"+items[0][0:i].strip().replace("'","''")+"','"+items[0][i+1:].strip().replace("'","''")+"'"
         if items[2]==".":
             items[2]="0.0"
         out=list()
         out.append(varName+"("+items[0]+")")
         out.append(items[2])
         print >> g,'{:<30} {:>10}'.format(out[0],out[1])
         # g.write(varName+"("+items[0]+")"+"       "+items[2]+"\n")
   f.close()
   g.close()
   return

def ImportGDX():
	path = SolverStudio.GetGAMSPath()
	source = SolverStudio.WorkingDirectory() + "\\outputGDXdata.gdx"
	path = path.replace("gams.exe","gdxdump")
	index = len(source)-source[::-1].index("\\")-1
	GDXName = source[index+1:len(source)-4]
	tempTxt=  SolverStudio.WorkingDirectory() + "\\" + GDXName + "_temp.txt"
	path = path.replace("\\","\\\\")
	source = source.replace("\\","\\\\")
	tempTxt = tempTxt.replace("\\","\\\\")
	command = "\"" + path + "\" " + "\"" + source + "\"" + " output = \"" + tempTxt
	os.system(command)
	myVars={}
	csvPath = tempTxt.replace("_temp.txt",".txt")
	f=open(tempTxt,"r")
	g=open(csvPath,"w")
	row=0
	lineCount=0
	start=False
	printLine = False
	ID = ""
	print "## Loading GDX result file"
	f.readline()
	f.readline()
	while True:
		lineCount=lineCount+1
		line = f.readline()
		if not line: break
		if line == "": 
			ID = ""
			printLine = False
		if "$offempty" in line: break
		if "Equation " in line: printLine = False
		if "Variable" in line:
			printLine = True
			startIndex = line.index("Variable")+9
			if "/;"  in line:
				endIndex = line.index(" ",startIndex+1)
				try:
					newIndex = line.index("(",startIndex+1)
					endIndex = min(endIndex,newIndex)
				except: pass
				varName = line[startIndex:endIndex]
				try:
					varValue = line[line.index("/L ")+3:line.index(" /;")]
				except: varValue = ""
				g.write(varName + "\n")
				g.write(varValue + "\n")
				ID="var1"
			else:
				endIndex = line.index("(",startIndex)
				varName = line[startIndex:endIndex]
				g.write(varName+"\n")
				ID = "var2"
			start = True
			startRow = lineCount
			continue
		line = line.replace("/;","").replace(",","")
		line = line.replace("'.'","','")
		if ID=="var2" and ("'.M " in line or "'.UP " in line or "'.LO " in line): 
			lineCount = lineCount - 1
			continue
		if ID =="var2": line=line.replace("'.L ","' ")
		if start and lineCount==startRow+1:
			dim = line.count("','")+1
			if ID == "var2": dim = dim +1   
		if start and line=="\n":
			if ID == "var2":
				myVars[varName]=[lineCount-startRow-1,dim]
			if ID == "var1":
				myVars[varName]=[lineCount-startRow-1,0]
			start=False
		if printLine: g.write(line)
		if lineCount%100000 == 0: print "##Reached line " + str( lineCount)
	f.close()
	g.close()
	print "## Found " + str(len(myVars)) + " variables"
	print
	
def ConvertGDXResults(resultsfilename):
	f=open('outputGDXdata.txt',"r")
	g=open(resultsfilename,"w")
	start = False
	while True:
		line = f.readline()
		if not line: break
		line=line.strip()
		name = line
		try:
			dataItem = SolverStudio.DataItems[name]
			if dataItem == None:
				value = f.readline().strip()
				if value != "": g.write(name + " " + value+ "\n")
				continue
			for i in dataItem:
				dataItem[i]=0
			start = True
		except: start = False
		while start:
			line = f.readline()
			if line == "\n": 
				g.write("\n")
				start = False
				break
			line = line.strip()+"\n"
			if not "'" in line:
				g.write(name + "  " + line)
				continue
			newLine = name + "(" + line.replace("' ","') ")
			g.write(newLine)
	f.close()
	g.close()
		
# Ensure the file "Sheet" that AMPL writes its results into is empty
# (We don't just delete it, as that require import.os to use os.remove)
resultsfilename = "Sheet"
g = open(resultsfilename,"w")
g.close()

# Create a 'data' file, that contains the appropriate data
dataFilename = "SheetData.dat"
g = open(dataFilename,"w")
g.write("* GAMS data file created by SolverStudio\n\n")

print "## Scanning model for sets and parameters"
Sets, Params = ScanModelFileForDataItems(SolverStudio.ModelFileName)
#print "Sets=",Sets
#print "Params=",Params

print "## Building GAMS input file for %s sets and %s parameters"%(len(Sets),len(Params))

print "## Writing simple parameters...",
# Create constants
count = 0
for (name,value) in SolverStudio.SingletonDataItems.iteritems():
	 if name in Params:
		 count += 1
		 WriteParam(g, name,value)
print count,"items written."
		 
print "## Writing sets...",
# Create sets
count = 0
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

with open(SolverStudio.ModelFileName,"r") as content_file:
	model = content_file.read()

try:
	SolverStudio.GetGAMSPath()
	BuildXMLGDX1(model)
	print "## Sent with data as GDX file"
except:
	BuildXML(model,dataFilename)
	print "## Sent with data in the model file"
# BuildXML(model,dataFilename)
# print "##Sent with data in the model file"
	
	
print "\n## Running GAMS on NEOS"
print "## GAMS model file:",SolverStudio.ModelFileName	
print "## GAMS data file:",dataFilename + "\n"
succeeded = runGAMSonNEOS(resultsfilename)

if succeeded:
	print "\n## GAMS completed successfully.\n## Reading results..."
	try:
		SolverStudio.GetGAMSPath()
		BookName = Application.ActiveWorkbook.Name
		SheetName = Application.ActiveSheet.Name
		ImportGDX()
		Application.Workbooks(BookName).Worksheets(SheetName).Select()
		ConvertGDXResults(resultsfilename)
		print "## Results obtained from GDX result file"
	except:
		GetResults(resultsfilename)
		FixResults()
		print "## Results obtained from NEOS text file"
	# GetResults(resultsfilename)
	# FixResults()
	# print "## Results obtained from NEOS text file"
	g = open(resultsfilename,"r")
	lineCount = 0
	# Read a file such as: 
	# flow('A','1')      500.00
	itemsLoaded = list()
	while True :
		line = g.readline()
		if not line: break
		lineCount += 1
		line = line.replace("("," ").replace("'","'").replace(","," ").replace(")"," ")
		items = csv.reader([line], delimiter=' ', skipinitialspace=True, quotechar="'").next()
		#print items
		if len(items)==0: continue
		if len(items) == 2:
			name = items[0]
			value = items[1]
			globals()[name]=value
			if not name in itemsLoaded: itemsLoaded.append(name)
			continue
		# get name and index
		name = items[0]
		index = items[1:-1]
		# convert any values into floats that we can
		for i in range(len(index)):
			try:
				index[i] = float(index[i])
			except:
				pass
		# get value
		try:
			value = float(items[-1])
		except:
			raise TypeError("When reading the results for data item '"+repr(name)+"', an error occurred in line "+repr(line)+" when converting the last item '"+repr(items[-1])+"' to a number. (Line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")
		# print name,index,value
		dataItem = SolverStudio.DataItems[name]
		try:
			if dataItem[tuple(index)]<> value:
				dataItem[tuple(index)]=value
		except:
			raise TypeError("When reading the results for data item "+repr(name)+", an error occurred in line "+repr(line)+" when assigning the value "+repr(items[-1])+" to index "+repr(tuple(index))+". (Line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")            
		if not name in itemsLoaded:
			itemsLoaded.append(name)
	if len(itemsLoaded)==0:
		print "## No results were loaded into the sheet."
	else:
		print "## Results loaded for data items:", itemsLoaded
else:
   print "## GAMS did not complete; no solution is available."
print "## Done"