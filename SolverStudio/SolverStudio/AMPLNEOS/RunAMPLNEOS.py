# Updated 2012.07.20 to interpret quoted indices with spaces (by using csv.reader)
# Updated 2012.07.20 to interpret quoted indices with spaces (by using csv.reader)
# Updated 20120814 to solve both GMPL and AMPL models
# This allows us to parse AMPL model files to distinguish between 
# singletons that can be params and single-member sets 
#
# v0.21, 2013.06.03:  Close the 'sheet' file
# 20140407: Added code to replace all pairs of LF LF now appearing in NEOS output with a single LF; a bug report has been filed with NEOS.
# 20140620 Removed any :g15 formatting as this is locale specific, and so fails in non-English countries
# 20141104: pATCHED TO ALLOW TIMESOUTS in NEOS comms
# 20141105 Convert any "]]>" into "]] >" in the model so it can be embedded in a CData object (for which ]]> is special)

language = "AMPL"
isAMPL = language == "AMPL"
useNEOS = True

from _winreg import *

import csv

if useNEOS:
	import sys
	import xmlrpclib
	import time
	import os
	import re
	import socket

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
		#f.write('['+','.join([fmt(t) for t in s])+']')
		return fmtTuple(s)
	else:
		return fmtSngl(s)

def WriteIndex(f,s):
	if type(s) is tuple:
		#f.write('['+','.join([fmt(t) for t in s])+']')
		f.write('['+fmt(s)+']')
	else:
		f.write(fmt(s))

def WriteItem(f,s):
	if type(s) is tuple:
		f.write( "("+fmt(s)+")" )
	else:
		f.write(fmt(s))

def WriteParam(f,name,value):
	f.write("param "+name+" := "+fmt(value)+";\n\n")

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
	solveCommand=False
	dataCommand=False
	displayCommand = False
	solver=None;
	Variables=list();
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
		elif tokens[0].lower()=="solve":
			solveCommand=True
		elif tokens[0].lower()=="data":
			dataCommand=True
		elif tokens[0].lower()=="display":
			displayCommand=True
		elif tokens[0].lower() == "var":
			Variables.append(tokens[1])
		elif tokens[0].lower() == "option":
			if tokens[1].lower() == "solver":
				solver = tokens[2]
	f.close()
	return Sets, Params,solveCommand,dataCommand,displayCommand,Variables,solver
	
def FixBadCommands(model,solveCommand,dataCommand,displayCommand,Variables):
	if not displayCommand:
		if solveCommand and not dataCommand:
			r= re.compile(r'^solve\s*;',re.IGNORECASE | re.MULTILINE)
			model = r.sub("data SheetData.dat;",model)
			model = model + "\n\nsolve; #SolverStudio: Line modified\n\n"
		if not solveCommand and dataCommand:
			model = model + "\nsolve; #SolverStudio: Line modified\n\n" 
		if not solveCommand and not dataCommand:
			model = model + "\n\ndata SheetData.dat; #SolverStudio: Line modified" + "\n\nsolve; #SolverStudio: Line modified\n\n"
		for i in Variables:
			model = model + "display " + i + " > Sheet; #SolverStudio: Line modified\n\n"
	else:
		if solveCommand and not dataCommand:
			r= re.compile(r'^solve\s*;',re.IGNORECASE | re.MULTILINE)
			model = r.sub("data SheetData.dat; #SolverStudio: Line modified\n\nsolve;",model)
		if not solveCommand and dataCommand:
			r=re.compile(r'^display\s',re.IGNORECASE| re.MULTILINE)
			model = r.sub("solve; #SolverStudio: Line modified\n\ndisplay",model,1)
		if not solveCommand and not dataCommand:
			r=re.compile(r'^display\s',re.IGNORECASE| re.MULTILINE)
			model = r.sub("\ndata SheetData.dat; #SolverStudio: Line modified\n\nsolve; #SolverStudio: Line modified\n\ndisplay ",model,1)      
	return model

	
def DoSmallFixes(solveCommand,dataCommand,displayCommand,Variables,solver):
	with open(SolverStudio.ModelFileName,"r") as input_file:
		model=input_file.read()
	#Fix missing data,solve,display commands
	if not (displayCommand and solveCommand and dataCommand):
		print "## Fixing model commands"
		model=FixBadCommands(model,solveCommand,dataCommand,displayCommand,Variables)

	with open(SolverStudio.ModelFileName,"w") as dest_file:
		dest_file.write(model)
	return

def ConvertModelForNEOS(model):
	r = re.compile(r'\bSheetData.dat\b',re.IGNORECASE)
	model = r.sub('ampl.dat',model)
	solver=None
	r = re.compile(r"^\s*option\s+solver\s+(\w+)\s*;",re.IGNORECASE | re.MULTILINE)
	m=r.findall(model)
	if m:
		solver = m[-1] # the last solver command
		model=r.sub('#SolverStudio: solver command deleted for NEOS',model) # delete the line(s)
	r = re.compile(r'^\s*display\s+([\.\w]+)\s*>\s*sheet\s*;',re.IGNORECASE | re.MULTILINE)
	model=r.sub(r'_display \1; #SolverStudio: output modified for NEOS',model)
	model = model + "\n\nend;\n"
	return (solver,model)

# Replace any "]]>" with "]] >" because "]]>" is the end-of-CData delimiter   
def MakeModelSafeForCData(model):
	r = re.compile(r']]>',re.IGNORECASE)
	model = r.sub(']] >',model)
	return model

def GetNEOSCategory(solver):
	solverList = {
		'asa':'go',   # ASA
		'bonmin':'minco',   # Bonmin
		'bpmpd':'lp',   # bpmpd
		'cbc':'milp',   # Cbc
		'condor':'ndo',   # condor
		'conopt':'nco',   # CONOPT
		'cplex':'milp',   # Cplex added 2015
		'couenne':'minco',   # Couenne
		'feaspump':'milp',   # feaspump
		'filmint':'minco',   # FilMINT
		'filter':'nco',   # filter
		'gurobi':'lp',   # Gurobi
		'gurobi':'milp',   # Gurobi
		'icos':'go',   # icos
		'ipopt':'nco',   # Ipopt
		'knitro':'cp',   # KNITRO
		# 'knitro':'nco',   # KNITRO
		'lancelot':'nco',   # LANCELOT
		'l-bfgs-b':'bco',   # L-BFGS-B
		'loqo':'nco',   # LOQO
		'minlp':'minco',   # MINLP
		'minos':'nco',   # MINOS
		'minto':'milp',   # MINTO
		'mosek':'lno',   # MOSEK
		#'mosek':'lp',   # MOSEK
		#'mosek':'nco',   # MOSEK
		'muscod-ii':'miocp',   # MUSCOD-II
		'nsips':'sio',   # nsips
		'ooqp':'lp',   # OOQP
		'path':'cp',   # PATH
		'pgapack':'go',   # PGAPack
		'pswarm':'go',   # PSwarm
		'qsopt_ex':'milp',   # qsopt_ex
		'scip':'go',   # scip
		'scip':'milp',   # scip
		'scip':'minco',   # scip
		'snopt':'nco',   # SNOPT
		# 'xpressmp':'lp',   # XpressMP
		'xpressmp':'milp'  # XpressMP
		}
	return solverList.get(solver.lower(),None)
	
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
				job = str(items[0])
				if job.find(str(jobNumber)) != -1:
					position = numQueued
			except:
				pass
			
	return numRunning, numQueued, position
	
def encodeGzip(in_file):
	import gzip
	import base64
	out_gz="SheetData.dat.gz"
	with open(in_file,'rb') as orig_file:
		with gzip.open(out_gz,'wb') as zipped_file:
			zipped_file.writelines(orig_file)   
	with open(out_gz,'rb') as zipped_file:
		encodegzip = base64.encodestring(zipped_file.read());
	return encodegzip

def NeosWarningEnabled():
	try:
		keyVal = 'Software\\SolverStudio\\NEOS'
		key = OpenKey(HKEY_CURRENT_USER, keyVal, 0, KEY_ALL_ACCESS)
		result = QueryValueEx(key,"ShowNeosWarning")
		CloseKey(key)
		return result[0] != "False"
	except: 
		return True  

def ConfirmNEOS():
	import clr
	clr.AddReference('System.Windows.Forms')
	clr.AddReference('System.Drawing')
	from System.Drawing import Point,Color
	from System.Windows.Forms import LinkBehavior,LinkArea,LinkLabel,CheckBox,FormBorderStyle,Form, Label,Button, ToolStripMenuItem, MessageBox, MessageBoxButtons, DialogResult, FormStartPosition

	class WarningForm(Form):
		def __init__(self):
			self.Text = "SolverStudio: NEOS Warning"
			self.FormBorderStyle = FormBorderStyle.FixedDialog    
			self.Height=235
			self.Width = 310
			# self.Warning = Label()
			# self.Warning.Text = "Models solved using via NEOS will be available in the public domain for perpetuity."
			# self.Warning.Location = Point(10,10)
			# self.Warning.Height  = 30
			# self.Warning.Width = 380

			self.myOK=Button()
			self.myOK.Text = "Continue"
			self.myOK.Width = 100
			self.myOK.Location = Point(self.Width/2 - self.myOK.Width - 30, self.Height - 65)
			self.myOK.Click += self.AddChoice

			self.myCancel=Button()
			self.myCancel.Text = "Cancel"
			self.myCancel.Width = 100
			self.myCancel.Location = Point(self.Width/2 + 30, self.Height - 65)

			self.check1 = CheckBox()
			self.check1.Text = "Never show NEOS warnings again."
			self.check1.Location = Point(12, 80)
			self.check1.Width = 250
			self.check1.Height = 130
			
			self.link1 = LinkLabel()
			self.link1.Location = Point(10, 10)
			self.link1.Width = self.Width-15
			self.link1.Height = 120
			self.link1.LinkClicked += self.OpenTCs
			self.link1.VisitedLinkColor = Color.Blue;
			self.link1.LinkBehavior = LinkBehavior.HoverUnderline;
			self.link1.LinkColor = Color.Navy;
			msg = "By using NEOS via SolverStudio you agree to the NEOS-Server terms and conditions, and accept that your model will become publicly available.\n\n"
			msg += "Some solvers require an email address, which you can set using the AMPL menu. "
			email = SolverStudio.GetRegistrySetting("NEOS","Email","")
			if email=="":
				msg += "You have not provided an email address, and so some solvers may fail." 
			else:
				msg += "Your email address ("+email+") will be sent to NEOS." 
			# msg += "\n\nContinue?"
			self.link1.Text = msg
			self.link1.LinkArea = LinkArea(60,20)
			
			self.AcceptButton = self.myOK
			self.CancelButton = self.myCancel

			self.Controls.Add(self.myOK)      
			self.Controls.Add(self.myCancel)      
			self.Controls.Add(self.link1)      
			# self.Controls.Add(self.Warning)   
			self.Controls.Add(self.check1)
			# self.CenterToScreen()
			self.StartPosition = FormStartPosition.CenterParent;

		def AddChoice(self,sender,event):
			keyVal = 'Software\\SolverStudio\\NEOS'
			try:
				key = OpenKey(HKEY_CURRENT_USER, keyVal, 0, KEY_ALL_ACCESS)
			except:
				key = CreateKey(HKEY_CURRENT_USER, keyVal)
			SetValueEx(key, "ShowNeosWarning", 0, REG_SZ, str(not self.check1.Checked))
			CloseKey(key)
			self.DialogResult = DialogResult.OK
			self.Close()
		def OpenTCs(self,sender,event):
			import webbrowser
			webbrowser.open("""http://www.neos-server.org/neos/termofuse.html""")
	
	# Actual code for ConfirmNEOS
	myForm = WarningForm()
	# result = myForm.ShowDialog()
	result = SolverStudio.ShowDialogWithExcelAsParent(myForm)
	return result == DialogResult.OK

def runAMPLonNEOS(modelFileName, dataFileName, resultsfilename):
	import clr # Need this to define System.IO.IOException exception
	clr.AddReference('System')  # Need this to define System.IO.IOException exception; note that System.IO is defined in 
										# assembly mscorlib, not System.IO, so clr.AddReference('System.IO') is wrong as there is no 
										# such assembly. This fails for others, even tho' it works ok for me
	import System
	oldDirectory = SolverStudio.ChangeToLanguageDirectory()
	with open('NEOSTemplate.xml', 'r') as content_file:
		xml= content_file.read()
	with open('NEOS.run', 'r') as content_file:
		commands= content_file.read()
	SolverStudio.SetCurrentDirectory(oldDirectory)
	with open(modelFileName, 'r') as content_file:
		model= content_file.read()   
	data_size = os.path.getsize(dataFileName)
	if data_size <10000:   
		with open(dataFileName, 'r') as content_file:
			data= content_file.read()
			typeStart="<![CDATA["
			typeEnd="]]>"
	else: #compress large files
		typeStart = "<base64>"
		typeEnd = "</base64>"
		data= encodeGzip(dataFileName)
	inputtype='AMPL'
	
	(solver, model) = ConvertModelForNEOS(model)
	model = MakeModelSafeForCData(model)
	if solver == None:
		# This will use the solver and category in the drop down menu
		solver = SolverStudio.ActiveModelSettings.GetValueOrDefault("AMPLNeosSolver","Cbc")
		category = SolverStudio.ActiveModelSettings.GetValueOrDefault("AMPLNeosCategory","milp")
	else:
		# This will use the Solver from the model file, and match it to a category in GetNEOSCategory
		category = GetNEOSCategory(solver)
	
	if SolverStudio.ActiveModelSettings.GetValueOrDefault("ShortQueueLength",False):
		priority="short"
	else:
		priority="long"
	email = SolverStudio.GetRegistrySetting("NEOS","Email","") # TODO: Encode this to be safe, eg urllib.urlencode?
	xml=xml%(category, solver, inputtype, priority, email, model, typeStart, data, typeEnd, commands)
	with open('ampl.mod','w') as content_file:
		content_file.write(model)
	with open('NEOSJob.xml','w') as content_file:
		content_file.write(xml)
	NEOS_HOST="neos-server.org"
	NEOS_PORT=3332
	try:
		neos=xmlrpclib.ServerProxy("http://%s:%d" % (NEOS_HOST, NEOS_PORT)) # AJM: .server is deprecated
		queue = neos.printQueue()
	except:
		print "## Error: An error occurred connecting to NEOS, perhaps because there is no internet connection or NEOS is currently offline."
		raise Exception("NEOS Connection Error: An error occurred connecting to NEOS, perhaps because there is no internet connection or NEOS is currently offline. No results are available.")
	numRunning, numQueued, position = parseNEOSQueue(queue, -1)
	print "## NEOS Status: %d running, %d queued jobs" % (numRunning, numQueued)
	jobNumber=0
	status = "NotStarted"
	completed = False
	# Change to a 60s timeout (by default, there is no timeout)
	oldTimeout = socket.getdefaulttimeout()
	socket.setdefaulttimeout(60)
	showLog = False
	try:
		(jobNumber,password) = neos.submitJob(xml)
		if (jobNumber == 0):
			sys.stdout.write("## NEOS job submission failed\nError=%s\n" % password)
			return False
		sys.stdout.write("## Submitted JobNumber=%d, password=%s\n" % (jobNumber,password))

		status="Waiting for NEOS status..."
		status = neos.getJobStatus(jobNumber, password)
		if showLog: print "::::  status = ",status
		if status=="Waiting":
			print "## NEOS Status: Waiting in queue..."
			counter = 0
			while status=="Waiting":
				time.sleep(5) # Changed from 1 to 5 to avoid NEOs rejecting too many requests (20150520)
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
						print "## Waiting:%3d/%2d in queue (%2d jobs running)" % (posn, numQueued, numRunning)
		if status == "Running":
			print "## NEOS Status: Running..."
			offset=0
			while status == "Running":
				time.sleep(5) #  # Changed from 1 to 5 to avoid NEOs rejecting too many requests (20150520)
				try:
					if showLog: print ":::: calling getIntermediateResults(",jobNumber,",",password+",",offset,")"
					(msg,offset) = neos.getIntermediateResultsNonBlocking(jobNumber,password,offset)
					#(msg,offset) = neos.getIntermediateResults(jobNumber,password,offset)
					if showLog: print ":::: getIntermediateResults result="
					sys.stdout.write(msg.data)
					if showLog: print ":::: getIntermediateResults result ends"
				except System.IO.IOException as ex:
					if showLog: print ":::: getIntermediateResults timed out"
				if showLog: print ":::: calling getJobStatus(", jobNumber,","+password+")"
				status = neos.getJobStatus(jobNumber, password)
				if showLog: print "::::  status =",status
		if status=="Done":
			if showLog: print "calling getFinalResults(",jobNumber,",",password,")"
			msg = neos.getFinalResults(jobNumber, password).data
			if showLog: print "::::  .data="
			if showLog: print msg
			if showLog: print "::::  .data ends"
			# As of 7/4/2014, NEOS separates lines by LF LF (i.e. 0A 0A); remove all the second LF's
			msg=msg.replace("\n\n","\n") # Quick fix for NEOS double line feeds
			sys.stdout.write("## NEOS Status: Job Done\nJob Output:\n")
			sys.stdout.write(msg)
			# print "Hex output"
			# print ":".join("{:02x}".format(ord(c)) for c in msg) # Print string in hdex for debuggging
			print
			with open(resultsfilename,'w') as content_file:
				content_file.write(msg)
			completed = True
		else:
			print "## Error: NEOS reported %s" % status
	except System.IO.IOException as ex:
		print "## TimeOut/IO Error: No NEOS response in "+str(socket.getdefaulttimeout())+"s."
		raise Exception("TimeOut/IO Error: The NEOS server did not respond within "+str(socket.getdefaulttimeout())+" seconds, and so the job has been cancelled by SolverStudio. No results are available.")
	except System.ThreadAbortException  as ex: # The thread is being aborted as a result of the user cancelling
		print "## Aborted: NEOS run aborted by the user."
		raise Exception("Aborted: NEOS run aborted by the user.")
	except System.Exception as ex: # .Net error
		print "## Error: A system error occurred while communicating with NEOS\nError Message: "+ex.Message
		raise Exception("Error: An system error occurred while communicating with NEOS. No results are available.\n\nError is:\n"+ex.Message)
	except Exception as ex: # A general Python exception
		print "## Error: Python reported an error while communicating with NEOS.\nError Message: "+str(ex)
		raise Exception("Error: Python reported an error while communicating with NEOS. No results are available.\n\nError is:\n"+str(ex))
	finally:
		if jobNumber != 0 and status!="Done" and status!="NotStarted":
			try:
				neos.killJob(jobNumber, password, '')
				print "## Killed NEOS job %d" % jobNumber
			except:
				pass
		socket.setdefaulttimeout(oldTimeout)
	return completed

def readAMPLresultsStyle2(resultsfilename):
	try:
		g = open(resultsfilename,"r")
		lineCount = 0
		itemsLoaded = ""
		itemsChanged = False
		standardErrorHelp = "Please check that your model has a a 'display_1col' option, followed by one or more 'display' commands each displaying one item; e.g.\n\noption display_1col 9999999;\ndisplay rosteredOn > Sheet;\ndisplay rosteredOff > Sheet;\n"
		while True :
			line = g.readline()
			if not line: break
			lineCount += 1
			if line=="": continue # skip blank lines
			items = csv.reader([line], delimiter=' ', skipinitialspace=True, quotechar="'").next()
			if len(items)==0: continue
			if items[0] != '_display': continue # Note that this skips any empty items that are output on a single line as, eg: nbr; #empty
			numIndices = int(items[1])
			numValues = int(items[2])
			numRows = int(items[3])
			name = g.readline()
			name = name.strip()  # remove line feed at end
			escapedName = SolverStudio.EscapedDataItemName(name) # The Python name used internally
			lineCount += 1
			if name=="":
					raise Exception("An unexpected entry has occurred when writing the new values to the sheet; the data item name is missing in the temporary file.\n\n"
					+standardErrorHelp+"\n\n(The error occurred in the string '"+repr(line)+"' on line "+repr(lineCount)+" of file "+repr(resultsfilename)+".)\n\n")
			try:      
				dataItem = SolverStudio.DataItems[escapedName]
			except:
				raise Exception("A value was written to the sheet for data item '"+repr(name)+"', but this data item has not been defined in your spreadsheet.\n\n"
				+standardErrorHelp+"\n\n(The error occurred in the string "+repr(line)+" on line "+repr(lineCount)+" of file "+repr(resultsfilename)+".)")
			changed = False
			for i in range(numRows):
				line = g.readline()
				lineCount += 1
				items = csv.reader([line], delimiter=',', skipinitialspace=True, quotechar="'").next()
				if len(items) != numIndices+numValues:
					raise Exception("An unexpected number of items found when reading the new value for data item '"+repr(name)+"'.\n\n"
					+standardErrorHelp+"\n\n(The error occurred in the string "+repr(line)+" on line "+repr(lineCount)+" of file "+repr(resultsfilename)+".)")
				if numIndices==0:
					value = items[-1]
					try:
						value = float(items[-1])
					except:
						pass	#raise Exception("When processing the value written to the sheet for data item '"+repr(name)+"', an error occurred when converting the item '"+repr(items[-1])+"' to a number.\n\n"+standardErrorHelp+"\n\n(The error occurred in string '"+repr(line)+"' in line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")
					try:
						if dataItem!=value:
							# This fails inside a function; needs globals() instead
							# vars()[name]=value # Update the python variable of this name (without creating a new variable, which would not be seen by SolverStudio               
							globals()[escapedName]=value # Update the python variable of this name (without creating a new variable, which would not be seen by SolverStudio)
							changed = True
					except:
						raise Exception("When writing new values for data item "+repr(name)+" to the sheet, an error occurred when assigning the value "+repr(items[-1])+" to "+repr(name)+".\n\n"+standardErrorHelp+"\n\n(The error occurred in string '"+repr(line)+"' in line "+repr(lineCount)+" of file "+repr(resultsfilename)+".)")
				else:
						value = items[-1]
						try:
							value = float(items[-1])
						except:
							pass	# raise Exception("When writing new values for data item "+repr(name)+" to the sheet, an error occurred in line "+repr(line)+" when converting the last item '"+repr(items[-1])+"' to a number.\n\n"
							#+standardErrorHelp+"\n\n(Line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")
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
			if changed:
				itemsLoaded = itemsLoaded+name+"* "
				itemsChanged = True
			else:
				itemsLoaded = itemsLoaded+name+" "
	finally:
		g.close()
	if itemsLoaded=="":
		print "## No results were loaded into the sheet."
	elif not itemsChanged:
		print "## Results loaded for data items:", itemsLoaded
	else :
		print "## Results loaded for data items:", itemsLoaded
		print "##   (*=data item values changed on sheet)"

def DoSolveOnNeos():
	dataFilename = "SheetData.dat"
	g = open(dataFilename,"w")
	g.write("# "+language+" data file created by SolverStudio\n\n")
	g.write("data;\n\n")

	print "## Scanning model for sets and parameters"
	Sets, Params,solveCommand,dataCommand,displayCommand,Variables,solver = ScanModelFileForDataItems(SolverStudio.ModelFileName)

	if SolverStudio.ActiveModelSettings.GetValueOrDefault("FixMinorErrors",True):
		DoSmallFixes(solveCommand,dataCommand,displayCommand,Variables,solver)
		
	print "## Building input file for %s sets and %s parameters"%(len(Sets),len(Params))

	print "## Writing simple parameters...",
	count = 0
	for (escapedName,value) in SolverStudio.SingletonDataItems.iteritems():
		name = SolverStudio.UnEscapedDataItemName(escapedName)
		if name in Params:
			count += 1
			WriteParam(g, name,value)
	print count,"items written."
			
	print "## Writing sets...", #  Sets may be Python singletons or lists
	count = 0
	for (escapedName,value) in SolverStudio.SingletonDataItems.iteritems():
		name = SolverStudio.UnEscapedDataItemName(escapedName)
		if name in Sets:
			count += 1
			WriteSingletonSet(g, name,value)
	for (escapedName,value) in SolverStudio.ListDataItems.iteritems():
		name = SolverStudio.UnEscapedDataItemName(escapedName)
		if name in Sets:
			count += 1
			WriteSet(g, name,value)
	for (escapedName,value) in SolverStudio.DictionaryDataItems.iteritems():
		name = SolverStudio.UnEscapedDataItemName(escapedName)
		if name in Sets:
			count += 1
			WriteIndexedSet(g, name,value)
	print count,"items written."
		
	print "## Writing indexed parameters...",
	count = 0
	for (escapedName,value) in SolverStudio.DictionaryDataItems.iteritems():
		name = SolverStudio.UnEscapedDataItemName(escapedName)
		if name in Params:
			count += 1
			WriteIndexedParam(g, name,value)
	print count,"items written."

	g.close()

	resultsfilename = "Sheet"
	g = open(resultsfilename,"w")
	g.close()

	print "## Running "+language+"..."
	print "## "+language+" model file:",SolverStudio.ModelFileName
	print "## "+language+" data file:",dataFilename+"\n"
	if useNEOS:
		if language=="AMPL":
			succeeded = runAMPLonNEOS(SolverStudio.ModelFileName, dataFilename, resultsfilename)
		else:
			raise Exception("SolverStudio does not yet support GAMS on NEOS.")
		if succeeded:
			print "## "+language+" run completed on NEOS."
		else:
			print "## "+language+" did not complete; no solution is available. The sheet has not been changed."
		readAMPLresultsStyle2(resultsfilename)   
	else:
		if language=="AMPL":
			exitCode = SolverStudio.RunExecutable(SolverStudio.GetAMPLPath(),'"'+SolverStudio.ModelFileName+'"', addExeDirectoryToPath=True)
		else:
			exitCode = SolverStudio.RunExecutable(SolverStudio.GetGLPSolPath(),"--model \""+SolverStudio.ModelFileName+"\""+" --data \""+dataFilename+"\"")
		if exitCode==0:
			print "## "+language+" run completed."
			readAMPLresults(resultsfilename)   
		else:
			print "## "+language+" did not complete (Error Code %d); no solution is available. The sheet has not been changed." % exitCode

	print
	print "## Done"

# The main run code must be in a subroutine to avoid clashes with SolverStudio globals (which can cause COM exceptions in, eg, "for (name,value) in SolverStudio.SingletonDataItems.iteritems()" if value exists as a data item)
def DoRun():
	WarningShown = SessionSettings.GetValueOrDefault("AMPLNEOSWarningShown",False)
	if not WarningShown and NeosWarningEnabled():
		if ConfirmNEOS():
			SessionSettings["AMPLNEOSWarningShown"] = True
			DoSolveOnNeos()
	else:
		DoSolveOnNeos()

DoRun()
