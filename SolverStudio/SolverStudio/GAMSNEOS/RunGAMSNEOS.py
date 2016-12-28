# 20141104: pATCHED TO ALLOW TIMESOUTS in NEOS comms
# 20141105 Convert any "]]>" into "]] >" in the model so it can be embedded in a CData object (for which ]]> is special)

#import csv
import os
import sys
import xmlrpclib
import time
import re
import base64
import zipfile
from shutil import copy2
import socket

from _winreg import *

def ScanModelFileForDataItems(modelFileName):
	unusedKeys = ["positive","equations","model","solve","scalar"]#,"variables","display"]
	Sets= list()
	Params = list()
	Tables = list()
	Variables = {}
	AllVars=list()
	unloadCmd = False
	#singlelineKeyWords = { "set" : Sets, 'parameter' : Params, 'table' : Tables }
	keyWords = { "sets" : Sets, 'parameters' : Params, 'tables' : Params, "set" : Sets, 'parameter' : Params, 'table' : Params,"variables":AllVars}#,"display":Variables }
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
				#elif items != AllVars:
				#	print tokens[0]
					#raise Exception("## SolverStudio: Not data item exists for: " + tokens[0] + ". Unable to create GDX file. Create a data item first")
		if tokens[0].lower()=="display":
			for i in tokens[1:]:
				level=0
				if len(i)>3:
					if i[-2]==".":
						if i[-1]=="m": level = 1
						i=i[:-2]
				if i in SolverStudio.DataItems and not i in Variables:
					Variables[i]=level
		if ";" in tokens:
			items = None
	f.close()
	if Variables == {}:
		for i in AllVars:
			Variables[i]=0
	return Sets, Params, Variables
	
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
          try:
             job = str(items[0])
             if job.find(str(jobNumber)) != -1:
                position = numQueued
          except:
             pass
          
   return numRunning, numQueued, position

# Replace any "]]>" with "]] >" because "]]>" is the end-of-CData delimiter   
def MakeModelSafeForCData(model):
   r = re.compile(r']]>',re.IGNORECASE)
   model = r.sub(']] >',model)
   return model

def BuildXML(model,dataFilename):  
   oldDirectory = SolverStudio.ChangeToLanguageDirectory()
   with open('NEOSTemplate.xml', 'r') as content_file:
     xml= content_file.read()
   SolverStudio.SetCurrentDirectory(oldDirectory)
   # with open('GAMSdata.gdx',"r") as data_file:
      # data = base64.encodestring(data_file.read());
   data = encodeGzip(dataFilename,dataFilename+'.gz')
   if SolverStudio.ActiveModelSettings.GetValueOrDefault("ShortQueueLength",False):
      priority="short"
   else:
      priority="long"
   solver =  SolverStudio.ActiveModelSettings.GetValueOrDefault("GAMSNeosSolver",'cbc')
   category =  SolverStudio.ActiveModelSettings.GetValueOrDefault("GAMSNeosCategory",'milp')
   email = SolverStudio.GetRegistrySetting("NEOS","Email","")  # TODO: Encode this to be safe, eg urllib.urlencode?
   xml=xml%(category,solver,priority,email,"<![CDATA[",MakeModelSafeForCData(model),"]]>",data)
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
	
def runGAMSonNEOS(zipfilename):
	import clr # Need this to define System.IO.IOException exception
	clr.AddReference('System')  # Need this to define System.IO.IOException exception; note that System.IO is defined in 
										# assembly mscorlib, not System.IO, so clr.AddReference('System.IO') is wrong as there is no 
										# such assembly. This fails for others, even tho' it works ok for me
	import System
	f=open("NEOSJob.xml")
	xml=f.read()
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
				time.sleep(5)  # Changed from 1 to 5 to avoid NEOs rejecting too many requests (20150520)
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
						#if posn == -1:
						#	status = neos.getJobStatus(jobNumber, password)
						#	if status == "Waiting":
						#		print "Position in queue could not be found."
						#else:
						print "## Waiting:%3d/%2d in queue (%2d jobs running)" % (posn, numQueued, numRunning)
		if status == "Running":
			print "## NEOS Status: Running..."
			offset=0
			while status == "Running":
				time.sleep(5)  # Changed from 1 to 5 to avoid NEOs rejecting too many requests (20150520)
				try:
					if showLog: print ":::: calling getIntermediateResults(",jobNumber,",",password+",",offset,")"
					(msg,offset) = neos.getIntermediateResults(jobNumber,password,offset)
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
			sys.stdout.write("## NEOS Status: Job Done\n")#Job Output:\n")
			sys.stdout.write(msg)
			completed=(not "USER ERROR(S) ENCOUNTERED" in msg)
			if completed:
				print "## Downloading GDX result file"
				try:
					from urllib import urlretrieve
					url = "http://neos-server.org/neos/jobs/"+str(jobNumber/10000)+"0000/"+str(jobNumber)+"-"+str(password)+"-solver-output.zip"
					zipfilename,hdrs = urlretrieve(url,zipfilename) # This may get an HTML file containing: ... The requested URL /neos/jobs/3750000/3754262-WtCmufVq-solver-output.zip was not found on this server. ...
					# TODO  hdrs contains "Last-Modified: Thu, 11 Jun 2015 09:07:34 GMT" and "ETag: ..." and "Accept-Ranges: bytes:" only if a .zip was actually downloaded
					print "## Download complete"
				except:
					print "## Unable to download GDX result file"
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

def FixInput(model,Sets,Params):
	sList = ""
	for i in Sets+Params:
		sList = sList + i + ", "
	r=None
	l=re.compile(r'^\$\s*GDXIN\s*"SheetData.gdx".*\n*^\$\s*LOAD.*',re.IGNORECASE | re.MULTILINE | re.DOTALL)
	if not re.search(l,model):
		r1=re.compile(r'^\$\s*GDXIN\s*"SheetData.gdx".*',re.IGNORECASE | re.MULTILINE)
		r2=re.compile(r'^\s*\$\s*include\s*"SheetData.dat".*',re.IGNORECASE | re.MULTILINE)
		if re.search(r1,model): r=r1
		if re.search(r2,model): r = r2
		if r:
			model=r.sub('\n$GDXIN "in.gdx"\n***SolverStudio: Above Line Modified\n$LOAD '+sList[:-2]+"\n***SolverStudio: Above Line Modified\n$GDXIN\n***SolverStudio: Above Line Modified\n",model)
		else:
			r=re.compile(r'^\s*solve ',re.IGNORECASE|re.MULTILINE)
			model=r.sub('\n$GDXIN "in.gdx"\n***SolverStudio: Above Line Modified\n$LOAD '+sList[:-2]+"\n***SolverStudio: Above Line Modified\n$GDXIN\n***SolverStudio: Above Line Modified\nsolve ",model)
	return model

	
def FixOutput(model,Variables):
	r=re.compile(r'^display.*',re.IGNORECASE | re.MULTILINE)
	m = re.search(r,model)
	if not m==None: 
		line=m.group(0)
		for i in Variables:
			if Variables[i]==0:
				if (not (i+".l") in line) and (i in line):
					line=line.replace(i,i+".l")
		model = r.sub(line,model)
	return model

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
			msg += "Some solvers require an email address, which you can set using the GAMS menu. "
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

def DoSolveOnNeos():
	#Check gdxdclib.dll in temp dir. UNUISED NOW AS GAMS DLL'S ON SolverStudio's PATH
	#if SolverStudio.Is64Bit():
	#	DLLpath=SolverStudio.GetGAMSDLLPath() + "\\gdxdclib64.dll"
	#else:
	#	DLLpath=SolverStudio.GetGAMSDLLPath() + "\\gdxdclib.dll"
	#if not os.path.exists(SolverStudio.WorkingDirectory() + "\\gdxdclib.dll"):
	#	copy2(DLLpath,SolverStudio.WorkingDirectory() + "\\gdxdclib.dll")

	resultsfilename = "out.gdx"
	f = open(resultsfilename,"w")
	f.close()
	
	zipfilename = "results.zip"
	g = open(zipfilename,"w")
	g.close()

	print "## Scanning model for sets and parameters"
	Sets, Params, Variables = ScanModelFileForDataItems(SolverStudio.ModelFileName)
	#print "Sets=",Sets
	#print "Params=",Params

	with open(SolverStudio.ModelFileName,"r") as inFile:
			model = inFile.read()
	if SolverStudio.ActiveModelSettings.GetValueOrDefault("FixMinorErrors",True):
		model = FixInput(model,Sets,Params)
	model = FixOutput(model,Variables)
	model = model.replace("SheetData.gdx","in.gdx")
	with open(SolverStudio.ModelFileName,"w") as outFile:
			outFile.write(model)
			
	dataFilename = "SheetData.gdx"
	print "## Building GAMS input file for %s sets and %s parameters"%(len(Sets),len(Params))

	PGX = SolverStudio.OpenGDX(dataFilename,False)
	for i in Sets:
		SolverStudio.WriteSetToGDX(PGX,i,SolverStudio.DataItems)
	for i in Params:
		SolverStudio.WriteParamToGDX(PGX,i,SolverStudio.DataItems)
	SolverStudio.CloseGDX(PGX)
	print "## GAMS input file built successfully."

	with open(SolverStudio.ModelFileName,"r") as content_file:
		model = content_file.read()
	BuildXML(model,dataFilename)
	built = True
	succeeded = False

	if built:
		print "\n## Running GAMS on NEOS"
		print "## GAMS model file:",SolverStudio.ModelFileName	
		print "## GAMS data file:",dataFilename + "\n"
		succeeded = runGAMSonNEOS(zipfilename)

	if succeeded:
		# Unzip data
		# The .zip file may be an HTML file-not-found error, so we trap thhis
		try:
			fh = open(zipfilename, 'r')
			try:
				z = zipfile.ZipFile(fh) # May fail if we retrieved an error message, not a .zip file
				z.extract(resultsfilename)
			except:
				succeeded = False
			fh.close()
		except:
			succeeded = False

	if succeeded:
		print "\n## GAMS completed successfully.\n## Reading results..."
		# Read Results
		try:
			PGX = SolverStudio.OpenGDX(resultsfilename)
			itemsLoaded = list()
			for i in Variables:
				try:
					SolverStudio.ReadSymbolFromGDX(PGX,i,Variables[i],SolverStudio.DataItems)
					itemsLoaded.append(i)
				except:
					print "## Unable to load: " + i
			SolverStudio.CloseGDX(PGX)
			if len(itemsLoaded)==0:
				print "## No results were loaded into the sheet."
			else:
				print "## Results loaded for data items:", itemsLoaded
		except:
			print "## Error reading results"
	else:
		print "## GAMS did not complete; no solution is available."
	print "## Done"

# The main run code must be in a subroutine to avoid clashes with SolverStudio globals (which can cause COM exceptions in, eg, "for (name,value) in SolverStudio.SingletonDataItems.iteritems()" if value exists as a data item)
def DoRun():
	WarningShown = SessionSettings.GetValueOrDefault("GAMSNEOSWarningShown",False)
	if not WarningShown and NeosWarningEnabled():
		if ConfirmNEOS():
			 SessionSettings["GAMSNEOSWarningShown"] = True
			 DoSolveOnNeos()
	else:
		DoSolveOnNeos()

DoRun()
