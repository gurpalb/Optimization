import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
# clr.AddReference('Microsoft.Office.Tools.Ribbon');
import webbrowser
import os
import re
#from System.Windows.Forms import Application, Form, Label, ToolStripMenuItem
from System.Windows.Forms import Form, ToolStripMenuItem, MessageBox, MessageBoxButtons, DialogResult, Cursor, Cursors, Button, TextBox, Label, FormBorderStyle, LinkBehavior,LinkArea,LinkLabel

clr.AddReference('System.Drawing')
from System.Drawing import Point,Color

import version
# import InstallAMPLCML

def VisitAMPLWebPage(key, e):
  webbrowser.open('http://ampl.com')
  
def VisitNEOSWebPage(key, e):
  webbrowser.open('http://www.neos-server.org/')
  
def VisitAMPLBookWebPage(key, e):
  webbrowser.open('http://www.ampl.com/BOOK/download.html')
  
def OpenAMPLpdf(key,e):
  webbrowser.open('http://www.ampl.com/REFS/amplmod.pdf')

def OpenDataFile(key,e):
  # We need to get access to the temporary directory in our SolverStudio object...
  workingDir = SolverStudio.WorkingDirectory()
  dataFileName = "SheetData.dat"
  dataFilePath = os.path.realpath(os.path.normpath(os.path.join(workingDir, dataFileName)))
  #form = Form(Text="workingDir="+workingDir+"\ndataFileName="+dataFileName)
  #form.Controls.Add(Label(Text="workingDir="+workingDir+"\ndataFileName="+dataFileName))
  #form.Controls.Add(Label(Text="dataFilePath="+dataFilePath))
  #form.ShowDialog()
  #if os.path.exists( dataFilePath ):
  SolverStudio.ShowTextFile(dataFilePath) # This will show the file and update as the file is changed (or created)

def About(key,e):
  MessageBox.Show("SolverStudio AMPLNEOS Processor:\n"+version.versionString)

def MenuTester2(key,e):
  #form = Form(Text="Hello World Form Menu Tester")
  #label = Label(Text="Hello World!")
  #form.Controls.Add(label)
  #form.ShowDialog()
  pass

def InstallAMPLMenuHandler(key,e):
  SolverStudio.RunInDialog("This will download the AMPL Student Edition from AMPL.com, and install it into the SolverStudio directory. It will not replace any other installed AMPL versions you may have.",InstallAMPLWorker)
  
def UpdateSolverListWorker(key,e):
	try:
		import xmlrpclib
		NEOS_HOST="neos-server.org"
		NEOS_PORT=3332
		neos=xmlrpclib.ServerProxy("http://%s:%d" % (NEOS_HOST, NEOS_PORT))
		##=============================================================
		##Currently there is only a single NEOS template file. As such only solvers with
		##matching templates can be used.
		##=============================================================
		oldDirectory = SolverStudio.ChangeToLanguageDirectory()
		with open("NEOSSolverList.txt","r") as in_file:
			currentSolvers = eval(in_file.read())
		print
		print
		print "## Getting list of Solvers"
		currentSolvers, categoryDict = ChooseSolver()
		listOfSolvers = neos.listAllSolvers()
	except:
		choice = MessageBox.Show("Unable to contact the NEOS Server. Please check your internet connection and try again.","SolverStudio",MessageBoxButtons.OK)
		return
	amplSolvers = {}
	for i in listOfSolvers:
		if "AMPL" in i:
			colIndex=i.index(":")
			category = i[:colIndex]
			solver = i[colIndex+1:-5]
			# if solver=="Gurobi": continue
			if category in amplSolvers:
				amplSolvers[category] = amplSolvers[category] + [solver]
			else:
				amplSolvers[category] = [solver]
	print "## Checking SolverStudio compatibility"
	template = neos.getSolverTemplate('milp','Cbc','AMPL')
	template = template[template.index("</inputMethod>"):]
	newSolvers={}
	for i in amplSolvers:
		for j in amplSolvers[i]:
			testXML = neos.getSolverTemplate(i,j,'AMPL')
			testXML = testXML[testXML.index("</inputMethod>"):]
			if testXML == template:
				if i in newSolvers:
					newSolvers[i]=newSolvers[i] + [j]
				else:
					newSolvers[i]=[j]
	added = {}
	removed = {}
	print
	#Check for changes
	if not newSolvers == currentSolvers: #Something changed
		for i in newSolvers: #for all new categories
			if not i in currentSolvers:
				## Added new category
				added[i]=newSolvers[i]
				continue
			if newSolvers[i]==currentSolvers[i]: continue #no change
			for j in currentSolvers[i]: #for all solvers in changed category
				if not j in newSolvers[i]:
					####removed j from category i
					if i in removed:
						removed[i].append(j)
					else:
						removed[i] = [j]
			for j in newSolvers[i]:
				if not j in currentSolvers[i]:
					####added j to category i
					if i in added:
						added[i].append(j)
					else:
						added[i] = [j]
		for i in currentSolvers:
			exists = False
			for j in newSolvers:
				if i == j:
					exists = True
			if exists == False:
				#### category removed
				removed[i]=currentSolvers[i]
		for i in added:
			for j in added[i]:
				print "## Added " + j + " to " + categoryDict[i]
		for i in removed:
			for j in removed[i]:
				if i in categoryDict:
					print "## Removed " + j + " from " + categoryDict[i]
				else:
					print "## Removed " + j + " from " + i
		with open("NEOSSolverList.txt","w") as dest_file:
			dest_file.write(str(newSolvers))
	else:
		print
		print "## AMPL Solver list is already up to date."
	SolverStudio.SetCurrentDirectory(oldDirectory)
	print
	print "## DONE."
	print "## Restart SolverStudio to use updated solvers."
	##=============================================================	

def UpdateSolverListMenuHandler(key,e):
	SolverStudio.RunInDialog("Updating the NEOS Solver List for AMPL...",UpdateSolverListWorker,True,True,True)

def ChooseSolver():
	import xmlrpclib
	NEOS_HOST="neos-server.org"
	NEOS_PORT=3332
	neos=xmlrpclib.ServerProxy("http://%s:%d" % (NEOS_HOST, NEOS_PORT))
	categoryDict = neos.listCategories()
	oldDirectory = SolverStudio.ChangeToLanguageDirectory()
	with open("NEOSSolverList.txt","r") as in_file:
		sGoodSolvers = in_file.read()
	SolverStudio.SetCurrentDirectory(oldDirectory)
	return eval(sGoodSolvers), categoryDict
	
def ReadSolversFiles():
	oldDirectory = SolverStudio.ChangeToLanguageDirectory()
	with open("NEOSSolverList.txt","r") as in_file:
		sGoodSolvers = in_file.read()
	with open("NEOSCategoryList.txt","r") as in_file:
		categoryDict = in_file.read()
	SolverStudio.SetCurrentDirectory(oldDirectory)
	return eval(sGoodSolvers), eval(categoryDict)

class NEOSEmailForm(Form):

	def __init__(self):
		self.Text = "SolverStudio NEOS Email Address"
		self.FormBorderStyle = FormBorderStyle.FixedDialog    
		self.Height= 140
		self.Width = 460

		self.label = Label()
		self.label.Text = "Some commercial solvers on NEOS, such as CPLEX, require a valid email address. Please enter the Email address, if any, you wish to send to NEOS when you solve:"
		self.label.Location = Point(10,10)
		self.label.Height  = 35
		self.label.Width = 430

		self.KeyText = TextBox()
		self.KeyText.Location = Point(10,46)
		self.KeyText.Width = 430
		self.KeyText.Height  = 20
		self.KeyText.Text = SolverStudio.GetRegistrySetting("NEOS","Email","")
		self.KeyText.AcceptsReturn = False;

		self.bOK=Button()
		self.bOK.Text = "OK"
		self.bOK.Location = Point(270,76)
		# self.bOK.Width = 100
		self.bOK.Click += self.SetEmail

		self.link1 = LinkLabel()
		self.link1.Location = Point(10, 78)
		self.link1.Width = 250
		self.link1.Height = 40
		self.link1.LinkClicked += self.OpenTCs
		self.link1.VisitedLinkColor = Color.Blue;
		self.link1.LinkBehavior = LinkBehavior.HoverUnderline;
		self.link1.LinkColor = Color.Navy;
		self.link1.Text = "View NEOS terms and conditions"
		self.link1.LinkArea = LinkArea(5,25)

		self.bCancel=Button()
		self.bCancel.Text = "Cancel"
		self.bCancel.Location = Point(363,76)

		self.AcceptButton = self.bOK
		self.CancelButton = self.bCancel

		self.Controls.Add(self.label)
		self.Controls.Add(self.KeyText)	# Add first to get focus
		self.Controls.Add(self.link1)
		self.Controls.Add(self.bOK)
		self.Controls.Add(self.bCancel)
		self.CenterToScreen()

	def SetEmail(self,sender, event):
		SolverStudio.SetRegistrySetting("NEOS","Email",self.KeyText.Text)
		self.DialogResult = DialogResult.OK

	def OpenTCs(self,sender,event):
		import webbrowser
		webbrowser.open("""http://www.neos-server.org/neos/termofuse.html""")

def SetNEOSEmail(key,e):
	myNeosEmailForm = NEOSEmailForm()
	myNeosEmailForm.ShowDialog()

def DoNEOSTest():
	import socket
	import xmlrpclib
	clr.AddReference('System')  # Need this to define System.IO.IOException exception; note that System.IO is defined in 
													# assembly mscorlib, not System.IO, so clr.AddReference('System.IO') is wrong as there is no 
													# such assembly. This fails for others, even tho' it works ok for me
	import System
	clr.AddReference('System.Reflection')  # Need this to define System.IO.IOException exception; note that System.IO is defined in assembly mscorlib, not System.IO, so clr.AddReference('System.IO') is wrong as there is no such assembly. This fails for others, even tho' it works ok for me
	clr.AddReference('System.Net');
	clr.AddReference('System.Net.NetworkInformation');

	if not System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable():
		MessageBox.Show("This computer does not appear to have an active network connection, and so it cannot communicate with the NEOS servers.", "SolverStudio NEOS Test: Failed")
		#return

	NEOS_HOST="neos-server.org"
	NEOS_PORT=3332

	# Do a full test
	# Change to a 60s timeout (by default, there is no timeout)
	oldTimeout = socket.getdefaulttimeout()
	socket.setdefaulttimeout(60)
	neos=xmlrpclib.ServerProxy("http://%s:%d" % (NEOS_HOST, NEOS_PORT))

	xmlrpcSucceeded = False
	xmlrpcResult = "Unable to connect to NEOS."
	try:
		xmlrpcResult = "The connection to the NEOS solvers at "+NEOS_HOST+" appears to be working.\nThe NEOS host reported: "+neos.ping()
		xmlrpcSucceeded = True
	except System.IO.IOException as ex:
		print "System.IO.IOException"    
		xmlrpcResult = "NEOS Solvers Connection: Fail (timeout after  "+str(socket.getdefaulttimeout())+"s.)\n  No response from NEOS XML-RPC server "+NEOS_HOST+":"+str(NEOS_PORT)+"\n  The NEOS solvers may be offline\n"
	except System.Exception as ex: # .Net error
		xmlrpcResult = "NEOS Solvers Connection: Fail ("+ex.Message+")\n  No response from NEOS XML-RPC server "+NEOS_HOST+":"+str(NEOS_PORT)+"\n  The NEOS solvers may be offline\n"
	except Exception as ex: # A general Python exception
		xmlrpcResult = "NEOS Solvers Connection: Fail ("+ex.Message+")\n  No response from NEOS XML-RPC server "+NEOS_HOST+":"+str(NEOS_PORT)+"\n  The NEOS solvers may be offline\n"
	finally:
		socket.setdefaulttimeout(oldTimeout)

	print "xmlrpcResult="+xmlrpcResult

	if xmlrpcSucceeded:
		MessageBox.Show(xmlrpcResult,"SolverStudio NEOS Test: Success")
		return

	# Try a sequence of tests
	# Do a DNS lookup
	NeosDNSResult = "Unable to find IP address"
	dnsResolutionSucceeded = False;
	try:
		NeosIPAddressList = System.Net.Dns.GetHostByName(NEOS_HOST).AddressList;
		ipAddresses = ""
		for address in NeosIPAddressList:
				if ipAddresses <> "":
					ipAddresses = ipAddress+", "
				ipAddresses = ipAddresses+address.ToString()
		NeosDNSResult = "DNS Lookup: Success ("+ipAddresses+")\n  The NEOS server's address is known.\n"
		dnsResolutionSucceeded = True;
	except Exception as ex2:
		NeosDNSResult = "DNS Lookup: Fail ("+str(ex2)+")\n  The NEOS server's address is not known\n"

	print "NeosDNSResult="+NeosDNSResult

	# Try a PING test
	pingSender = System.Net.NetworkInformation.Ping()
	try:
		pingReply = pingSender.Send(NEOS_HOST)
		pingSucceeded = pingReply.Status == System.Net.NetworkInformation.IPStatus.Success;
		pingResult = "Ping "+NEOS_HOST+": "+ (("Success ("+pingReply.RoundtripTime.ToString()+" ms roundtrip)\n  The NEOS computers are online\n") if pingSucceeded else ("Fail ("+pingReply.Status.ToString()+")\n  The NEOS computers may be offline\n"))
	except System.Net.NetworkInformation.PingException as ex:
		pingResult = "Ping "+NEOS_HOST+": Fail\n  The NEOS computers could not be pinged ("+str(ex.InnerException.Message)+")\n"
	except Exception as ex:
		pingResult = "Ping "+NEOS_HOST+": Fail\n  The NEOS computers could not be pinged ("+ex.Message+")\n"

	print "pingResult="+pingResult

	# Try a PING test of Google 8.8.8.8
	googlePingSender = System.Net.NetworkInformation.Ping()
	try:
		googlePingReply = pingSender.Send("8.8.8.8")
		googlePingSucceeded = googlePingReply.Status == System.Net.NetworkInformation.IPStatus.Success;
		googlePingResult = "Ping 8.8.8.8: "+ ("Success\n  Your internet connection works\n" if googlePingSucceeded else ("Fail ("+googlePingReply.Status.ToString()+")\n  Your internet connection does not appear to be working\n"))
	except System.Net.NetworkInformation.PingException as ex:
		googlePingResult = "Ping 8.8.8.8: Fail\n  The Google machine at 8.8.8.8 could not be pinged ("+str(ex.InnerException.Message)+")\n"
	except Exception as ex:
		googlePingResult = "Ping 8.8.8.8: Fail\n  The Google machine at 8.8.8.8 could not be pinged ("+ex.Message+")\n"

	print "googlePingResult ="+googlePingResult 

	MessageBox.Show("The NEOS solvers could not be reached. Please use the following to help diagnose the problem.\n\nNEOS_HOST="+NEOS_HOST+"\n"+NeosDNSResult+"\n"+xmlrpcResult+"\n"+pingResult+"\n"+googlePingResult,"SolverStudio NEOS test: Fail")

def TestNEOSConnection(key,e):
	try:
		 Cursor.Current = Cursors.WaitCursor;
		 DoNEOSTest()
	finally:
		 Cursor.Current = Cursors.Default; 

# Return s as a string suitable for indexing, which means tuples are listed items in []
def Initialise():
	#Menu.DropDownItems.Add( ToolStripMenuItem("Test",None,MenuTester) )
	Menu.ToolStripMenuItem.MouseDown += Update
	Menu.Add( "View Last Data File",OpenDataFile)
	Menu.AddSeparator()
	global doFixChoice
	doFixChoice=Menu.Add( "Fix Minor Errors",doFixClickMenu)
	Menu.AddSeparator()
	global queueChoice
	queueChoice=Menu.Add( "Run in short NEOS queue",QueueClickMenu)
	Menu.Add( "Set Email Address for NEOS...",SetNEOSEmail)
	chooseSolverMenu = Menu.Add ("Choose Solver",emptyClickMenu)
	amplSolvers, categoryDict = ReadSolversFiles()
	global solverMenu
	solverMenu = {}
	try:
		for i in amplSolvers:
			solverMenu[i]=[]
			categoryMenu = ToolStripMenuItem(categoryDict[i],None,emptyClickMenu)
			solverMenu[i].append(categoryMenu)
			chooseSolverMenu.DropDownItems.Add(categoryMenu)
			solverMenu[i].append({})
			for j in amplSolvers[i]:
				solverMenu[i][1][j] = ToolStripMenuItem(j,None,SolverClickMenu)
				categoryMenu.DropDownItems.Add(solverMenu[i][1][j])
	except:
		choice = MessageBox.Show("A NEOS category is no longer supported by AMPLonNEOS.\n\nDo you wish to update the NEOS solver list now?","SolverStudio",MessageBoxButtons.OK)
		if choice == DialogResult.OK:
			SolverStudio.RunInDialog("NEOS Solver List Update for AMPL.",UpdateSolverListWorker,True,True,True)
	Menu.Add( "Update NEOS Solvers...",UpdateSolverListMenuHandler)
	Menu.Add( "Test NEOS Connection...",TestNEOSConnection)
	Menu.AddSeparator()
	Menu.Add( "Open AMPL.com",VisitAMPLWebPage)
	Menu.Add( "Open online AMPL book",VisitAMPLBookWebPage)
	Menu.Add( "Open online AMPL PDF documentation",OpenAMPLpdf)
	# Menu.Add( "-",None) This does not work :-(
	Menu.Add( "Open neos-server.org",VisitNEOSWebPage)
	Menu.AddSeparator()
	Menu.Add( "About AMPL on NEOS Processor",About)

def Update(key,e):
   # The Solver is either the value saved, or the solver specified in the model file
	solver = SolverStudio.ActiveModelSettings.GetValueOrDefault("AMPLNeosSolver",'Cbc').lower()
	category = SolverStudio.ActiveModelSettings.GetValueOrDefault("AMPLNeosCategory",'milp')
	# Look for a solver in the model, first
	model = SolverStudio.GetActiveModelText()
	r = re.compile(r"^\s*option\s+solver\s+(\w+)\s*;",re.IGNORECASE | re.MULTILINE)
	m=r.findall(model)
	if m:
		solver = m[-1].lower() # the last solver command
		category = None # We'll take any category
   
	for i in solverMenu:
		solverMenu[i][0].Checked = False
		for j in solverMenu[i][1]:
			if (category==i or category==None) and solver == j.lower():
				solverMenu[i][1][j].Checked = True
				solverMenu[i][0].Checked = True
				SolverStudio.ActiveModelSettings["AMPLNeosSolver"]=j
				SolverStudio.ActiveModelSettings["AMPLNeosCategory"]=i
			else:
				solverMenu[i][1][j].Checked = False
	queueChoice.Checked = SolverStudio.ActiveModelSettings.GetValueOrDefault("ShortQueueLength",False)
	doFixChoice.Checked = SolverStudio.ActiveModelSettings.GetValueOrDefault("FixMinorErrors",True)
	
def emptyClickMenu(key,e):
	return
  
def SolverClickMenu(key,e):
	if key.Checked == False:
		for i in solverMenu:
			solverMenu[i][0].Checked = False
			for j in solverMenu[i][1]:
				if solverMenu[i][1][j] == key:
					key.Checked = True
					solverMenu[i][0].Checked = True
					SolverStudio.ActiveModelSettings["AMPLNeosSolver"]=j
					SolverStudio.ActiveModelSettings["AMPLNeosCategory"]=i
				else:
					solverMenu[i][1][j].Checked = False
		UpdateActiveModelText()


def UpdateActiveModelText():
	model = SolverStudio.GetActiveModelText()
	if model != "No ActiveWorkBookMgr":
		solver=SolverStudio.ActiveModelSettings["AMPLNeosSolver"].lower()
		# r = re.compile(r"^\s*option\s+solver\s+(\w+)\s*;( ## SolverStudio: Line Modified)?",re.IGNORECASE | re.MULTILINE)
		r = re.compile(r"(.*)(^\s*option\s+solver\s+(\w+(-+\w+)*)\s*;( ## SolverStudio: Line Modified)?)",re.IGNORECASE | re.MULTILINE | re.DOTALL)
		if re.search(r,model):
			# The model has a solver specified, so change it (otherwise we do not put the model into the file)
			# model=r.sub(r'option solver '+solver+r"; ## SolverStudio: Line Modified",model)
			# r = re.compile(r"(.*)(^\s*option\s+solver\s+(\w+)\s*;( ## SolverStudio: Line Modified)?)",re.IGNORECASE | re.MULTILINE | re.DOTALL)
			model=r.sub(r'\1option solver '+solver+r'; ## SolverStudio: Line Modified',model)
			SolverStudio.SetActiveModelText(model)
		#else:
		#	r = re.compile(r"^\s*solve\s*;",re.IGNORECASE | re.MULTILINE)
		#	if re.search(r,model):
		#		model=r.sub(r'\r\noption solver '+solver+r"; ## SolverStudio: Line Modified\r\nsolve;",model)
		#	else:
		#		model = model + '\r\noption solver '+solver+'; ## SolverStudio: Line Modified\r\nsolve;'
		# SolverStudio.SetActiveModelText(model)
	return
	
def CheckSolver():
	model = SolverStudio.GetModelText()
	r = re.compile(r"(.*)(^\s*option\s+solver\s+(\w+(-+\w+)*)\s*;( ## SolverStudio: Line Modified)?)",re.IGNORECASE | re.MULTILINE | re.DOTALL)
	m=re.search(r,model)
	return m.groups(0)[2]
	
def QueueClickMenu(key,e):
  queueChoice.Checked=not queueChoice.Checked
  SolverStudio.ActiveModelSettings["ShortQueueLength"]=queueChoice.Checked
  
def doFixClickMenu(key,e):
  doFixChoice.Checked=not doFixChoice.Checked
  SolverStudio.ActiveModelSettings["FixMinorErrors"]=doFixChoice.Checked

