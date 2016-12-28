import clr
clr.AddReference('System.Windows.Forms')
# clr.AddReference('System.Drawing')
# clr.AddReference('Microsoft.Office.Tools.Ribbon');
import webbrowser
import os
import os.path
import subprocess 
import version
from _winreg import *
from System.Windows.Forms import  OpenFileDialog, Form, ToolStripMenuItem, MessageBox, MessageBoxButtons, DialogResult, Cursor, Cursors, Button, TextBox, Label, FormBorderStyle, LinkBehavior,LinkArea,LinkLabel

clr.AddReference('System.Drawing')
from System.Drawing import Point,Color

#Application, LinkBehavior,LinkArea,LinkLabel,CheckBox, Form, Label, ToolStripMenuItem, MessageBox, MessageBoxButtons, Button, TextBox, FormBorderStyle, DialogResult, 
# from System.Drawing import Point,Color
from shutil import copy2

def OpenGAMSWebSite(key,e):
  webbrowser.open("""http://www.gams.com/""")
  
def OpenNEOSWebSite(key,e):
  webbrowser.open("""http://www.neos-server.org/""")
  
def OpenGAMSOnlineDocumentation(key,e):
  webbrowser.open("""http://www.gams.com/docs/document.htm""")

def OpenGAMSDownload(key,e):
  webbrowser.open("""http://www.gams.com/download/""")

# def OpenModelInGamsIDE(key,e):
  # oldDirectory = SolverStudio.ChangeToLanguageDirectory()
  # gamsIDEPath=GetEXEPath(exeName="gameside.exe", exeDirectory="gams*", mustExist=true)
  # subprocess.call([gamsIDEPath, SolverStudio.ModelFileName]) 
  # SolverStudio.SetCurrentDirectory(oldDirectory)
  
def OpenInputDataInGamsIDE(key,e):
	try:
		oldDirectory = SolverStudio.ChangeToLanguageDirectory()
		gamsIDEPath=SolverStudio.GetEXEPath(exeName="gamside.exe", exeDirectory="gams*", mustExist=True)
		subprocess.call([gamsIDEPath, SolverStudio.WorkingDirectory()+"\\SheetData.gdx"]) 
		SolverStudio.SetCurrentDirectory(oldDirectory)
	except:
		choice = MessageBox.Show("Unable to open last data file. You need a version of GAMS installed to view GDX files. You can instead use the menu item \"GAMS > Import a GDX file > SheetData.gdx\" to open it in Excel.","SolverStudio",MessageBoxButtons.OK)
	
	
def OpenDataFile(key,e):
	# We need to get access to the temporary directory in our SolverStudio object...
	oldDirectory = SolverStudio.ChangeToWorkingDirectory()
	workingDir = SolverStudio.WorkingDirectory()
	dataFileName = "SheetData.dat"
	dataFilePath = os.path.realpath(os.path.normpath(os.path.join(workingDir, dataFileName)))
	if os.path.exists( dataFilePath ):
		SolverStudio.ShowTextFile(dataFilePath) # This will show the file and update as the file is changed (or created)
	SolverStudio.SetCurrentDirectory(oldDirectory)

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
#======================================================================
	gamsSolvers = {}
	for i in listOfSolvers:
		if "GAMS" in i:
			colIndex=i.index(":")
			category = i[:colIndex]
			solver = i[colIndex+1:-5]
			# if solver=="Gurobi": continue
			if category in gamsSolvers:
				gamsSolvers[category] = gamsSolvers[category] + [solver]
			else:
				gamsSolvers[category] = [solver]
	print "## Checking SolverStudio compatibility"
	template = neos.getSolverTemplate('milp','Cbc','GAMS')
	template = template[template.index("</inputMethod>"):]
	newSolvers={}
	for i in gamsSolvers:
		for j in gamsSolvers[i]:
			testXML = neos.getSolverTemplate(i,j,'GAMS')
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
		print "## GAMS Solver list is already up to date."
	SolverStudio.SetCurrentDirectory(oldDirectory)
	print
	print "## DONE."
	print "## Restart SolverStudio to use updated solvers."
	##=============================================================	

def UpdateSolverListMenuHandler(key,e):
	SolverStudio.RunInDialog("Updating the NEOS Solver List for GAMS...",UpdateSolverListWorker,True,True,True)

def ChooseSolver():
	oldDirectory = SolverStudio.ChangeToLanguageDirectory()
	# try:
		# import xmlrpclib
		# NEOS_HOST="neos-server.org"
		# NEOS_PORT=3332
		# neos=xmlrpclib.ServerProxy("http://%s:%d" % (NEOS_HOST, NEOS_PORT))
		# categoryDict = neos.listCategories()
	# except:
		# with open("NEOSCategoryList","r") as in_file:
			# categoryDict = eval(in_file.read())		
	with open("NEOSCategoryList.txt","r") as in_file:
		categoryDict = eval(in_file.read())
	#choice = MessageBox.Show("Unable to contact the NEOS Server. Please check your internet connection and try again.","SolverStudio",MessageBoxButtons.OK)
	with open("NEOSSolverList.txt","r") as in_file:
		sGoodSolvers = in_file.read()
	SolverStudio.SetCurrentDirectory(oldDirectory)
	return eval(sGoodSolvers), categoryDict

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

def ImportGDXMenuHandler(key,e):
	# oldDirectory = SolverStudio.ChangeToWorkingDirectory()
	#if SolverStudio.Is64Bit():
	#	DLLpath=SolverStudio.GetGAMSDLLPath() + "\\gdxdclib64.dll"
	#else:
	#	DLLpath=SolverStudio.GetGAMSDLLPath() + "\\gdxdclib.dll"
	#if not os.path.exists(SolverStudio.WorkingDirectory() + "\\gdxdclib.dll"):
	#	copy2(DLLpath,SolverStudio.WorkingDirectory() + "\\gdxdclib.dll")
	gdxPath = GetGDXPath(key,e)
	if gdxPath == "CANCELLED": return
	initialScreenUpdating = Application.ScreenUpdating
	try:
		Application.ScreenUpdating = False
		varChoice = MessageBox.Show("Would you like to import the variables only?","SolverStudio",MessageBoxButtons.YesNo)
		if varChoice  == DialogResult.Yes:
			varOnly = True
		else:
			varOnly = False
		groupChoice = MessageBox.Show("Would you like to group the data items for easier viewing?","SolverStudio",MessageBoxButtons.YesNo)
		if groupChoice  == DialogResult.Yes:	
			group = True
		else:
			group = False
		PGX = SolverStudio.OpenGDX(gdxPath)
		SolverStudio.ImportGDX(PGX,gdxPath,varOnly,group)
		SolverStudio.CloseGDX(PGX)
	finally:
		Application.ScreenUpdating = initialScreenUpdating
		SolverStudio.SetCurrentDirectory(oldDirectory)
	return
	
def GetGDXPath(sender,event):
	# Ask the user for the path to a GDX file
	dialog = OpenFileDialog()
	dialog.Filter = "GDX files (*.gdx)|*.gdx"
	dialog.InitialDirectory = SolverStudio.WorkingDirectory()
	myDialog = dialog.ShowDialog()
	if (myDialog == DialogResult.OK):
		return dialog.FileName
	elif (myDialog == DialogResult.Cancel):
		return "CANCELLED"

# Return s as a string suitable for indexing, which means tuples are listed items in []
def Initialise():
	#Menu.DropDownItems.Add( ToolStripMenuItem("Test",None,MenuTester) )
	Menu.ToolStripMenuItem.MouseDown += Update
	# We now have DLL's available via the SolverStudio PATH env variable, and so don't need this check
	# DLLPath = CheckDLL()
	#try:
	#	if DLLPath[0]==";": DLLPath=DLLPath[1:]
	#	if not os.path.exists(DLLPath + "\\gmszlib1.dll"):
	#		choice = MessageBox.Show("The GAMS compression library not found. Check that gmszlib1.dll still exists in the GAMSDLLs folder included in SolverStudio. If not you will need to download a new copy of SolverStudio.","SolverStudio",MessageBoxButtons.OK)
	#except:
	#	choice = MessageBox.Show("The GAMS compression library not found. Check that gmszlib1.dll still exists in the GAMSDLLs folder included in SolverStudio. If not you will need to download a new copy of SolverStudio.","SolverStudio",MessageBoxButtons.OK)
		
	# Menu.Add( "View last text data file",OpenDataFile)
	Menu.Add( "View last GAMS input data file in GAMSIDE...",OpenInputDataInGamsIDE)
	Menu.AddSeparator()
	Menu.Add( "Import a GDX file...",ImportGDXMenuHandler)
	#Menu.Add( "Create a GDX file using the Model Data",SaveGDXMenuHandler)
	Menu.AddSeparator()
	global doFixChoice
	doFixChoice=Menu.Add( "Fix Minor Errors",FixErrorClickMenu)
	Menu.AddSeparator()
	global queueChoice
	queueChoice=Menu.Add( "Run in short NEOS queue",QueueClickMenu)
	Menu.Add( "Set Email Address for NEOS...",SetNEOSEmail)
	global solverMenu
	chooseSolverMenu = Menu.Add ("Choose Solver",emptyClickMenu)
	gamsSolvers, categoryDict = ChooseSolver()
	solverMenu = {}
	try:
		for i in gamsSolvers:
			solverMenu[i]=[]
			categoryMenu = ToolStripMenuItem(categoryDict[i],None,emptyClickMenu)
			solverMenu[i].append(categoryMenu)
			chooseSolverMenu.DropDownItems.Add(categoryMenu)
			solverMenu[i].append({})
			for j in gamsSolvers[i]:
				solverMenu[i][1][j] = ToolStripMenuItem(j,None,SolverClickMenu)
				categoryMenu.DropDownItems.Add(solverMenu[i][1][j])
	except:
		choice = MessageBox.Show("A NEOS category is no longer supported by GAMSonNEOS.\n\nDo you wish to update the NEOS solver list now?","SolverStudio",MessageBoxButtons.OK)
		if choice == DialogResult.OK:
			SolverStudio.RunInDialog("NEOS Solver List Update for GAMS.",UpdateSolverListWorker,True,True,True)
	Menu.Add( "Update NEOS Solvers...",UpdateSolverListMenuHandler)
	Menu.Add( "Test NEOS Connection...",TestNEOSConnection)
	Menu.AddSeparator()
	Menu.Add( "Open GAMS web page",OpenGAMSWebSite) 
	Menu.Add( "Open GAMS online documentation",OpenGAMSOnlineDocumentation) 
	Menu.Add( "Open GAMS download web page",OpenGAMSDownload) 
	Menu.AddSeparator()
	Menu.Add( "Open the NEOS web page",OpenNEOSWebSite)
	Menu.AddSeparator()
	Menu.Add( "About GAMSNEOS Processor",About)
  #Menu.Add( "Open model in GAMS IDE",OpenModelInGamsIDE)
  #Menu.Add( "Open README file",OpenREADME) 
  #Menu.Add( "Open COPYING file",OpenCOPYING)

def About(key,e):
	MessageBox.Show("SolverStudio GAMSNEOS Processor:\n"+version.versionString)
	
def Update(key,e):
	solver = SolverStudio.ActiveModelSettings.GetValueOrDefault("GAMSNeosSolver",'Cbc')
	category = SolverStudio.ActiveModelSettings.GetValueOrDefault("GAMSNeosCategory",'milp')
	for i in solverMenu:
		solverMenu[i][0].Checked = False
		for j in solverMenu[i][1]:
			if category == i and solver == j:
				solverMenu[i][1][j].Checked = True
				solverMenu[i][0].Checked = True
				SolverStudio.ActiveModelSettings["GAMSNeosSolver"]=j
				SolverStudio.ActiveModelSettings["GAMSNeosCategory"]=i
			else:
				solverMenu[i][1][j].Checked = False
	queueChoice.Checked = SolverStudio.ActiveModelSettings.GetValueOrDefault("ShortQueueLength",False)
	doFixChoice.Checked=SolverStudio.ActiveModelSettings.GetValueOrDefault("FixMinorErrors",True)
	
def emptyClickMenu(key,e):
	return
  
def SolverClickMenu(key,e):
	for i in solverMenu:
		solverMenu[i][0].Checked = False
		for j in solverMenu[i][1]:
			if solverMenu[i][1][j] == key:
				key.Checked = True
				solverMenu[i][0].Checked = True
				SolverStudio.ActiveModelSettings["GAMSNeosSolver"]=j
				SolverStudio.ActiveModelSettings["GAMSNeosCategory"]=i
			else:
				solverMenu[i][1][j].Checked = False
				
def QueueClickMenu(key,e):
  queueChoice.Checked=not queueChoice.Checked
  SolverStudio.ActiveModelSettings["ShortQueueLength"]=queueChoice.Checked
  
def FixErrorClickMenu(key,e):
  doFixChoice.Checked=not doFixChoice.Checked
  SolverStudio.ActiveModelSettings["FixMinorErrors"]=doFixChoice.Checked
  # queueChoice.Checked=(SolverStudio.ActiveModelSettings["ShortQueueLength"] == True)
  
#def CheckDLL():
#	sysPath = os.environ["PATH"]
#	DLLDir = SolverStudio.GetGAMSDLLPath()
#	if not DLLDir in sysPath:
#		if not sysPath[-1]==";": DLLDir = ";" + DLLDir
#		os.environ["PATH"] = sysPath + DLLDir
#	return DLLDir

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
