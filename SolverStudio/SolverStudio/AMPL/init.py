import clr
clr.AddReference('System.Windows.Forms')
# clr.AddReference('Microsoft.Office.Tools.Ribbon');
import webbrowser
import os 

#from System.Windows.Forms import Application, Form, Label, ToolStripMenuItem
from System.Windows.Forms import Form, Label, ToolStripMenuItem, MessageBox,MessageBoxButtons,DialogResult, OpenFileDialog

import version
import InstallAMPLCML

import csv
import re
import sys # for exception handling via sys.exc_info()[0]

def writeItem(g,n):
	wroteItem = False
	for i in n:
		if i[1]!="": 
			wroteItem = True
			g.write(i[1]+"\n")
	return wroteItem

def readSet(f,g,line):
	r = re.compile(r"^\s*set:?\s*(\w+):?\s*\:\=\s*(.*)\s*;",re.IGNORECASE | re.MULTILINE)
	r1 = re.compile(r"(\()?(.*?)(?(1)\))",re.IGNORECASE | re.MULTILINE)
	m = re.search(r,line)
	setName = m.group(1)
	print "Found set: " + setName
	dim = None
	size = 0
	if m:
		g.write(setName+"\n")
		n = r1.findall(m.group(2))
		wroteItem = False
		for i in n:
			if i[1]!="": 
				wroteItem = True
				g.write(i[1]+"\n")
				size = size + 1
				if dim == None: dim = i[1].count(",")
		if not wroteItem:
			r1 = re.compile(r'(\')?(.*?)(?(1)\')',re.IGNORECASE | re.MULTILINE)
			n = r1.findall(m.group(2))
			for i in n:
				if i[1]!="": 
					wroteItem = True
					g.write("'" + i[1] + "'\n")
					size = size + 1
					if dim == None: dim = i[1].count(",")
			if not wroteItem:
				list = m.group(2).split()
				for i in list:
					if i!="": 
						g.write(i+"\n")
						size = size + 1
						if dim == None: dim = i.count(",")
		if not setName in mySets:
			mySets[setName] = [dim+1,size]
	g.write("\n")
	
def readParam(f,g,line):
	if " (tr):" in line:
		transpose = True
	else:
		transpose = False
	if ":=" in line:
		r = re.compile(r"^\s*param:?\s+(\w+):?\s+(\(tr\))?:?(\n)?\s+((.*)(\:\=$))?",re.IGNORECASE | re.MULTILINE)
		m = re.search(r,line)
		if m:	
			colIndices = m.group(5)
			tableCols = len(colIndices.split())

		else:
			r=re.compile(r"^\s*param:?\s+(\w+):?\s*\:\=",re.IGNORECASE)
			m=re.search(r,line)
			colIndices = None
			tableCols=0
		paramName = m.group(1)
	else:
		r = re.compile(r"^\s*param:?\s+(\w+):?\s+(\(tr\))?:?",re.IGNORECASE | re.MULTILINE)
		m = re.search(r,line)
		paramName = m.group(1)
		line = f.readline()
		colIndices = line[:-3]
		tableCols = len(line.split())-1
	colPrint=False
	print "Found parameter: " + paramName
	g.write(paramName + "\n")
	dim = None
	size = 0
	while True:
		doBreak = False
		line = f.readline()
		line = line.strip().replace(";"," ; ")
		if not line: break
		if line == " ; ": 
			g.write("\n")
			break
		if ";" in line: 
			doBreak=True
			line = line.replace(";","")
		if line[0]==":":
			print "## There was an adding the parameter: " + paramName
		rs = re.compile(r"\s+")
		line = rs.sub(" ",line)
		# r = re.compile(r"^\s*(\[)?(.*?)(?(1)\])\s+(.*)",re.IGNORECASE | re.MULTILINE)
		r = re.compile(r"^\s*(\[)?(.*?)(?(1)\])\s+(.*)",re.IGNORECASE | re.MULTILINE|re.DOTALL)
		rr = re.compile(r'"(.*?)"([, ])',re.IGNORECASE | re.DOTALL)
		mm = re.search(rr,line)
		if mm: 
			grp1 = mm.group(1).replace("'","''")
			line = rr.sub(" '" + grp1 + "'" + mm.group(2),line)
		m = re.search(r,line)
		if m: 
			if dim==None: dim = m.group(2).count(",")
			if not colPrint and colIndices != None:
				g.write(","*(dim+1)+colIndices+"\n")
				colPrint = True
			g.write(m.group(2).strip() + " " + m.group(3) + "\n")
			size = size+1
		if doBreak: 
			g.write("\n")
			break
	if not paramName in myParams:
		myParams[paramName] = [dim+2,size,tableCols,transpose]



def AddSet(setName,sheetName,row):
	Rows = mySets[setName][1]
	Cols = mySets[setName][0]
	if Rows==0: Rows=1
	Application.Worksheets(sheetName).Names.Add(Name = setName, RefersToR1C1= "='" + sheetName + "'!R" + str(row+1) + "C1:R" + str(row+Rows)  + "C" + str(Cols) )
	print "Added set: " + setName
	return row+Rows+1

def AddParam(paramName,sheetName,row):
	if paramName == "amt2": print myParams[paramName]
	Rows = myParams[paramName][1]
	Cols = myParams[paramName][0]
	tableCols = myParams[paramName][2]
	if Rows==0: Rows=1
	if tableCols!=0:
		if myParams[paramName][3]==True:
			rowIndex = ".columnindex"
			colIndex = ".rowindex"
		else:
			rowIndex = ".rowindex"
			colIndex = ".columnindex"
		Application.Worksheets(sheetName).Names.Add(Name = paramName, RefersToR1C1= "='" + sheetName+"'!R" + str(row+2) + "C" + str(Cols)  + ":R" + str(row+Rows+1)  + "C" + str(Cols + tableCols -1))
		Application.Worksheets(sheetName).Names.Add( Name = paramName + rowIndex, RefersToR1C1 = "='" + sheetName + "'!R" + str(row+2)+"C1:R" + str(row+Rows+1)  + "C" + str(Cols-1))
		Application.Worksheets(sheetName).Names.Add( Name = paramName + colIndex, RefersToR1C1 = "='" + sheetName + "'!R" + str(row+1)+"C"+str(Cols)+":R" + str(row+1)  + "C" + str(Cols + tableCols - 1))
		Application.Worksheets(sheetName).Names.Add( Name = paramName + "rowindex.dirn", RefersToR1C1 = '="Row"')
		Application.Worksheets(sheetName).Names.Add( Name = paramName + "columnindex.dirn", RefersToR1C1 = '="Column"')
		Application.Worksheets(sheetName).Names(paramName + ".rowindex").Visible = False
		Application.Worksheets(sheetName).Names(paramName + ".columnindex").Visible = False		
	elif Cols != 0:
		Application.Worksheets(sheetName).Names.Add(Name = paramName, RefersToR1C1= "='" + sheetName+"'!R" + str(row+1) + "C" + str(Cols)  + ":R" + str(row+Rows)  + "C" + str(Cols))
		Application.Worksheets(sheetName).Names.Add( Name = paramName + ".rowindex", RefersToR1C1 = "='" + sheetName + "'!R" + str(row+1)+"C1:R" + str(row+Rows)  + "C" + str(Cols-1))
		Application.Worksheets(sheetName).Names.Add( Name = paramName + ".rowindex.dirn", RefersToR1C1 = '="Row"')
		Application.Worksheets(sheetName).Names(paramName + ".rowindex").Visible = False
	else:
		Application.Worksheets(sheetName).Names.Add(Name = paramName, RefersToR1C1= "='" + sheetName + "'!R" + str(row+1) + "C1")
	print "Added parameter: " + paramName
	return row + Rows + 1
	
def ImportAMPLData(inPath):
	initialScreenUpdating = Application.ScreenUpdating
	try:
		print
		print "## Starting"
		Application.ScreenUpdating = False
		global mySets
		global myParams
		mySets = {}
		myParams = {}
		r = re.compile(".*(\\\)(.+)\.dat",re.IGNORECASE | re.DOTALL)
		DestSheetName = re.search(r,inPath).groups(1)[1]
		outPath = SolverStudio.WorkingDirectory()+"\\"+DestSheetName+".txt"
		f = open(inPath,"r")
		g = open(outPath,"w")
		line = f.readline().strip()
		g.write("'" + line + "'\n\n")
		while True:
			line = f.readline()
			if not line: break
			line = line.strip()
			items = csv.reader([line], delimiter=' ', skipinitialspace=True, quotechar="'").next()
			if len(items)==0: continue
			items[0] = items[0].replace(":","")
			if items[0].lower() == "set":
				readSet(f,g,line)
				continue
			if items[0].lower() == "param":
				readParam(f,g,line)
				continue
		f.close()
		g.close()
		print "\n## Found " + str(len(mySets)) + " sets."
		print "## Found " + str(len(myParams)) + " parameters.\n"
		r = re.compile(".*(\\\)(.+)\.txt",re.IGNORECASE | re.DOTALL)
		DestSheetName = re.search(r,outPath).groups(1)[1]
		BookName = Application.ActiveWorkBook.Name
		OrigSheetName = Application.ActiveSheet.Name
		Application.Workbooks.OpenText(Filename=outPath, Origin= 3, StartRow=1, DataType=1, TextQualifier=2, ConsecutiveDelimiter=True, Tab=False, Semicolon=False, Comma=True, Space=True, Other=False, TrailingMinusNumbers=True)
		Application.Sheets(DestSheetName).Move( After=Application.Workbooks(BookName).Sheets(OrigSheetName))
		emptyRows=0
		row = 1
		setsAdded = 0
		paramsAdded = 0
		sheetName = Application.ActiveSheet.Name
		while emptyRows<5:
			cellValue = Application.ActiveSheet.Cells(row,1).text
			if cellValue == "": 
				emptyRows = emptyRows+1
				row = row+1
				continue
			emptyRows = 0
			if cellValue in mySets:
				row = AddSet(cellValue,sheetName,row)
				setsAdded = setsAdded + 1
			elif cellValue in myParams:
				row = AddParam(cellValue,sheetName,row)
				paramsAdded = paramsAdded + 1
			row = row+1
		print "\n## Added " + str(setsAdded) + " sets to the data items editor."
		print "## Added " + str(paramsAdded) + " parameters to the data items editor.\n"
	except:
		print "\nERROR: Unable to import AMPL file.\n"
	finally:
		print "## DONE"
		Application.ScreenUpdating = initialScreenUpdating


def VisitAMPLWebPage(key, e):
  webbrowser.open('http://ampl.com')
  
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

def OpenAMPLShell(key,e):
	preferSolverStudioAMPL=SolverStudio.GetRegistrySetting("AMPL","PreferSolverStudioAMPL",1)==1
	try:
		# Get AMPL exe
		exeFile =SolverStudio.GetAMPLPath(preferSolverStudioAMPL=preferSolverStudioAMPL, mustExist=True)
	except:
		raise Exception("AMPL does not appear to be installed. Please install the student version using the AMPL menu in SolverStudio or purchase the commercial version.")
	try:
		oldDirectory = SolverStudio.ChangeToWorkingDirectory()
		# Start AMPL, showing the version info.
		SolverStudio.StartEXE(exeFile, "-v", addExeDirectoryToPath=True, addSolverPath=True, addAtStartOfPath=preferSolverStudioAMPL)
	finally:
		SolverStudio.SetCurrentDirectory(oldDirectory)
  
def About(key,e):
  MessageBox.Show("SolverStudio AMPL Processor:\n"+version.versionString)

def MenuTester2(key,e):
  #form = Form(Text="Hello World Form Menu Tester")
  #label = Label(Text="Hello World!")
  #form.Controls.Add(label)
  #form.ShowDialog()
  pass
  
def InstallAMPLWorker(workManager, workInfo):
  oldDirectory = SolverStudio.ChangeToLanguageDirectory()
  InstallAMPLCML.DoInstall(workManager, workInfo)
  SolverStudio.SetCurrentDirectory(oldDirectory)

def ImportAMPLWorker(key, e):
  with open(SolverStudio.WorkingDirectory()+ "\\sourcePath","r") as tmp:
     inPath = tmp.read()
  ImportAMPLData(inPath)

def ImportMenuHandler(key,e):
	sourcePath = GetAMPLDataPath(key,e)
	if sourcePath == "CANCELLED": return
	with open(SolverStudio.WorkingDirectory() + "\\sourcePath","w") as tmp:
		tmp.write(sourcePath)
	SolverStudio.RunInDialog("About to Import AMPL data file",ImportAMPLWorker)
	
def GetAMPLDataPath(key,e):
	dialog = OpenFileDialog()
	dialog.Filter = "AMPL data files (*.dat)|*.dat"
	myDialog = dialog.ShowDialog()
	if (myDialog == DialogResult.OK):
		return dialog.FileName
	elif (myDialog == DialogResult.Cancel):
		return "CANCELLED"

def InstallAMPLMenuHandler(key,e):
  SolverStudio.RunInDialog("This will download the AMPL Student Edition from AMPL.com, and install it into the SolverStudio directory. It will not replace any other installed AMPL versions you may have.",InstallAMPLWorker)
  UpdateSolverStudioAMPLMenu()

def GetLicenseMgr():
	exeFile =SolverStudio.GetAMPLPath(preferSolverStudioAMPL=False, mustExist=False)
	if exeFile == None:
		MessageBox.Show("AMPL does not appear to be installed on this machine.")
		return None
	AMPLLicenseMgrPath = SolverStudio.GetAssociatedAMPLLicenseMgr(exeFile)
	if AMPLLicenseMgrPath == None:
		MessageBox.Show("An AMPL license server does not appear to be installed on this machine.\nNote that the student version of AMPL does not use a license manager.","SolverStudio")
		return None
	return AMPLLicenseMgrPath

def StartAMPLLicenseServerMenuHandler(key,e):
	AMPLLicenseMgrPath = SolverStudio.GetRunningAMPLLicenseMgrPath()
	if AMPLLicenseMgrPath != None:
		MessageBox.Show("An AMPL license server is already running:\n"+AMPLLicenseMgrPath,"SolverStudio")
		return
	AMPLLicenseMgrPath = GetLicenseMgr()
	if AMPLLicenseMgrPath != None:
		succeeded, error = SolverStudio.StartAMPLLicenseMgr(AMPLLicenseMgrPath,"start")
		if succeeded:
			MessageBox.Show("AMPL License Manager started successfully:\n"+AMPLLicenseMgrPath+"\n\nNote that SolverStudio will normally start the license manager automatically if it is needed.","SolverStudio")
		else:
			MessageBox.Show("The AMPL License Manager failed to start.\n\nLicense Manager:\n"+AMPLLicenseMgrPath+"\n\nError:\n"+error,"SolverStudio")

def StopAMPLLicenseServerMenuHandler(key,e):
	AMPLLicenseMgrPath = SolverStudio.GetRunningAMPLLicenseMgrPath()
	if AMPLLicenseMgrPath == None:
		MessageBox.Show("No running AMPL license server was found.","SolverStudio")
		return
	SolverStudio.StartEXE(AMPLLicenseMgrPath,"stop", createWindow = False)
	MessageBox.Show("Stopping AMPL License Manager:\n"+AMPLLicenseMgrPath,"SolverStudio")

def ShowAMPLLicenseServerStatusMenuHandler(key,e):
	AMPLLicenseMgrPath = SolverStudio.GetRunningAMPLLicenseMgrPath()
	if AMPLLicenseMgrPath != None:
		output = ""
		try:
			output = output+SolverStudio.RunEXE_GetOutput(AMPLLicenseMgrPath,"licshow",60000, useLoopingWaitForExit=True)+"\nampl_lic.exe status:\n"+SolverStudio.RunEXE_GetOutput(AMPLLicenseMgrPath,"status",60000, useLoopingWaitForExit=True)+"\nampl_lic.exe netstatus:\n"+SolverStudio.RunEXE_GetOutput(AMPLLicenseMgrPath,"netstatus",60000, useLoopingWaitForExit=True)+"\nampl_lic.exe ipranges:\n"+ SolverStudio.RunEXE_GetOutput(AMPLLicenseMgrPath,"ipranges",60000, useLoopingWaitForExit=True)
		except Exception as ex:
			output = "Unable to get the AMPL License Manager status.\nError: " + str(ex)
		MessageBox.Show(output,"SolverStudio AMPL License Status")
		return
	AMPLLicenseMgrPath = GetLicenseMgr() # this will show a message box if AMPLLicenseMgrPath==None
	if AMPLLicenseMgrPath != None:
		MessageBox.Show("An AMPL license server was found:\n"+AMPLLicenseMgrPath+"\nbut is not currently running.\n\nPlease start the license manager and try again to get more information.","SolverStudio")
		return

def ShowAMPLLicenseFileMenuHandler(key,e):
	ampl_lic_path = None
	ampl_lic_source = ""
	# First, look for ampl_lic file in the same folder as ampl.exe
	exeFile = SolverStudio.GetAMPLPath(preferSolverStudioAMPL=False, mustExist=False)
	if exeFile != None:
		ampl_lic_path = SolverStudio.GetFileInSameDirectory(exeFile,"ampl.lic", mustExist=False)
		ampl_lic_source = "found in AMPL directory"
	#Next, look for an environment file AMPL_LICFILE that specifies the LICFILE path, 
	# eg AMPL_LICFILE = C:\Ampl-Cplex-Gurobi-2013\ampl.lic
	if ampl_lic_path == None:
		ampl_lic_path = os.getenv("AMPL_LICFILE")
		ampl_lic_source = "specified by %AMPL_LICFILE%"
	# Finally, search the path for ampl.lic	
	if ampl_lic_path == None:
		ampl_lic_path = SolverStudio.FindFileOnSystemPath("ampl.lic")
		ampl_lic_source = "found on system %PATH%"
	if ampl_lic_path == None:
		MessageBox.Show("No ampl.lic file could be found.","SolverStudio")
		return
	try:
		with open(ampl_lic_path,"rb") as inFile:
			licenseText = inFile.read()
	except Exception as ex:
		MessageBox.Show("An error occurred try to read the AMPL license file '"+ampl_lic_path+"' "+ampl_lic_source+"\n\n"+str(ex),"SolverStudio")
		return
	MessageBox.Show("AMPL License File '"+ampl_lic_path+"' "+ampl_lic_source+":\n\n"+licenseText,"SolverStudio")


# Return s as a string suitable for indexing, which means tuples are listed items in []
def Initialise():
  #Menu.DropDownItems.Add( ToolStripMenuItem("Test",None,MenuTester) )
  Menu.ToolStripMenuItem.MouseDown += UpdateFixMinorErrorsMenu
  Menu.Add( "Open AMPL Shell",OpenAMPLShell)
  Menu.Add( "View Last Data File",OpenDataFile)
  Menu.AddSeparator()
  Menu.Add( "Import Data File...",ImportMenuHandler)
  Menu.AddSeparator()
  global menuFixMinorErrors
  menuFixMinorErrors=Menu.Add( "Fix Minor Errors",ClickFixMinorErrorsMenu)
  Menu.AddSeparator()
  Menu.Add( "Open AMPL.com",VisitAMPLWebPage)
  Menu.Add( "Open online AMPL book",VisitAMPLBookWebPage)
  Menu.Add( "Open online AMPL PDF documentation",OpenAMPLpdf)
  Menu.AddSeparator()
  Menu.Add( "Start AMPL License Server",StartAMPLLicenseServerMenuHandler)
  Menu.Add( "Stop AMPL License Server",StopAMPLLicenseServerMenuHandler)
  Menu.Add( "Show AMPL License Server Status",ShowAMPLLicenseServerStatusMenuHandler)
  Menu.Add( "View AMPL License File",ShowAMPLLicenseFileMenuHandler)
  Menu.AddSeparator()
  Menu.Add( "Install AMPL Student Version...",InstallAMPLMenuHandler)
  global menuPreferSolverStudioAMPL
  menuPreferSolverStudioAMPL=Menu.Add( "Use SolverStudio's AMPL Student Version",ClickSolverStudioAMPLMenu)
  UpdateSolverStudioAMPLMenu()
  # Menu.Add( "-",None) This does not work :-(
  Menu.AddSeparator()
  Menu.Add( "About AMPL Processor",About)

def UpdateFixMinorErrorsMenu(key,e):
  menuFixMinorErrors.Checked=SolverStudio.ActiveModelSettings.GetValueOrDefault("FixMinorErrors",True)
  
def ClickFixMinorErrorsMenu(key,e):
  menuFixMinorErrors.Checked=not menuFixMinorErrors.Checked
  SolverStudio.ActiveModelSettings["FixMinorErrors"]=menuFixMinorErrors.Checked

def UpdateSolverStudioAMPLMenu():
  if SolverStudio.GetSolverStudioAMPLPath()==None:
    menuPreferSolverStudioAMPL.Checked = False
    menuPreferSolverStudioAMPL.Enabled = False
  else:
    menuPreferSolverStudioAMPL.Enabled = True
    menuPreferSolverStudioAMPL.Checked = (SolverStudio.GetRegistrySetting("AMPL","PreferSolverStudioAMPL",1) == 1)
  
def ClickSolverStudioAMPLMenu(key ,e):
  menuPreferSolverStudioAMPL.Checked=not menuPreferSolverStudioAMPL.Checked
  SolverStudio.SetRegistrySetting("AMPL","PreferSolverStudioAMPL",int(menuPreferSolverStudioAMPL.Checked))
