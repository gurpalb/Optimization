# This unused file allows the user to import a GDX file using a text file
# It requires a version of GAMS to be installed but does not use the propietary GAMS dll's.

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
# clr.AddReference('Microsoft.Office.Tools.Ribbon');
import webbrowser
import os
import os.path
import subprocess 
from _winreg import *
from System.Windows.Forms import  LinkBehavior,LinkArea,LinkLabel,CheckBox, Form, Label, ToolStripMenuItem, MessageBox, MessageBoxButtons, Button, TextBox, FormBorderStyle, DialogResult, OpenFileDialog #Application,
from System.Drawing import Point,Color


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
		choice = MessageBox.Show("Unable to open last data file. You need a version of GAMS installed to view GDX files. You can always use \"GAMS>Import a GDX file>SheetData.gdx\" to open it in excel.","SolverStudio",MessageBoxButtons.OK)
	
	
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
		with open("NEOSSolverList","r") as in_file:
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
			if solver=="Gurobi": continue
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
		with open("NEOSSolverList","w") as dest_file:
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
	SolverStudio.RunInDialog("This will update the NEOS Solver List for GAMS.",UpdateSolverListWorker,True,True,True)

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
	with open("NEOSCategoryList","r") as in_file:
		categoryDict = eval(in_file.read())
	#choice = MessageBox.Show("Unable to contact the NEOS Server. Please check your internet connection and try again.","SolverStudio",MessageBoxButtons.OK)
	with open("NEOSSolverList","r") as in_file:
		sGoodSolvers = in_file.read()
	SolverStudio.SetCurrentDirectory(oldDirectory)
	return eval(sGoodSolvers), categoryDict
	
def ImportGDXWorker(key,e):
	initialScreenUpdating = Application.ScreenUpdating
	try:
		Application.ScreenUpdating = False
		with open("C:\\Temp\\sourcepath","r") as tmp:
			sourcePath = tmp.read()
		path = SolverStudio.GetGAMSPath()
		varChoice = MessageBox.Show("Would you like to import the variables only?","SolverStudio",MessageBoxButtons.YesNo)
		if varChoice == DialogResult.Yes:
			varOnly = True
		elif varChoice == DialogResult.No:
			varOnly = False
		else:
			return
		groupChoice = MessageBox.Show("Would you like to group the data items for easier viewing?","SolverStudio",MessageBoxButtons.YesNo)
		if groupChoice == DialogResult.Yes:
			group = True
		elif groupChoice == DialogResult.No:
			group = False
		else:
			return
		print  "\n## About to start GDXDUMP"
		path = path.replace("gams.exe","gdxdump")
		index = len(sourcePath)-sourcePath[::-1].index("\\")-1
		GDXName = sourcePath[index+1:len(sourcePath)-4]
		tempTxt=  SolverStudio.WorkingDirectory() + "\\" + GDXName + "_temp.txt"
		path = path.replace("\\","\\\\")
		sourcePath = sourcePath.replace("\\","\\\\")
		tempTxt = tempTxt.replace("\\","\\\\")
		command = "\"" + path + "\" " + "\"" + sourcePath + "\"" + " output = \"" + tempTxt
		if (os.system(command)==0):
			print  "## GDXDUMP completed"
		else:
			print  "## Error: GDXDUMP failed"
			return
		mySets={}
		myParams={}
		myScalars={}
		myVars={}
		csvPath = tempTxt.replace("_temp.txt",".txt")
		f=open(tempTxt,"r")
		g=open(csvPath,"w")
		row=0
		lineCount=0
		start=False
		printLine = False
		ID = ""
		print  "## About to start import"
		f.readline()
		f.readline()
		maxDim=0
		while True:
			if varOnly: dim2 = 0
			lineCount=lineCount+1
			line = f.readline()
			if not line: break
			if line == "": 
				ID = ""
				printLine = False
			if "$offempty" in line: break
			if "Equation " in line: printLine = False
			if ("Set " in line or "Parameter " in line) and not varOnly:
				printLine = True
				start=True
				startRow=lineCount
				if "Set " in line:
					name =line[4:line.index("(")]
					ID="set"
				if "Parameter " in line:
					name = line [10:line.index("(")]
					ID="param"
				g.write(name+"\n")
				dim2=line.count("*")
				if ID=='param': dim2 = dim2 + 1
				# if  "/ /;" in line: g.write("\n")
				continue
			if "Scalar" in line and not varOnly:
				printLine = True
				scalarName = line[7:line.index("/")-1]
				scalarValue = line[line.index("/",1)+2:line.index("/;")-1]
				myScalars[scalarName]= scalarValue
				g.write(scalarName + "\n")
				g.write(scalarValue + "\n")
				# if scalarValue == " ": g.write("\n")
				ID = "scalar"
				continue
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
					# if varValue == "": g.write("\n")
					ID="var1"
				else:
					endIndex = line.index("(",startIndex)
					varName = line[startIndex:endIndex]
					g.write(varName+"\n")
					ID = "var2"
				start = True
				startRow = lineCount
				# if  "/ /;" in line: g.write("\n")
				continue
			line = line.replace("/;","").replace(",","")
			line = line.replace("'.'","','")
			if ID=="var2" and ("'.M " in line or "'.UP " in line or "'.LO " in line): 
				lineCount = lineCount - 1
				continue
			if ID =="var2": line=line.replace("'.L ","' ")
			if start and lineCount==startRow+1:
				dim = line.count("','")+1
				if ID == "param" or ID == "var2": dim = dim +1   
			
			if start and line=="\n":
				if dim2>dim: dim=dim2
				if dim>maxDim: maxDim = dim
				if ID == "set":
					mySets[name]=[lineCount-startRow-1,dim]
				if ID == "param":
					myParams[name]=[lineCount-startRow-1,dim]
				if ID == "var2":
					myVars[varName]=[lineCount-startRow-1,dim]
				if ID == "var1":
					myVars[varName]=[lineCount-startRow-1,0]
				start=False
				if lineCount-startRow - 1 == 0 and ID != "var1": 
					g.write("\n")
			if printLine: g.write(line)
			# if startRow == lineCount + 1 and line == "": g.write("\n")
			if lineCount%10000 == 0: print  "## Reached line " + str( lineCount)
		f.close()
		g.close()
		print  "\n## Found " + str(len(myScalars)) + " scalars\n## Found " +str(len(mySets)) + " sets\n## Found " + str(len(myParams)) + " parameters\n## Found " + str(len(myVars)) + " variables"
		BookName = Application.Activeworkbook.Name
		origSheetName = Application.ActiveSheet.Name
		
		from System import Array
		colArray=Array.CreateInstance(object,maxDim)
		for i in range(0,maxDim):
			a=Array.CreateInstance(object,2)
			a[0]=i+1
			a[1]=2
			colArray[i]=a
		
		Application.Workbooks.OpenText(Filename=csvPath, Origin= 3, StartRow=1, DataType=1, TextQualifier=2, ConsecutiveDelimiter=True, Tab=False, Semicolon=False, Comma=True, Space=True, Other=False, TrailingMinusNumbers=True, FieldInfo = colArray)
		Application.Sheets(GDXName).Move( After=Application.Workbooks(BookName).Sheets(origSheetName))
		sheetName = Application.ActiveSheet.Name
		
		
		#Application.Workbooks(BookName).Worksheets(origSheetName).Select()
		print  "\n## About to add data to the Data Items Editor"
		row = 1
		lastRow = 1
		setsAdded=0
		scalarsAdded=0
		parametersAdded=0
		variablesAdded=0
		emptyCellCount=0
		while emptyCellCount<4:
			name = Application.Worksheets(sheetName).Cells(row,1).text
			if name == "":
				emptyCellCount=emptyCellCount+1
				row=row+1
				continue
			else:
				emptyCellCount=0
			newRows=0	
			if name in myScalars: newRows = 1
			if name in myParams: newRows = myParams[name][0]
			if name in mySets: newRows = mySets[name][0]
			if name in myVars: newRows = myVars[name][0]
			if row+newRows >1000000:
				Application.Worksheets(sheetName).Range(Application.Cells(row,1),Application.Cells(1048576,16348)).Clear()
				lastRow = lastRow + row - 1
				Application.Workbooks.OpenText(Filename=csvPath, Origin= 3, StartRow=lastRow, DataType=1, TextQualifier=2, ConsecutiveDelimiter=True, Tab=False, Semicolon=False, Comma=True, Space=True, Other=False, TrailingMinusNumbers=True)
				Application.Sheets(GDXName).Move( After=Application.Workbooks(BookName).Sheets(origSheetName))
				sheetName = Application.ActiveSheet.Name
				#Application.Workbooks(BookName).Worksheets(origSheetName).Select()
				row = 1
				continue
				
			if name in myScalars:
				value=myScalars[name]    
				Application.Worksheets(sheetName).Names.Add(Name = name, RefersToR1C1= "='" + sheetName + "'!R" + str(row+1) + "C1")
				if group:
					Application.Worksheets(sheetName).Rows(str(row+1)).Select()
					Application.Selection.Rows.Group()
				row=row+2
				scalarsAdded=scalarsAdded+1
				continue
			if name in mySets:
				Rows = mySets[name][0]
				Cols = mySets[name][1]
				if Rows==0: Rows=1
				Application.Worksheets(sheetName).Names.Add(Name = name, RefersToR1C1= "='" + sheetName + "'!R" + str(row+1) + "C1:R" + str(row+Rows)  + "C" + str(Cols) )
				if group:
					Application.Worksheets(sheetName).Rows(str(row+1) + ":" + str(row+Rows)).Select()
					Application.Selection.Rows.Group()
				row = row+Rows+1
				setsAdded = setsAdded +1
				continue
			if name in myParams or name in myVars:
				if name in myParams:
					Rows = myParams[name][0]
					Cols = myParams[name][1]
					parametersAdded = parametersAdded +1
				else:
					Rows = myVars[name][0]
					Cols = myVars[name][1]
					variablesAdded = variablesAdded +1
				if Rows==0: Rows=1
				if Cols != 0:
					Application.Worksheets(sheetName).Names.Add(Name = name+"_index", RefersToR1C1= "='" + sheetName + "'!R" + str(row+1)+"C1:R" + str(row+Rows)  + "C" + str(Cols-1))
					Application.Worksheets(sheetName).Names.Add(Name = name, RefersToR1C1= "='" + sheetName+"'!R" + str(row+1) + "C" + str(Cols)  + ":R" + str(row+Rows)  + "C" + str(Cols))
					Application.Worksheets(sheetName).Names.Add( Name = name + ".rowindex", RefersToR1C1 = "='" + sheetName + "'!" + name + "_index")
					#
					Application.Worksheets(sheetName).Names.Add( Name = name + ".rowindex.dirn", RefersToR1C1 = '="Row"')
					#
					Application.Worksheets(sheetName).Names(name + ".rowindex").Visible = False
					Application.Worksheets(sheetName).Names(name + "_index").Visible = False
				else:
					Application.Worksheets(sheetName).Names.Add(Name = name, RefersToR1C1= "='" + sheetName + "'!R" + str(row+1) + "C1")
				if group:
					Application.Worksheets(sheetName).Rows(str(row+1) + ":" + str(row+Rows)).Select()
					Application.Selection.Rows.Group()
				row=row+Rows+1
				continue
			row=row+1
		if group:
			Application.Worksheets(sheetName).Outline.ShowLevels(RowLevels=1)
			Application.Workbooks(BookName).Worksheets(sheetName).cells(1,1).select()
		Application.ScreenUpdating = True
		print  "\n## Added " + str(scalarsAdded) + " scalars\n## Added " + str(setsAdded) + " sets\n## Added " + str(parametersAdded) + " parameters\n## Added " + str(variablesAdded) + " variables\n## Finished importing GDX file"
	except:
		MessageBox.Show("Unable to import GDX file.","SolverStudio",MessageBoxButtons.OK)
	finally:
		Application.ScreenUpdating = initialScreenUpdating
	return
		
def ImportGDXMenuHandler(key,e):
	sourcePath = GetGDXPath(key,e)
	if sourcePath == "CANCELLED": return
	with open("C:\\Temp\\sourcePath","w") as tmp:
		tmp.write(sourcePath)
	SolverStudio.RunInDialog("Importing large GDX files can take some time. Are you sure you want to start?\n",ImportGDXWorker)
	
def GetGDXPath(sender,event):
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
	NeosWarning()
	DLLPath = CheckDLL()
	try:
		if DLLPath[0]==";": DLLPath=DLLPath[1:]
		if not os.path.exists(DLLPath + "\\gmszlib1.dll"):
			choice = MessageBox.Show("The GAMS compression library not found. Check that gmszlib1.dll still exists in the GAMSDLLs folder included in SolverStudio. If not you will need to download a new copy of SolverStudio.","SolverStudio",MessageBoxButtons.OK)
	except:
		choice = MessageBox.Show("The GAMS compression library not found. Check that gmszlib1.dll still exists in the GAMSDLLs folder included in SolverStudio. If not you will need to download a new copy of SolverStudio.","SolverStudio",MessageBoxButtons.OK)
		
	# Menu.Add( "View last text data file",OpenDataFile)
	Menu.Add( "View last GAMS input data file in GAMSIDE",OpenInputDataInGamsIDE)
	Menu.AddSeparator()
	Menu.Add( "Import a GDX file",ImportGDXMenuHandler)
	#Menu.Add( "Create a GDX file using the Model Data",SaveGDXMenuHandler)
	Menu.AddSeparator()
	global doFixChoice
	doFixChoice=Menu.Add( "Fix Minor Errors",FixErrorClickMenu)
	global queueChoice
	queueChoice=Menu.Add( "Run in short queue",QueueClickMenu)
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
		choice = MessageBox.Show("A NEOS category is no longer supported by GAMSonNEOS.\nLaunch NEOS solver list updater.","SolverStudio",MessageBoxButtons.OK)
		if choice == DialogResult.OK:
			SolverStudio.RunInDialog("This will update the NEOS Solver List for GAMS.",UpdateSolverListWorker,True,True,True)
	Menu.AddSeparator()
	Menu.Add( "Open GAMS web page",OpenGAMSWebSite) 
	Menu.Add( "Open GAMS online documentation",OpenGAMSOnlineDocumentation) 
	Menu.Add( "Open GAMS download web page",OpenGAMSDownload) 
	Menu.AddSeparator()
	Menu.Add( "Update NEOS Solvers",UpdateSolverListMenuHandler)
	Menu.Add( "Open the NEOS web page",OpenNEOSWebSite)

  #Menu.Add( "Open model in GAMS IDE",OpenModelInGamsIDE)
  #Menu.Add( "Open README file",OpenREADME) 
  #Menu.Add( "Open COPYING file",OpenCOPYING)
  
def Update(key,e):
	solver = SolverStudio.ActiveModelSettings.GetValueOrDefault("GAMSNeosSolver",'Cbc')
	category = SolverStudio.ActiveModelSettings.GetValueOrDefault("GAMSNeosCategory",'milp')
	for i in solverMenu:
		solverMenu[i][0].Checked = False
		for j in solverMenu[i][1]:
			if category == i and solver == j:
				solverMenu[i][1][j].Checked = True
				solverMenu[i][1][j].Checked = True
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
  
def CheckDLL():
	sysPath = os.environ["PATH"]
	DLLDir = SolverStudio.GetGAMSDLLPath()
	if not DLLDir in sysPath:
		if not sysPath[-1]==";": DLLDir = ";" + DLLDir
		os.environ["PATH"] = sysPath + DLLDir
	return DLLDir

def NeosWarning():
	try:
		keyVal = 'Software\\SolverStudio\\NEOS'
		key = OpenKey(HKEY_CURRENT_USER, keyVal, 0, KEY_ALL_ACCESS)
		bool = QueryValueEx(key,"ShowNeosWarning")
		if (bool[0]!="False"):
			myForm = WarningForm()
			myForm.ShowDialog()
		CloseKey(key)
	except: 
		myForm = WarningForm()
		myForm.ShowDialog()

class WarningForm(Form):
	def __init__(self):
		self.Text = "SolverStudio"
		self.FormBorderStyle = FormBorderStyle.FixedDialog    
		self.Height=150
		self.Width = 400
		self.Warning = Label()
		self.Warning.Text = "Note: Models solved using via NEOS will be available in the public domain for perpetuity."
		self.Warning.Location = Point(10,10)
		self.Warning.Height  = 30
		self.Warning.Width = 380
		self.myOK=Button()
		self.myOK.Text = "OK"
		self.myOK.Location = Point(310,90)
		self.myOK.Click += self.AddChoice
	
		self.check1 = CheckBox()
		self.check1.Text = "Don't show this message again"
		self.check1.Location = Point(10, 80)
		self.check1.Width = 250
		self.check1.Height = 50
		
		self.link1 = LinkLabel()
		self.link1.Location = Point(10, 50)
		self.link1.Width = 380
		self.link1.Height = 40
		self.link1.LinkClicked += self.OpenTCs
		self.link1.VisitedLinkColor = Color.Blue;
		self.link1.LinkBehavior = LinkBehavior.HoverUnderline;
		self.link1.LinkColor = Color.Navy;
		self.link1.Text = "By using NEOS via SolverStudio you also agree to the NEOS-Server terms and conditions."
		self.link1.LinkArea = LinkArea(65,20)
		
		self.AcceptButton = self.myOK
		
		self.Controls.Add(self.myOK)		
		self.Controls.Add(self.link1)		
		self.Controls.Add(self.Warning)	
		self.Controls.Add(self.check1)

		self.CenterToScreen()
	
	def OpenTCs(self,sender,event):
		import webbrowser
		webbrowser.open("""http://www.neos-server.org/neos/termofuse.html""")
	
	def AddChoice(self,sender,event):
		keyVal = 'Software\\SolverStudio\\NEOS'
		try:
			key = OpenKey(HKEY_CURRENT_USER, keyVal, 0, KEY_ALL_ACCESS)
		except:
			key = CreateKey(HKEY_CURRENT_USER, keyVal)
		SetValueEx(key, "ShowNeosWarning", 0, REG_SZ, str(not self.check1.Checked))
		CloseKey(key)	
		self.Close()