import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
# clr.AddReference('Microsoft.Office.Tools.Ribbon');
import webbrowser
import os
import os.path
import subprocess 

from System.Windows.Forms import  Form, Label, ToolStripMenuItem, MessageBox, MessageBoxButtons, Button, TextBox, FormBorderStyle, DialogResult, OpenFileDialog #Application,
from System.Drawing import Point

def OpenGAMSWebSite(key,e):
	import webbrowser
	webbrowser.open("""http://www.gams.com/""")
  
def OpenGAMSOnlineDocumentation(key,e):
	import webbrowser
	webbrowser.open("""http://www.gams.com/docs/document.htm""")

def OpenGAMSDownload(key,e):
	import webbrowser
	webbrowser.open("""http://www.gams.com/download/""")

# def OpenModelInGamsIDE(key,e):
	# import subprocess 
	# oldDirectory = SolverStudio.ChangeToLanguageDirectory()
	# gamsIDEPath=SolverStudio.GetEXEPath(exeName="gameside.exe", exeDirectory="gams*", mustExist=True)
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
  
def OpenGAMSResultFile(key,e):
  oldDirectory = SolverStudio.ChangeToWorkingDirectory()
  if os.path.exists( "model.lst" ):
       SolverStudio.ShowTextFile("model.lst")
  SolverStudio.SetCurrentDirectory(oldDirectory)

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
				Application.Workbooks.OpenText(Filename=csvPath, Origin= 3, StartRow=lastRow, DataType=1, TextQualifier=2, ConsecutiveDelimiter=True, Tab=False, Semicolon=False, Comma=True, Space=True, Other=False, TrailingMinusNumbers=True, FieldInfo = colArray)
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
	myDialog = dialog.ShowDialog()
	if (myDialog == DialogResult.OK):
		return dialog.FileName
	elif (myDialog == DialogResult.Cancel):
		return "CANCELLED"

def Initialise():
  Menu.ToolStripMenuItem.MouseDown += Update
  #Menu.DropDownItems.Add( ToolStripMenuItem("Test",None,MenuTester) )
  Menu.Add( "View last GAMS result file",OpenGAMSResultFile) 
  Menu.Add( "View last GAMS input data file in GAMSIDE",OpenInputDataInGamsIDE)
  Menu.AddSeparator()
  global doFixChoice
  doFixChoice=Menu.Add( "Fix Minor Errors",FixErrorClickMenu)
  Menu.Add( "Import a GDX file",ImportGDXMenuHandler)
  Menu.AddSeparator()
  Menu.Add( "Open GAMS web page",OpenGAMSWebSite) 
  Menu.Add( "Open GAMS online documentation",OpenGAMSOnlineDocumentation) 
  Menu.Add( "Open GAMS download web page",OpenGAMSDownload) 

  #Menu.Add( "Open model in GAMS IDE",OpenModelInGamsIDE)
  #Menu.Add( "Open README file",OpenREADME) 
  #Menu.Add( "Open COPYING file",OpenCOPYING)
  
def Update(key,e):
	doFixChoice.Checked=SolverStudio.ActiveModelSettings.GetValueOrDefault("FixMinorErrors",True)
	
def FixErrorClickMenu(key,e):
  doFixChoice.Checked=not doFixChoice.Checked
  SolverStudio.ActiveModelSettings["FixMinorErrors"]=doFixChoice.Checked