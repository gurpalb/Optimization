import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
# clr.AddReference('Microsoft.Office.Tools.Ribbon');
import webbrowser
import os
import os.path
# import subprocess 
from shutil import copy2
import version
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
		gamsIDEPath=SolverStudio.GetGAMSIDEPath()
		args="\""+SolverStudio.WorkingDirectory()+"\\SheetData.gdx\""
		SolverStudio.StartEXE(gamsIDEPath, args)
		SolverStudio.SetCurrentDirectory(oldDirectory)
	except:
		choice = MessageBox.Show("Unable to open last data file. You need a version of GAMS installed to view GDX files. You can always use \"GAMS>Import a GDX file>SheetData.gdx\" to open it in Excel.","SolverStudio",MessageBoxButtons.OK)
  
def OpenGAMSResultFile(key,e):
  oldDirectory = SolverStudio.ChangeToWorkingDirectory()
  if os.path.exists( "model.lst" ):
       SolverStudio.ShowTextFile("model.lst")
  SolverStudio.SetCurrentDirectory(oldDirectory)
		
def ImportGDXMenuHandler(key,e):
	oldDirectory = SolverStudio.ChangeToWorkingDirectory()
	#if SolverStudio.Is64Bit():
	#	DLLpath=SolverStudio.GetGAMSDLLPath() + "\\gdxdclib64.dll"
	#else:
	#	DLLpath=SolverStudio.GetGAMSDLLPath() + "\\gdxdclib.dll"
	#if not os.path.exists(SolverStudio.WorkingDirectory() + "\\gdxdclib.dll"):
	#	copy2(DLLpath,SolverStudio.WorkingDirectory() + "\\gdxdclib.dll")
	gdxPath = GetGDXPath(key,e)
	if gdxPath == "CANCELLED": return
	initialScreenUpdating = Application.ScreenUpdating
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
	Application.ScreenUpdating = initialScreenUpdating
	SolverStudio.SetCurrentDirectory(oldDirectory)
	return
		
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
  Menu.Add( "View last GAMS input data file in GAMSIDE...",OpenInputDataInGamsIDE)
  Menu.AddSeparator()
  Menu.Add( "Import a GDX file...",ImportGDXMenuHandler)
  Menu.AddSeparator()
  global doFixChoice
  doFixChoice=Menu.Add( "Fix Minor Errors",FixErrorClickMenu)
  Menu.AddSeparator()
  Menu.Add( "Open GAMS web page",OpenGAMSWebSite) 
  Menu.Add( "Open GAMS online documentation",OpenGAMSOnlineDocumentation) 
  Menu.Add( "Open GAMS download web page",OpenGAMSDownload) 
  Menu.AddSeparator()
  Menu.Add( "About GAMS Processor",About)
  #Menu.Add( "Open model in GAMS IDE",OpenModelInGamsIDE)
  #Menu.Add( "Open README file",OpenREADME) 
  #Menu.Add( "Open COPYING file",OpenCOPYING)
def About(key,e):
	MessageBox.Show("SolverStudio GAMS Processor:\n"+version.versionString)  
def Update(key,e):
	doFixChoice.Checked=SolverStudio.ActiveModelSettings.GetValueOrDefault("FixMinorErrors",True)
	
def FixErrorClickMenu(key,e):
  doFixChoice.Checked=not doFixChoice.Checked
  SolverStudio.ActiveModelSettings["FixMinorErrors"]=doFixChoice.Checked