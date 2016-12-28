import clr
clr.AddReference('System.Windows.Forms')
# clr.AddReference('Microsoft.Office.Tools.Ribbon');
import webbrowser
import os 
import version
clr.AddReference('System.Drawing')
from System.Windows.Forms import LinkBehavior,LinkArea,LinkLabel,FormBorderStyle,Form, Label,Button,MessageBox, MessageBoxButtons, DialogResult, ToolStripMenuItem
from System.Drawing import Point,Color

def VisitPYOMOWebPage(key, e):
  webbrowser.open('http://www.pyomo.org')
  
def VisitPYOMODownloadPage(key, e):
  webbrowser.open('http://www.pyomo.org/installation/')

def VisitPYOMODocumentationPage(key, e):
  webbrowser.open('http://www.pyomo.org/documentation/')

def OpenDataFile(key,e):
  # We need to get access to the temporary directory in our SolverStudio object...
  workingDir = SolverStudio.WorkingDirectory()
  dataFileName = "SheetData.dat"
  dataFilePath = os.path.realpath(os.path.normpath(os.path.join(workingDir, dataFileName)))
  SolverStudio.ShowTextFile(dataFilePath) # This will show the file and update as the file is changed (or created)

# Now unused  
#def CheckSolversPath():
#	solversDir = SolverStudio.GetSolversPath()
#	sysPath=os.environ["PATH"]
#	if not solversDir in sysPath:
#		if sysPath[-1] != ";": solversDir=";"+solversDir
#		os.environ["PATH"]=sysPath + solversDir
	
# Return s as a string suitable for indexing, which means tuples are listed items in []
def Initialise():
	Menu.ToolStripMenuItem.MouseDown += UpdateMenu
	# CheckSolversPath()
	#Menu.DropDownItems.Add( ToolStripMenuItem("Test",None,MenuTester) )
	Menu.Add( "View Last Data File",OpenDataFile)
	Menu.AddSeparator()
	global solverMenu
	solverMenu={}
	# This does a check of solvers in the SOlvers directory, and also adds all popular solvers not in the list.
	chooseSolverMenu = Menu.Add ("Choose Solver",emptyClickMenu)
	pyomoSolvers = []
	for path in SolverStudio.GetSolversPaths():
		pyomoSolvers = pyomoSolvers + os.listdir(path)
	# pyomoSolvers2 = ["Bonmin","CBC","CPlex","Couenne","GLPK", "Gurobi","IPOpt","LPSolve"]
	pyomoSolvers2 = ["Bonmin","CPlex","Couenne","GLPK", "Gurobi","IPOpt"]
	pyomoSolvers2lc = [x.lower() for x in pyomoSolvers2]
	for i in pyomoSolvers:
		if i[-4:].lower()!=".exe" and i[-4:].lower()!=".bat": continue # only show .exe and .bat files
		i=i.replace(".exe","").strip() # Keep the case for showing in the menu
		if not i.lower() in pyomoSolvers2lc:
			pyomoSolvers2.append(i)
	pyomoSolvers3 = sorted(pyomoSolvers2, key=str.lower)
	for i in pyomoSolvers3:
		solver = ToolStripMenuItem(i,None,SolverClickMenu)
		solverMenu[i]=solver
		chooseSolverMenu.DropDownItems.Add(solver)
	Menu.AddSeparator()	
	Menu.Add( "Open Pyomo web site",VisitPYOMOWebPage)
	Menu.Add( "Open Pyomo online documentation",VisitPYOMODocumentationPage)
	Menu.Add( "Open Pyomo install page",VisitPYOMODownloadPage)
	Menu.AddSeparator()
	Menu.Add( "About Pyomo Processor",About)
	if SolverStudio.GetPYOMOPath(False) == None:
		myForm = DownloadForm()
		myForm.ShowDialog()	

def About(key,e):
	MessageBox.Show("SolverStudio Pyomo Processor:\n"+version.versionString)
	
class DownloadForm(Form):
	def __init__(self):
		self.Text = "SolverStudio"
		self.FormBorderStyle = FormBorderStyle.FixedDialog	 
		self.Height=150
		self.Width = 400
		self.label1 = Label()
		self.label1.Text = "Pyomo is not installed by default in SolverStudio. You should install it yourself from the link below.\n\nBe sure to add the location of 'pyomo.exe' (usually '\\PythonXY\\Scripts') to the system path environment variable."
		self.label1.Location = Point(10,10)
		self.label1.Height  = 75
		self.label1.Width = 380
		self.myOK=Button()
		self.myOK.Text = "OK"
		self.myOK.Location = Point(310,90)
		self.myOK.Click += self.CloseForm	
		self.link1 = LinkLabel()
		self.link1.Location = Point(10, 95)
		self.link1.Width = 380
		self.link1.Height = 20
		self.link1.LinkClicked += VisitPYOMODownloadPage
		self.link1.VisitedLinkColor = Color.Blue;
		self.link1.LinkBehavior = LinkBehavior.HoverUnderline;
		self.link1.LinkColor = Color.Navy;
		self.link1.Text = 'http://www.pyomo.org/installation/'
		self.link1.LinkArea = LinkArea(0,49)
		self.AcceptButton = self.myOK
		self.Controls.Add(self.myOK)		
		self.Controls.Add(self.link1)		
		self.Controls.Add(self.label1)
		self.CenterToScreen()
	
	def CloseForm(self,sender,event):
		self.Close()  
		
def UpdateMenu(key,e):
	solver = GetSolverFromSheet()
	if solver == None: solver = SolverStudio.ActiveModelSettings.GetValueOrDefault("COOPRSolver",'cbc').lower().strip()
	for i in solverMenu:
		if solver == i.lower():
			solverMenu[i].Checked = True
			SolverStudio.ActiveModelSettings["COOPRSolver"]=i
		else:
			solverMenu[i].Checked = False

def emptyClickMenu(key,e):
	return

# Get the cell on the sheet (if any) that contains the name of the chosen solver for Pyomo in a 'PyomoOptions' named range
def GetSolverSheetRange():
	r = None
	try:
		r = Application.ActiveSheet.Range("PyomoOptions")
	except:
		return None
	if r.Columns.Count <> 2:
		return None
	try:
		for row in range(r.Rows.Count):
			try:
				key = r.Cells(row+1,1).Value2
				key=key.lower().strip()
				if key=="solver" or key=="-solver" or key=="--solver":
					return r.Cells(row+1,2)
			except:
				continue
	except:
		return None

def GetSolverFromSheet():
	# Get the name of the solver as defined on the sheet in a PyomoOptions named range
	r = GetSolverSheetRange()
	if r==None: return None
	try:
		return str(r.Value2).lower().strip()
	except:
		return None

def SolverClickMenu(key,e):
	for i in solverMenu:
		solverMenu[i].Checked=False
	key.Checked = True
	SolverStudio.ActiveModelSettings["COOPRSolver"]=str(key) # Why is this key, and not key.Text?
	# Update the solver choice on the sheet, if it is defined on the sheet
   # Note: The name "COOPRSolver" comes from the original COOPR/Pyomo name for the package
	r = GetSolverSheetRange()
	if r!=None:
		try:
			r.Value = str(key.Text).lower() # update value on sheet. 20151029: Ensure it is in lower case
		except:
			pass
