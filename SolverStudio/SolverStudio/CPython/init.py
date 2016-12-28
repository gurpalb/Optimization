import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import LinkBehavior,LinkArea,LinkLabel,FormBorderStyle,Form, Label,Button,MessageBox, MessageBoxButtons, DialogResult
from System.Drawing import Point,Color
import os
import os.path
import version
def OpenCPythonWebSite(key,e):
	import webbrowser
	webbrowser.open("""http://www.python.org/""")
  
def OpenCPythonOnlineDocumentation(key,e):
	import webbrowser
	webbrowser.open("""http://www.python.org/doc/""")

def OpenCPythonDownload(key,e):
	import webbrowser
	webbrowser.open("""http://www.python.org/getit/""")
	
def OpenActivePythonDownload(key,e):
	import webbrowser
	webbrowser.open("""http://www.activestate.com/activepython/downloads/""")
  
def OpenResultFile(key,e):
  oldDirectory = SolverStudio.ChangeToWorkingDirectory()
  if os.path.exists( "CPythonResults.py" ):
       SolverStudio.ShowTextFile("CPythonResults.py")
  SolverStudio.SetCurrentDirectory(oldDirectory)

def Initialise():
	#Menu.DropDownItems.Add( ToolStripMenuItem("Test",None,MenuTester) )
	Menu.Add( "View last data file",OpenResultFile) 
	Menu.AddSeparator()
	Menu.Add( "Open Python web page",OpenCPythonWebSite) 
	Menu.Add( "Open Python online documentation",OpenCPythonOnlineDocumentation) 
	Menu.AddSeparator()
	Menu.Add( "Open CPython download web page",OpenCPythonDownload)
	Menu.Add( "Open ActivePython download web page",OpenActivePythonDownload)
	Menu.AddSeparator()
	Menu.Add( "About CPython Processor",About)
	oldDirectory = SolverStudio.ChangeToLanguageDirectory()
	if SolverStudio.GetCPythonPath() == None:
		myForm = DownloadForm()
		myForm.ShowDialog()	
	SolverStudio.SetCurrentDirectory(oldDirectory)
def About(key,e):
	MessageBox.Show("SolverStudio CPython Processor:\n"+version.versionString)	
class DownloadForm(Form):
	def __init__(self):
		self.Text = "SolverStudio"
		self.FormBorderStyle = FormBorderStyle.FixedDialog    
		self.Height=150
		self.Width = 400
		self.label1 = Label()
		self.label1.Text = "Python is not installed by default in SolverStudio. You should install it yourself using one of the following links."
		self.label1.Location = Point(10,10)
		self.label1.Height  = 30
		self.label1.Width = 380
		self.myOK=Button()
		self.myOK.Text = "OK"
		self.myOK.Location = Point(310,90)
		self.myOK.Click += self.CloseForm	
		self.link1 = LinkLabel()
		self.link1.Location = Point(10, 50)
		self.link1.Width = 380
		self.link1.Height = 20
		self.link1.LinkClicked += OpenCPythonDownload
		self.link1.VisitedLinkColor = Color.Blue;
		self.link1.LinkBehavior = LinkBehavior.HoverUnderline;
		self.link1.LinkColor = Color.Navy;
		self.link1.Text = 'CPython: www.python.org/getit/'
		self.link1.LinkArea = LinkArea(9,21)
		self.link2 = LinkLabel()
		self.link2.Location = Point(10, 70)
		self.link2.Width = 380
		self.link2.Height = 40
		self.link2.LinkClicked += OpenActivePythonDownload
		self.link2.VisitedLinkColor = Color.Blue;
		self.link2.LinkBehavior = LinkBehavior.HoverUnderline;
		self.link2.LinkColor = Color.Navy;
		self.link2.Text = 'ActivePython: www.activestate.com/activepython/downloads/'
		self.link2.LinkArea = LinkArea(14,43)
		self.AcceptButton = self.myOK
		self.Controls.Add(self.myOK)		
		self.Controls.Add(self.link1)		
		self.Controls.Add(self.link2)		
		self.Controls.Add(self.label1)
		self.CenterToScreen()
	
	def CloseForm(self,sender,event):
		self.Close()