import clr
clr.AddReference('System.Windows.Forms')
# clr.AddReference('Microsoft.Office.Tools.Ribbon');
import webbrowser
import os 

#from System.Windows.Forms import Application, Form, Label, ToolStripMenuItem
from System.Windows.Forms import Form, Label, ToolStripMenuItem, MessageBox

import version

def OpenGMPLWikiSite(key,e):
  webbrowser.open("""http://en.wikibooks.org/wiki/GLPK/GMPL_%28MathProg%29""")

def VisitAMPLBookWebPage(key, e):
  webbrowser.open("""http://www.ampl.com/BOOK/download.html""")
  
def OpenGLPKWebSite(key,e):
  webbrowser.open("""http://www.gnu.org/s/glpk/""")
  
def OpenGMPLpdf(key,e):
  oldDirectory = SolverStudio.ChangeToLanguageDirectory()
  os.startfile("gmpl.pdf")  
  SolverStudio.SetCurrentDirectory(oldDirectory)

def OpenCOPYING(key,e):
  oldDirectory = SolverStudio.ChangeToLanguageDirectory()
  Application.Workbooks.Open(Filename="COPYING.", ReadOnly=True)   
  SolverStudio.SetCurrentDirectory(oldDirectory)

def OpenREADME(key,e):
  oldDirectory = SolverStudio.ChangeToLanguageDirectory()
  Application.Workbooks.Open(Filename="README.", ReadOnly=True)   
  SolverStudio.SetCurrentDirectory(oldDirectory)

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
  #   Application.Workbooks.Open(Filename=dataFilePath, ReadOnly=True)
  SolverStudio.ShowTextFile(dataFilePath) # This will show the file and update 
  
def About(key,e):
  MessageBox.Show("SolverStudio GMPL Processor:\n"+version.versionString)

def MenuTester2(key,e):
  #form = Form(Text="Hello World Form Menu Tester")
  #label = Label(Text="Hello World!")
  #form.Controls.Add(label)
  #form.ShowDialog()
  pass
  
# Return s as a string suitable for indexing, which means tuples are listed items in []
def Initialise():
  #Menu.DropDownItems.Add( ToolStripMenuItem("Test",None,MenuTester) )
  Menu.Add( "View Last Data File",OpenDataFile)
  Menu.AddSeparator()
  Menu.Add( "Open GMPL wiki page",OpenGMPLWikiSite) 
  Menu.Add( "Open GNU GLPK web page",OpenGLPKWebSite) 
  Menu.Add( "Open GMPL pdf manual",OpenGMPLpdf) 
  Menu.Add( "Open README file",OpenREADME) 
  Menu.Add( "Open COPYING file",OpenCOPYING)
  Menu.Add( "Open online AMPL book",VisitAMPLBookWebPage)
  Menu.AddSeparator()
  Menu.Add( "About GMPL Processor",About)
  
  
