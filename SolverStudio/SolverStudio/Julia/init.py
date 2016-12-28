import webbrowser
import os 

import clr
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import Form, Label, ToolStripMenuItem, Button, TextBox, FormBorderStyle,FolderBrowserDialog,OpenFileDialog,DialogResult,MessageBox,MessageBoxButtons, ScrollBars

clr.AddReference('System.Drawing')
from System.Drawing import Point

# from IronPython.Runtime.Calls import CallTarget0 - fails; been replaced by:
clr.AddReference('IronPython')
from IronPython.Compiler import CallTarget0
#see  "Delegates and CallTarget0" at http://www.voidspace.org.uk/ironpython/dark-corners.shtml

# We also need acces to Action as CallTarget1 etc don't seem to exist any more
clr.AddReference('System.Core')
from System import Action, Func

import version # Our SolverStudio version file

import urllib
import urllib2
from subprocess import call

import xml.etree.ElementTree as ET

lastFilePercent = 0

def OpenJuliaInterface(key,e):
	try:
		exeFile = SolverStudio.GetJuliaPath(mustExist = True);
	except:
		raise Exception("Julia cannot be found, perhaps because it is not installed.")
	try:
		oldDirectory = SolverStudio.ChangeToWorkingDirectory()
		modulePath = SolverStudio.WorkingDirectory() + "/SolverStudio.jl"
		# The following code, which we don't use now, includes the SolverStudio data items via SolverStudio.jl
		#if os.path.isfile(modulePath) :
		#	# Reason for complicated string is because having multiple sets of "" confuses CMD
		#	command_part_1 = '''-P include_string(""""using SolverStudio\\nprintln(string('S','o','l','v','e','r','S','t','u','d','i','o','''
		#	command_part_2 = '''' ','v','a','r','i','a','b','l','e','s',' ','i','m','p','o','r','t','e','d'))"""")'''
		#	command = command_part_1 + command_part_2
		#else :
		#	command = ""
		command = ""
		SolverStudio.StartEXE(exeFile, command, addExeDirectoryToPath=True, addSolverPath=True, addAtStartOfPath=True)
	finally:
		SolverStudio.SetCurrentDirectory(oldDirectory)
		
def ViewDataFile(key, e):
  dataFilePath = SolverStudio.WorkingDirectory() + "/SolverStudio.jl" 
  SolverStudio.ShowTextFile(dataFilePath)
  
def ViewResultsFile(key, e):
  resultsFilePath = SolverStudio.WorkingDirectory() + "/JuliaResults.py"
  SolverStudio.ShowTextFile(resultsFilePath)
		
def VisitJuliaWebPage(key,e):
  webbrowser.open('http://julialang.org/')

def VisitJuliaOptWebPage(key,e):
  webbrowser.open('http://www.juliaopt.org/')

def OpenJuMPDoc(key,e):
  webbrowser.open('https://jump.readthedocs.org/en/latest/')
  
def GetDownloadURL():
  uri = 'http://s3.amazonaws.com/julialang'
  tree = ET.parse(urllib2.urlopen(uri))
  root = tree.getroot()
  
  if root.tag[0] == "{":
    xmlns = "{" + root.tag[1:].partition("}")[0] + "}"
  else:
    xmlns = None
  
  contentString = xmlns + "Contents"
  keyString = xmlns + "Key"
  # TO-DO: Need a better way of finding bitness
  if 'PROGRAMFILES(X86)' in os.environ:
    bitness = "64"
  else:
    bitness = "32"
  latest = None
  
  for contents in root.findall(contentString):  
    key = contents.find(keyString).text
    key_words = key.split("/")
    if len(key_words) < 5 :
      continue
    if key_words[-1] == "" :
      continue
    if not (key_words[0] == "bin") :
      continue
    if not (key_words[1] == "winnt") :
      continue
    bitness_words = key_words[-1].split(".")
    if not (bitness_words[-1] == "exe"):
      continue
    if not (bitness_words[-2][-2:] == bitness):
      continue
    latest = key
  
  if latest == None:
    url = None
  else:
    url = uri + "/" + latest
  return url, bitness

def CallDownloadJulia(key, e):
  SolverStudio.RunInDialog("This will download and launch the Julia installer\n", DownloadJulia)
  
def DownloadJulia(workManager, workInfo):
  url, bitness = GetDownloadURL()
  if url == None:
    print "URL for download could not be obtained"
    return
  name = SolverStudio.WorkingDirectory() + "/" + url.split('/')[-1]
  try: 
    urllib2.urlopen(url)
  except urllib2.HTTPError, e:
    print "Download link could not be found"
    return
  try:
    print "\nDownloading Julia from " + url
    urlwords = url.split("/")
    version = urlwords[-1].split("-")[1]
    print "\nDownloading Julia version " + version + " installer for Windows " + bitness + " bit"
    name, hdrs = urllib.urlretrieve(url, name, reporter)
  except IOError, e:
    print "Can't retrieve Julia"
    return
  if (workManager.CancellationPending):
    print "\n\nThe download has been cancelled"
    try: 
      os.unlink(name)
    except:
      pass
    return
  print "\n\nDownload complete."
  print "Launching Julia installer (in another window)."
  os.startfile(name)
  print "SolverStudio installation steps completed."

def CallInstallJump(key, e):
  SolverStudio.RunInDialog("This will install/update JuMP and relevant packages. Make sure you have Julia installed",InstallJump)
  
def InstallJump(workManager, workInfo):
  setupFile = SolverStudio.WorkingDirectory() + "/JuMPsetup.jl"
  rebuildFile = SolverStudio.WorkingDirectory() + "/JuMPrebuild.jl"
  exeFile = SolverStudio.GetJuliaPath(mustExist = True);
  
  g = open(setupFile, "w")
  g.write('''println("\\nAdding JuMP\\n")''' + "\n")
  g.write('''Pkg.add("JuMP")''' + "\n")
  g.write('''println("\\nAdding Clp\\n")''' + "\n")
  g.write('''Pkg.add("Clp")''' + "\n")
  g.write('''println("\\nUpdating packages\\n")''' + "\n")
  g.write("Pkg.update()\n")
  g.write('''println("\\nBuilding Clp\\n")''' + "\n")
  g.write('''Pkg.build("Clp")''' + "\n")
  g.close()
  
  f = open(rebuildFile, "w")
  f.write('''println("\\nUpdating packages\\n")''' + "\n")
  f.write("Pkg.update()\n")
  f.write('''println("\\nRebuilding Clp\\n")''' + "\n")
  f.write('''Pkg.build("Clp")''' + "\n") 
  f.close()
  
  #if not os.path.isfile(exeFile) :
  #  print "ERROR: Julia could not be found"
  #  return
  
  if (workManager.CancellationPending):
    print "\n\nThe installation of JuMP has been cancelled"
    return
  
  print "\n\nInstalling JuMP"
  call([exeFile, setupFile])
  
  if (workManager.CancellationPending):
    print "\n\nThe installation of JuMP has been cancelled"
    return
  
  # Julia seems to require a reset for solvers to install/build correctly
  print "Rebuilding packages"
  call([exeFile, rebuildFile])
  
  if (workManager.CancellationPending):
    print "\n\nThe installation of JuMP has been cancelled"
    return
  
  print "JuMP, Cbc & Clp are added and built"
  print "All packages are updated"

def reporter(block_count, block_size, file_size):
  global lastFilePercent
  if file_size>0:
    filePercent = int(block_count*block_size/(0.01*file_size)+0.5)
    if filePercent>lastFilePercent: 
      if lastFilePercent % 10 == 0: print
      print "%d%%" % filePercent,
    lastFilePercent = filePercent
  else:
    print block_count," "    

def KillAnyRunningCommand():
	SolverStudio.InteractiveExes.GetInteractiveExe("Julia").KillAnyRunningCommand()

class ConsoleForm(Form):

	def __init__(self):
		self.Text = "SolverStudio Julia Console"
		self.FormBorderStyle = FormBorderStyle.FixedDialog    
		self.Height= 130
		self.Width = 460

		# self.Closed += CallTarget2(self.OnClosed)   # Kill any running command
		self.Closed += self.FormClosed # Never got this to work
		# self.ControlBox = False # Force the user to quit using the Close button which we can handle; getting a closing/closed event from the form is too hard

		self.label = Label()
		self.label.Text = "Please enter your Julia command:"
		self.label.Location = Point(10,10)
		self.label.Height  = 20
		self.label.Width = 230

		self.KeyText = TextBox()
		self.KeyText.Location = Point(10,30)
		self.KeyText.Width = 340
		self.KeyText.Height  = 20

		self.bRun=Button()
		self.bRun.Text = "Run"
		self.bRun.Location = Point(370,28)
		# self.bRun.Width = 100
		self.bRun.Click += self.Run
		
		self.bCancel=Button()
		self.bCancel.Text = "Close"
		self.bCancel.Location = Point(370,65)
		self.bCancel.Click += self.cancel

		self.AcceptButton = self.bRun
		self.CancelButton = self.bCancel

		self.Controls.Add(self.label)
		self.Controls.Add(self.KeyText)	# Add first to get focus
		self.Controls.Add(self.bRun)
		self.Controls.Add(self.bCancel)
		self.CenterToScreen()

	def FormClosed(self, sender, event):
		KillAnyRunningCommand()

	def cancel(self,sender,event):
		KillAnyRunningCommand()
		self.Close()
		
	# This did not work with a weird error
	#def commandCompletedCallback(self,processExited):
	#	# This gets called on a different thread; use invoke to call it on the creating thread as needed to access controls and forms
	#	if self.InvokeRequired:
	#		 self.Invoke( commandCompletedCallback(self,processExited) );
	#		 return
	#	self.bRun.Enabled = True

	def commandCompletedCallback(self,processExited):
		# commandCompletedCallback gets called on a different thread; use invoke to change to the creating thread as needed to access controls and forms
			# See http://www.ironpython.info/index.php?title=Invoking_onto_the_GUI_%28Control%29_Thread
		def EnableRunButton():
			self.bRun.Enabled = True
			self.bRun.Text = "Run"
		# self.bRun.Invoke( CallTarget0(EnableRunButton) );
		self.bRun.Invoke( CallTarget0(EnableRunButton) );   # Use CallTarget0 if IronPython is unable to convert to a delegate by itself; see http://www.voidspace.org.uk/ironpython/dark-corners.shtml

	def Run(self,sender,event):
		interactiveExe = SolverStudio.InteractiveExes.GetInteractiveExe("Julia")
		if not interactiveExe.IsRunning:
			SolverStudio.AppendToTaskPane( "## Starting Julia ("+interactiveExe.Executable+") console\r\n")
			interactiveExe.Start()
			# print "Working Dir =",interactiveExe.WorkingDirectory;
		commandToSend = self.KeyText.Text
		SolverStudio.AppendToTaskPane(">"+commandToSend+"\r\n") # Print does not work
		#try:
		self.bRun.Enabled = False
		self.bRun.Text = "Running..."
		interactiveExe.SendCommand(self.KeyText.Text, False, self.commandCompletedCallback) # NB: Cannot run in the main thread as it deadlocks if we wait for completion
		#except:
		#	self.bRun.Enabled = True

def SendCommandsToConsole_WorkerSub():
    myConsoleForm = ConsoleForm()
    #myConsoleForm.ShowDialog()
    SolverStudio.ShowDialogWithExcelAsParent(myConsoleForm)

def SendCommandsToConsole(key ,e):
     # We cannot run commands in the main thread as it deadlocks if we wait for completion (which we don't do now... so still needd this?)
    SolverStudio.RunInNonGUIThread(SendCommandsToConsole_WorkerSub)

def InitInteractiveExeConfig():
   SolverStudioName = "Julia" # eg CPython, used to identify the ExeFinder to use
   Arguments = "-i"
   CodeToRunAtStartup = None
   # Before each run, we replace SS_GenerateSolverStudioResultsFile so we can call it after running without ever running an old version which would print old data
   # CodeToRunBeforeCommand = "function SS_GenerateSolverStudioResultsFile(); end;"
   # CodeToRunBeforeCommand = """include("test.jl");"""
   CodeToRunBeforeCommand = """function SS_GenerateSolverStudioResultsFile(); end; print(\"\");""" # Note trailing print() to clear ans, and hence stop any output
      # Note ; to suppress any output of 'ans' (being the function) (but " does not work running in SolverStudio, but does work in Julia Command line??), along with a blank print to ensure ans is not set
   CodeToRunAfterCommand =  "print(\"\\n{0}\\n\")"   # CommandFinishedOutputText will be substituted in for {0}
   CodeToRunToQuit= "quit()"
   CommandFinishedOutputText = None # meaning default
   AddExeDirectoryToPath = False
   AddSolverPath = True
   AddAtStartOfPath = False
   KillChildProcesses = False
   SolverStudio.InteractiveExes.AddInteractiveExeConfig(
      SolverStudioName=SolverStudioName, 
      Arguments=Arguments,
      CodeToRunAtStartup=CodeToRunAtStartup,
      CodeToRunBeforeCommand = CodeToRunBeforeCommand,
      CodeToRunAfterCommand = CodeToRunAfterCommand,
      CodeToRunToQuit = CodeToRunToQuit,
      CommandFinishedOutputText = CommandFinishedOutputText,
      AddExeDirectoryToPath = AddExeDirectoryToPath, 
      AddSolverPath = AddSolverPath, 
      AddAtStartOfPath = AddAtStartOfPath, 
      KillChildProcesses = KillChildProcesses)
# TODO: Add in the path changing etc (but how does this work with the student ampl?)

def OpenExamples(key,e):
  webbrowser.open('https://github.com/JuliaOpt/JuMP.jl/tree/master/examples')
  
def About(key,e):
  MessageBox.Show("Julia: " + version.versionString)

global menuUseConsole

def UpdateConsoleModeMenu():
  menuUseConsole.Checked = (SolverStudio.GetRegistrySetting("Julia","RunInConsole",0) == 1)
  
def ClickConsoleModeMenu(key ,e):
  menuUseConsole.Checked=not menuUseConsole.Checked
  SolverStudio.SetRegistrySetting("Julia","RunInConsole",int(menuUseConsole.Checked))

def Initialise():
  Menu.Add( "Open External Julia Shell",OpenJuliaInterface)
  Menu.AddSeparator()
  Menu.Add( "View last data file", ViewDataFile)
  Menu.Add( "View last results file", ViewResultsFile)
  Menu.AddSeparator()
  global menuUseConsole
  menuUseConsole=Menu.Add( "Run Models using internal Julia Console",ClickConsoleModeMenu)
  UpdateConsoleModeMenu()
  Menu.Add( "Send Commands to internal Julia Console...",SendCommandsToConsole)
  Menu.AddSeparator()
  Menu.Add( "Open Julia website",VisitJuliaWebPage)
  Menu.Add( "Open JuliaOpt website",VisitJuliaOptWebPage)
  Menu.Add( "Open online JuMP documentation", OpenJuMPDoc)
  Menu.Add( "Open online JuMP examples", OpenExamples)
  Menu.AddSeparator()
  Menu.Add( "Download Julia and run installer...", CallDownloadJulia)
  Menu.Add( "Install/Update JuMP...", CallInstallJump)
  Menu.AddSeparator()
  Menu.Add( "About Julia Processor", About)
  InitInteractiveExeConfig()

 