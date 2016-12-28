import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
# clr.AddReference('Microsoft.Office.Tools.Ribbon');
import os
import sys
from System.Windows.Forms import Form, Label, ToolStripMenuItem, Button, TextBox, FormBorderStyle,FolderBrowserDialog,DialogResult,MessageBox,MessageBoxButtons, ScrollBars
from System.Drawing import Point
from _winreg import *
import version

def OpenGurobiWebSite(key,e):
	import webbrowser
	webbrowser.open("""http://www.gurobi.com/""")
  
def OpenGurobiOnlineDocumentation(key,e):
	import webbrowser
	webbrowser.open("""http://http://www.gurobi.com/documentation/5.0/reference-manual/node522""")

def OpenGurobiDownload(key,e):
	import webbrowser
	webbrowser.open("""http://www.gurobi.com/download/gurobi-optimizer""")
  
def OpenResultFile(key,e):
  oldDirectory = SolverStudio.ChangeToWorkingDirectory()
  if os.path.exists( "GurobiResults.py" ):
       SolverStudio.ShowTextFile("GurobiResults.py")
  SolverStudio.SetCurrentDirectory(oldDirectory)
  
def OpenDataFile(key,e):
  oldDirectory = SolverStudio.ChangeToWorkingDirectory()
  if os.path.exists( "GurobiData.py" ):
       SolverStudio.ShowTextFile("GurobiData.py")
  SolverStudio.SetCurrentDirectory(oldDirectory)
  
class ResultForm(Form):

	def __init__(self):
		self.Text = "SolverStudio"
		self.FormBorderStyle = FormBorderStyle.FixedDialog    
		self.Height=360
		self.Width = 370
  
		self.label = Label()
		self.label.Text = ""
		self.label.Location = Point(10,10)
		self.label.Height  = 20
		self.label.Width = 200
  
 		self.bOK=Button()
		self.bOK.Text = "OK"
		self.bOK.Location = Point(150,300)
		self.bOK.Click += self.OK
		
		self.Results = TextBox()
		self.Results.Multiline = True
		self.Results.ScrollBars = ScrollBars.Vertical
		self.Results.Location = Point(10,40)
		self.Results.Width = 350
		self.Results.Height = 250
		
		self.AcceptButton = self.bOK
		
		self.Controls.Add(self.label )
		self.Controls.Add(self.bOK)
		self.Controls.Add(self.Results)
		self.CenterToScreen()
  
	def OK(self,key,e):
		self.Close()
  
class InstallForm(Form):

	def __init__(self):
		try:
			GrbLicenceFile = os.environ["GRB_LICENSE_FILE"]
		except:
			GrbLicenceFile = os.environ["USERPROFILE"] + "\\gurobi.lic"
		self.Text = "SolverStudio Gurobi License Manager"
		self.FormBorderStyle = FormBorderStyle.FixedDialog    
		self.Height=365
		self.Width = 460

		self.label = Label()
		self.label.Text = "Please paste your license key here:"
		self.label.Location = Point(10,10)
		self.label.Height  = 20
		self.label.Width = 230

		self.bNewKey=Button()
		self.bNewKey.Text = "Get New Key..."
		self.bNewKey.Location = Point(335,28)
		self.bNewKey.Width = 100
		self.bNewKey.Click += self.NewKey
		
		self.bBrowse=Button()
		self.bBrowse.Text = "Browse..."
		self.bBrowse.Location = Point(335,83)
		self.bBrowse.Width = 100
		self.bBrowse.Click += self.browse
		
		self.label2 = Label()
		self.label2.Text = "grbgetkey"
		self.label2.Location = Point(10,32)
		self.label2.Height  = 15
		self.label2.Width = 55

		# We used to store license keys under 'Software\SolverStudio\Gurobi\'+os.getenv("COMPUTERNAME")
		# We now store them under 'Software\SolverStudio\Gurobi\'+os.getenv("COMPUTERNAME")+'\'+os.getenv("USERNAME")
		# as this works even if the HKEY_CURRENT_USER registry entries do not change for each user, as happens with no roaming profile
		# try:
			# # Get license key values from old location in registry
			# keyVal = 'Software\\SolverStudio\\Gurobi\\'+os.getenv("COMPUTERNAME")
			# key = OpenKey(HKEY_CURRENT_USER, keyVal, 0, KEY_ALL_ACCESS)
			# license = QueryValueEx(key,"LicenseString")[0]
			# licensepath = QueryValueEx(key,"GurobiPath")[0]
			# CloseKey(key)
			# # Write old key values to new registry location
			# keyVal = 'Software\\SolverStudio\\Gurobi\\'+os.getenv("COMPUTERNAME")+'\\'+os.getenv("USERNAME")
			# key = OpenKey(HKEY_CURRENT_USER, keyVal, 0, KEY_ALL_ACCESS)
			# SetValueEx(key, "LicenseString", 0, REG_SZ, license)
			# SetValueEx(key, "GurobiPath", 0, REG_SZ, licensepath)
			# CloseKey(key)
		# except:
			# self.KeyText.Text = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
		
		
		self.KeyText = TextBox()
		try:
			keyVal = 'Software\\SolverStudio\\Gurobi\\'+os.getenv("COMPUTERNAME")
			key = OpenKey(HKEY_CURRENT_USER, keyVal, 0, KEY_ALL_ACCESS)
			license = QueryValueEx(key,"LicenseString")
			self.KeyText.Text = license[0]
			CloseKey(key)
		except:
			self.KeyText.Text = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
		if self.KeyText.Text=="": self.KeyText.Text = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
		self.KeyText.Location = Point(65,30)
		self.KeyText.Width = 265
		
		self.label3 = Label()
		self.label3.Text = "Where would you like to save the license file?"
		self.label3.Location = Point(10,65)
		self.label3.Height  = 20
		self.label3.Width = 350

		self.KeyPath = TextBox()
		try:
			self.KeyPath.Text = GrbLicenceFile.replace("\\gurobi.lic","")
		except:
			self.KeyPath.Text = os.environ["USERPROFILE"]
		self.KeyPath.Location = Point(10,85)
		self.KeyPath.Width = 320

		self.Username = Label()
		grbLicenseFileEnvVar = os.getenv("GRB_LICENSE_FILE")
		if grbLicenseFileEnvVar == None:
			grbLicenseFileEnvVar ="[not defined]"
		self.Username.Text = "Machine Environment Variable Values:\nUSERNAME: " + os.environ.get("USERNAME") + "\n" + "HOSTNAME: " + os.getenv("COMPUTERNAME") + "\nGRB_LICENSE_FILE: " + grbLicenseFileEnvVar
		self.Username.Location = Point(10,115)
		self.Username.Height  = 65
		self.Username.Width = 330
		
		licenceFound = False
		licenseText = "No License Found."
		try:
		  with open(GrbLicenceFile,"rb") as inFile:
				licenseText = inFile.read()
				licenceFound = True
		except: 
			pass

		self.licenseLabel = Label()
		if (licenceFound):
			self.licenseLabel.Text = "Current License File (" + GrbLicenceFile + "):"
		else:
			self.licenseLabel.Text = "Current License File"
		self.licenseLabel.Location = Point(10,180)
		self.licenseLabel.Height  = 20
		self.licenseLabel.Width = 350
		
		self.LicenseText = TextBox()
		self.LicenseText.Multiline = True
		self.LicenseText.ScrollBars = ScrollBars.Vertical
		self.LicenseText.Text = licenseText
		self.LicenseText.Location = Point(10,200)
		self.LicenseText.Width = 435
		self.LicenseText.Height = 100
		self.LicenseText.SelectAll()
		
		self.bInstall=Button()
		self.bInstall.Text = "Install"
		self.bInstall.Location = Point(270,310)
		self.bInstall.Click += self.install

		self.bCancel=Button()
		self.bCancel.Text = "Cancel"
		self.bCancel.Location = Point(370,310)
		self.bCancel.Click += self.cancel

		self.AcceptButton = self.bInstall
		self.CancelButton = self.bCancel

		self.Controls.Add(self.KeyText)
		self.Controls.Add(self.bInstall)
		self.Controls.Add(self.bCancel)
		self.Controls.Add(self.bBrowse)
		self.Controls.Add(self.bNewKey)
		self.Controls.Add(self.label)
		self.Controls.Add(self.label2)
		self.Controls.Add(self.label3)
		self.Controls.Add(self.Username)
		self.Controls.Add(self.licenseLabel)
		self.Controls.Add(self.KeyText)
		self.Controls.Add(self.KeyPath)
		self.Controls.Add(self.LicenseText)
		self.CenterToScreen()
		
	def browse(self,sender,event):
		dialog = FolderBrowserDialog()
		dialog.SelectedPath = os.environ["USERPROFILE"]+"\\"
		if (dialog.ShowDialog(self) == DialogResult.OK):
			self.KeyPath.Text = dialog.SelectedPath
		
		
	def install(self,sender,event):
		# Strip any trailing \
		if self.KeyPath.Text[-1]=="\\":
			self.KeyPath.Text = self.KeyPath.Text[:-1]
		try:
			path = self.KeyPath.Text+"\\tmpFile"
			tmp = open(path,"w")
			tmp.close()
			os.remove(path)
		except:
			choice = MessageBox.Show("You do not have permission to save the Gurobi license to \""+self.KeyPath.Text+"\"\nChoose a different location.","SolverStudio",MessageBoxButtons.OK)
			return	
	
		if os.path.exists(self.KeyPath.Text+"\\gurobi.lic"):
			choice = MessageBox.Show("License file already exists in this location. Overwrite?","SolverStudio",MessageBoxButtons.YesNo)
			if choice == DialogResult.No:
				return

		# This worked on the older Gurobi code; try it first.
		succeeded = False
		data = ""
		tempDir = SolverStudio.WorkingDirectory()
		if not tempDir[-1]=="\\":	tempDir=tempDir+"\\"
		f=open(tempDir+"KeyPath.txt","w")
		f.write(self.KeyPath.Text)
		f.close()
		command = self.KeyText.Text + ' < "'+tempDir+'KeyPath.txt" > "'+tempDir+'GRBoutput.txt"'
		try:
			succeeded = (os.system("grbgetkey " + command)==0) # This finds grbgetkey somewhere on the PATH
		except:
			pass
		if succeeded:
			with open(tempDir+"GRBoutput.txt","rb") as output:
				data=output.read()
		
		# exitCode = SolverStudio.RunExecutable("grbgetkey",command,False,False)
		# does not work as it only saves to the Default Directory %USERPROFILE%
			
		if (not succeeded):
			# For GRBGETKEY(5.1.0) use:
			try:
				command = "--path=\"" + self.KeyPath.Text + "\" " + self.KeyText.Text # NB: Fails with a trailing \ in the path
				data = SolverStudio.RunEXE_GetOutput("grbgetkey",command,60000) # This finds grbgetkey somewhere on the PATH
			except:
				pass
				
		if data.strip() == "":
			data = "Running grbgetkey.exe failed, perhaps because Gurobi is not properly installed on your system.  Note that you need to re-start SolverStudio after installing Gurobi to allow the Gurobi files to be found."

		SolverStudio.AppendToTaskPane(data)
		
		myResultForm = ResultForm()
		myResultForm.Results.Text = data
		if succeeded:
			keyVal = 'Software\\SolverStudio\\Gurobi\\'+os.getenv("COMPUTERNAME")
			try:
				key = OpenKey(HKEY_CURRENT_USER, keyVal, 0, KEY_ALL_ACCESS)
			except:
				key = CreateKey(HKEY_CURRENT_USER, keyVal)
			SetValueEx(key, "GurobiPath", 0, REG_SZ, self.KeyPath.Text + "\\gurobi.lic")
			SetValueEx(key, "LicenseString", 0, REG_SZ, self.KeyText.Text)
			CloseKey(key)
			try:
				os.environ["GRB_LICENSE_FILE"] = self.KeyPath.Text + "\\gurobi.lic"
				if (os.environ["GRB_LICENSE_FILE"] == self.KeyPath.Text + "\\gurobi.lic"):
					envFail = False
				else:
					envFail = True
			except:
				envFail = False
			if (os.environ["USERPROFILE"] != self.KeyPath.Text):
				if envFail:
					data = "##############################################\r\nSolverStudio:  You have chosen to install the Gurobi License to a non-default location. SolverStudio was unable to set the environment variable 'GRB_LICENSE_FILE' to be \"" + self.KeyPath.Text + "\\gurobi.lic\".\r\n\r\nYou should set this yourself, or reinstall the license to the default location '"+os.environ["USERPROFILE"]+"'.\r\n##############################################\r\n\r\n"  + data
				else:
					data = "##############################################\r\nSolverStudio:  You have chosen to install the Gurobi License to a non-default location.\r\n\r\nSolverStudio succeeded in setting the environment variable 'GRB_LICENSE_FILE' to be '" + self.KeyPath.Text + "\\gurobi.lic'.\r\n##############################################\r\n\r\n" + data
			myResultForm.label.Text = "License installation SUCCEEDED"
		else:
			myResultForm.label.Text = "License installation FAILED"
		myResultForm.Results.Text = data	
		myResultForm.ShowDialog(self)
		self.Close()
		
	def cancel(self,sender,event):
		self.Close()
	
	def NewKey(self,sender,event):
		from webbrowser import open
		open("""http://www.gurobi.com/download/licenses/current""")
		
def LicenseGurobi(key,e):
    #import System.Windows.Forms.Application
    myInstallForm = InstallForm()
    myInstallForm.ShowDialog()
	 #System.Windows.Forms.Application.Run(myInstallForm)
	 
def Initialise():
	# AJM 20130326 Disable the following code which changes the GRB_LICENSE_FILE environment variable.
	# We now only do this when saving a new key to a non-default location
	# try:
	#	keyVal = 'Software\\SolverStudio\\Gurobi\\'+os.getenv("COMPUTERNAME")
	#	key = OpenKey(HKEY_CURRENT_USER, keyVal, 0, KEY_ALL_ACCESS)
	#	path = QueryValueEx(key,"GurobiPath")
	#	os.environ["GRB_LICENSE_FILE"] = path[0]
	#	CloseKey(key)
	#except: 
	#	pass
	#Menu.DropDownItems.Add( ToolStripMenuItem("Test",None,MenuTester) )
	Menu.Add( "View last data (Gurobi Input) file",OpenDataFile) 
	Menu.Add( "View last results (Gurobi Output) file",OpenResultFile) 
	Menu.AddSeparator()
	Menu.Add( "Install or Renew Gurobi License...",LicenseGurobi)
	Menu.AddSeparator()
	Menu.Add( "Open Gurobi web page",OpenGurobiWebSite) 
	Menu.Add( "Open Gurobi online documentation",OpenGurobiOnlineDocumentation) 
	Menu.Add( "Open Gurobi download web page",OpenGurobiDownload) 
	Menu.Add( "About Gurobi Processor...",About)
	
def About(key,e):
	MessageBox.Show("SolverStudio Gurobi Processor:\n"+version.versionString)