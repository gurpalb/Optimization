""" This file will download and unzip the student version of AMPL. No installer is run.
    The AMPL file is placed in an amplcml directory beside this Python file.
"""    

import os 
import urllib 
import zipfile  

lastFilePercent = 0

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
          
def getunzipped(workManager, workInfo, theurl, thedir):
    # Create path to download to (if needed)
    thedir = os.path.normpath(thedir)
    if thedir <> "":
        if not os.path.isdir(thedir):
            os.makedirs(thedir)    
    name = os.path.join(thedir, 'amplcml.zip')
    try:
        print "Downloading ",theurl,
        name, hdrs = urllib.urlretrieve(theurl, name, reporter)
    except IOError, e:
        print "Can't retrieve %r to %r: %s" % (theurl, thedir, e)
        print "The installation has failed."
        return
    print
    if (workManager.CancellationPending):
        print
        print "Cancelled; no files were installed, and the downloaded .zip file has been deleted."
        workInfo.Cancel = True # Report a Cancelled run
        try: 
            os.unlink(name) # delete any .zip file
        except:
            pass
        return
    try:
        z = zipfile.ZipFile(name)
    except zipfile.error, e:
        print "Bad zipfile (from %r): %s" % (theurl, e)
        return
    print "Extracting items: ",
    realdir = os.path.realpath(thedir)
    amplExe = "<ampl.exe not found in .zip file>"
    for n in z.namelist():
        print n
        # Check the .zip is putting files only in the sub-directory
        if not n.endswith('/'): 
            root, file = os.path.split(n)
            fullPath = os.path.normpath(os.path.join(thedir, root))
            directory = os.path.realpath(fullPath)
            if not directory.startswith(realdir):
                print "The zip file is trying to write to %r, which is not a sub-directory of %r" % (directory, thedir, e)
                return
            if n.lower().endswith("ampl.exe"): amplExe = os.path.realpath(os.path.normpath(os.path.join(thedir, n)))
        z.extract(n,thedir)
        #fname = os.path.join(directory, name)
        #f=file(fname, 'wb')
        #f.write(z.read(n))
        #f.close()
    print "z.fp=",z.fp
    z.close()
    print "z.fp=",z.fp
    print "Zip file extracted."
    try:
        #print "name=",name
        #print "real name =",os.path.realpath(name)
        # The following fails because, under IronPythin, this code still has about 5 handles open to the .zip file
        # These seem to be "weak" handles in the running this code will allow the .zip file to be written over (or at least
        # opened for writing!).
        os.unlink(name) # delete the .zip file
        print "Downloaded Zip file deleted."
    except Exception as err:
        print "The installation succeeded, but the downloaded zip file:\n  "+name+"\ncould not be deleted. The error was:\n  ",err
    print "\nAMPL student version downloaded successfully."
    print "The files are available at:"
    print realdir
    print "The AMPL executable is at:"
    print amplExe

def DoInstall(workManager, workInfo):    
    # Install the student AMPL version to the current directory
    getunzipped(workManager, workInfo, "http://www.ampl.com/NEW/TABLES/amplcml.zip",".")
