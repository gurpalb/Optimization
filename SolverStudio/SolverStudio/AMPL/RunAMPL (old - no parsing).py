# Updated 2012.02.23 to correctly interpret non-tuple indices
# Updated 2012.07.20 to interpret AMPL indices with spaces (by using csv.reader)
#
# This version of the AMPL runner does NOT parse the AMPL model file, but instead analyses the meta
# variables _SETS and _PARS. However, this happens AFTER the data is written, meaning that the
# data type (set or singleton value) is not known when the data it written to a file. Thus, it is not  
# possible to tell that a singleton is actually a member of a set.
# It would be possible to write all singletons as both a set and a singleton, but instead
# we follow the GMPL strategy of parsing the AMPL model file; see the new version for this.

import os 
import csv

# Return s as a string suitable for indexing, which means tuples are listed items in []
def WriteIndex(f,s):
   if type(s) is tuple:
      #f.write('['+','.join([repr(t) for t in s])+']')
      f.write('['+repr(s)[1:-1]+']')
   else:
      f.write(repr(s))

# Return s as a string suitable for an item in a set, which means tuples are listed items in ()
def WriteItem(f,s):
   if type(s) is tuple:
      #f.write( '('+','.join([repr(t) for t in s])+')' )
      f.write( repr(s) )
   else:
      f.write(repr(s))

def WriteParam(name,value):
   f = open(name+".dat","w")
   f.write("param "+name+" := "+repr(value)+";\n\n")
   f.close()

# set NUTR := A B1 B2 C ;
def WriteSet(name,value):
   f = open(name+".dat","w")
   f.write("set "+name+" := ")
   j = 0
   for i in value:
      if j>20:
         j=0
         f.write("\n")
      f.write(" ")
      WriteItem(f,i)
      j = j+1
   f.write(";\n\n")
   f.close()
  
# param demand := 
#   2 6 
#   3 7 
#   4 4 
#   7 8; 
# param sc  default 99.99 (tr) :

def WriteIndexedParam(name,value):
   f = open(name+".dat","w")
   f.write("param "+name+" :=")
   for (index,val) in value.iteritems():
      f.write("\n  ")
      WriteIndex(f,index)
      f.write(" ")
      WriteItem(f,val)
   f.write("\n;\n\n")
   f.close()

# Create a 'data' file, that contains if statements to read in the appropriate data
filename = "SheetData.dat"
g = open(filename,"w")
g.write("# AMPL data file created by SolverStudio\n\n\n")
g.write("# Work around an AMPL bug; _SETS and _PARS need to be displayed or read from to make them work\n")
g.write("# This has been confirmed with email with Bob Fourer Nov 2011.\n")
g.write("# He suggested the LET below; previously this displayed _SETS to a dummy file using >.\n")
# g.write("data;\n") 
# g.write("let DUMMY := _SETS;\n")
g.write("display _SETS > AMPLBugFile.txt;\n")
g.write("display _PARS > AMPLBugFile.txt;\n")

print "## Building AMPL input file for %s data items"%len(SolverStudio.DataItems)

# We avoid iterating thru DataItems and using, e.g., isinstance(value,str)  as isinstance seems very slow
print "## Writing constants"
# Create constants
for (name,value) in SolverStudio.SingletonDataItems.iteritems():
    WriteParam(name,value)
    g.write("if '"+name+"' in _PARS then {data "+name+".dat;}\n")

print "## Writing sets"
# Create sets
for (name,value) in SolverStudio.ListDataItems.iteritems():
    WriteSet(name,value)
    g.write("if '"+name+"' in _SETS then {data "+name+".dat;}\n")
   
print "## Writing indexed parameters"
# Create indexed params
for (name,value) in SolverStudio.DictionaryDataItems.iteritems():
    WriteIndexedParam(name,value)
    g.write("if '"+name+"' in _PARS then {data "+name+".dat;}\n")

g.close()

# Ensure the file "Sheet" that AMPL writes its results into is empty
# (We don't just delete it, as that require import.os to use os.remove)
resultsfilename = "Sheet"
g = open(resultsfilename,"w")
g.close()

print "## Running AMPL..."
print "## AMPL file:",SolverStudio.ModelFileName+"\n"

# Run AMPL, ensuring we add the AMPL directory to the path which allows the student version to find the solver such as CPLEX in this folder
exitCode = SolverStudio.RunExecutable(SolverStudio.GetAMPLPath(),'"'+SolverStudio.ModelFileName+'"', addExeDirectoryToPath=True)
print

if exitCode==0:
   print "## AMPL.EXE completed successfully."
   g = open(resultsfilename,"r")
   
   # Read a file such as: 
   #flow :=
   #A 1   0
   #B 5   0
   #;
   #age = 22
   # Amount [*] :=     <<<<< Note the [*]
   #    BEEF  60
   # CHICKEN   0
   #     GEL  40
   #  MUTTON   0
   #    RICE   0
   #   WHEAT   0
   # ;
   # TODO: We have switched to csv.reader; we can improve the code as a result
   lineCount = 0
   itemsLoaded = ""
   itemsChanged = False
   while True :
      line = g.readline()
      if not line: break
      lineCount += 1
      line = line.strip()
      if line=="": continue # skip blank lines
      # Parse the line into its space delimited components, recognising quoted items containing spaces
      items = csv.reader([line], delimiter=' ', skipinitialspace=True, quotechar="'").next()
      if len(items)==0: continue
      name = items[0]
      #print name,"="
      dataItem = SolverStudio.DataItems[name]
      changed = False
      if len(items)==3 and items[1]=="=":
         # Read a line like: Age = 22 or 'Boy Age' = 7
         try:
            value = float(items[-1])
         except:
            raise TypeError("When reading the results for data item '"+repr(name)+"', an error occurred in line "+repr(line)+" when converting the last item '"+repr(items[-1])+"' to a number. (Line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")
         try:
            if dataItem!=value:
               vars()[name]=value # Update the python variable of this name (without creating a new variable, which would not be seen by SolverStudio)
               changed = True
         except:
            raise TypeError("When reading the results for data item "+repr(name)+", an error occurred in line "+repr(line)+" when assigning the value "+repr(items[-1])+" to index "+repr(tuple(index))+". (Line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")
      else:
         # Read lines such as: 
         #flow :=    **OR, if we have one index**   flow [*] :=
         #'A B' 1   0
         #B 5   0
         #;
         if items[-1] != ":=":
            raise TypeError("When reading the results for data item '"+repr(name)+"', an error occurred in line "+repr(line)+"; '"+repr(items[-1])+"' was expected to be ':='. (Line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")
         while True:
            line = g.readline()
            items = csv.reader([line], delimiter=' ', skipinitialspace=True, quotechar="'").next()
            if len(items)==0: continue
            if items[0]==";": break
            if len(items)<2:
               raise TypeError("When reading the results for data item "+repr(name)+", the line "+repr(line)+" did not contain an index and a value. (Line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")
            try:
               value = float(items[-1])
            except:
               raise TypeError("When reading the results for data item '"+repr(name)+"', an error occurred in line "+repr(line)+" when converting the last item '"+repr(items[-1])+"' to a number. (Line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")
            index = items[0:-1]
            # Convert any values into floats
            for j in range(len(index)):
               try:
                  index[j]=float(index[j])
               except ValueError:
                  pass
            #print tuple(index),value
            try:
               if len(index)==1:
                  index2 = index[0]  # The index is a single value, not a tuple
               else:
                  index2 = tuple(index)# The index is a tuple of two or more items
               if dataItem[index2]!=value:
                  dataItem[index2]=value
                  changed = True
            except:
               raise TypeError("When reading the results for data item "+repr(name)+", an error occurred in line "+repr(line)+" when assigning the value "+repr(items[-1])+" to index "+repr(index2)+". (Line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")            
      if changed:
         itemsLoaded = itemsLoaded+name+"* "
         itemsChanged = True
      else:
         itemsLoaded = itemsLoaded+name+" "
   if itemsLoaded=="":
      print "## No results were loaded into the sheet."
   elif not itemsChanged:
      print "## Results loaded for data items:", itemsLoaded
   else :
      print "## Results loaded for data items:", itemsLoaded
      print "##   (*=data item values were changed on sheet)"
else:
   print "## AMPL did not complete (Error Code %d); no solution is available. The sheet has not been changed." % exitCode
print "## Done"
