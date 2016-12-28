# Return s as a string suitable for indexing, which means tuples are listed items in []
def WriteIndex(f,s):
   if type(s) is tuple:
      k = 0
      for i in s:
         k=k+1
         if isinstance(i,float):
             # We cannot have doubles as indices as we cannot have a '.'
             f.write( repr(int(i+0.5)) )
         elif isinstance(i,int):
             f.write( repr(i) )
         else:
             f.write( i )
         if k<>len(s):
             f.write( "." )
         #f.write('['+','.join([repr(t) for t in s])+']')
         #f.write(repr(s)[1:-1].replace(",","."))
   else:
      if isinstance(s,float):
          # We cannot have doubles as indices as we cannot have a '.'
          f.write( repr(int(s+0.5)) )
      elif isinstance(s,int):
          f.write( repr(s) )
      else:
          f.write( s )

# Return s as a string suitable for an item in a set, which means tuples are listed items in ()
def WriteItem(f,s):
   if type(s) is tuple:
      #f.write( '('+','.join([repr(t) for t in s])+')' )
      f.write( repr(s).replace(",",".") )
   else:
      if isinstance(s,float):
          # We cannot have doubles as indices as we cannot have a '.'
          f.write( repr(int(s+0.5)) )
      elif isinstance(s,int):
          f.write( repr(s) )
      else:
          f.write( s )

def WriteParam(f,name,value):
   f.write("param "+name+" := "+repr(value)+";\n\n")

# set NUTR := A B1 B2 C ;
def WriteSet(f,name,value):
   f.write("Set "+name+" / \n")
   k = 0
   for i in value:
      k=k+1
      f.write(" ")
      WriteItem(f,i)
      if k <> len(value):
         f.write(",\n")
   f.write(" / ;\n\n")

# param demand := 
#   2 6 
#   3 7 
#   4 4 
#   7 8; 
# param sc  default 99.99 (tr) :

def WriteIndexedParam(f,name,value):
   f.write("Parameter "+name+" /\n")
   for (index,val) in value.iteritems():
      f.write(" ")
      WriteIndex(f,index)
      f.write(" ")
      WriteItem(f,val)
      f.write("\n")
   f.write(" / ;\n\n")

def WriteTable(f,name,value):
   f.write("Parameter "+name+" /\n")
   for (index,val) in value.iteritems():
      f.write(" ")
      WriteIndex(f,index)
      f.write(" ")
      WriteItem(f,val)
      f.write("\n")
   f.write(" / ;\n\n")

# Read lines like:
# Sets
# i  canning plants
# j  markets / a b / ;
# ---- and ---- #
# Set i  canning plants / Seattle, San-Diego / ;

def ScanModelFileForDataItems(modelFileName):
   unusedKeys = ["variables","positive","equations","model","solve","scalar","display"]
   Sets= list()
   Params = list()
   Tables = list()
   #singlelineKeyWords = { "set" : Sets, 'parameter' : Params, 'table' : Tables }
   keyWords = { "sets" : Sets, 'parameters' : Params, 'tables' : Params, "set" : Sets, 'parameter' : Params, 'table' : Params}
   lineCount = 0
   f = open(SolverStudio.ModelFileName,"r")
   items = None
   for line in f:
      line = line.strip() # Remove white space, including the newline
      # Replace anything that can follow the name with a 'space' so the string splits correctly.
      l = line.replace("("," ").replace(";"," ; ") #.replace("<"," ").replace(">"," ").replace("{"," ").replace(":"," ").replace(","," ")
      lineCount += 1
      tokens = l.split()
      if len(tokens)==0: continue
      if tokens[0].lower() in keyWords.keys():
         items = keyWords[tokens[0].lower()]
         tokens = tokens[1:]
      elif tokens[0].lower() in unusedKeys:
         items = None
      if len(tokens)==0: continue
      if items <> None:
         if tokens[0]<>";" and not tokens[0] in items:
            # print "##  param ",tokens[1]
            items.append(tokens[0])
			
      #The command ends with a ; so search for this
      if ";" in tokens:
          items = None
   f.close()
   
   
   print "Sets=",Sets
   print "Params=",Params
   return Sets, Params

# Create a 'data' file, that contains the appropriate data
dataFilename = "SheetData.dat"
g = open(dataFilename,"w")
g.write("* GAMS data file created by SolverStudio\n\n")

print "## Scanning model for sets and parameters"
Sets, Params = ScanModelFileForDataItems(SolverStudio.ModelFileName)
#print "Sets=",Sets
#print "Params=",Params

print "## Building GAMS input file for %s sets and %s parameters"%(len(Sets),len(Params))

print "## Writing simple parameters...",
# Create constants
count = 0
for (name,value) in SolverStudio.SingletonDataItems.iteritems():
   if name in Params:
       count += 1
       WriteParam(g, name,value)
print count,"items written."
       
print "## Writing sets...",
# Create sets
count = 0
for (name,value) in SolverStudio.ListDataItems.iteritems():
   if name in Sets:
       count += 1
       WriteSet(g, name,value)
print count,"items written."
   
print "## Writing indexed parameters...",
# Create indexed params
count = 0
for (name,value) in SolverStudio.DictionaryDataItems.iteritems():
   if name in Params:
       count += 1
       WriteIndexedParam(g, name,value)
print count,"items written."

g.close()

# Ensure the file "Sheet" that AMPL writes its results into is empty
# (We don't just delete it, as that require import.os to use os.remove)
resultsfilename = "Sheet"
g = open(resultsfilename,"w")
g.close()

print "## Running GAMS ("+SolverStudio.GetGAMSPath()+")"
print "## GAMS model file:",SolverStudio.ModelFileName
print "## GAMS data file:",dataFilename+"\n"
exitCode = SolverStudio.RunExecutable(SolverStudio.GetGAMSPath()," \""+SolverStudio.ModelFileName+"\"")

if exitCode==0:
   print "## GAMS completed successfully.\n## Reading results..."
   g = open(resultsfilename,"r")
   lineCount = 0
   # Read a file such as: 
   # flow('A','1')      500.00
   itemsLoaded = list()
   while True :
      line = g.readline()
      if not line: break
      lineCount += 1
      line2 = line.replace("("," ").replace("'"," ").replace(","," ").replace(")"," ")
      line2 = line2.strip()
      if line2=="": continue # skip blank lines
      items = line2.split()
      if len(items)==0: continue
      # get name and index
      name = items[0]
      index = items[1:-1]
      # convert any values into floats that we can
      for i in range(len(index)):
         try:
            index[i] = float(index[i])
         except:
            pass
      # get value
      try:
         value = float(items[-1])
      except:
         raise TypeError("When reading the results for data item '"+repr(name)+"', an error occurred in line "+repr(line)+" when converting the last item '"+repr(items[-1])+"' to a number. (Line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")
      # print name,index,value
      dataItem = SolverStudio.DataItems[name]
      try:
         if dataItem[tuple(index)]<> value:
            dataItem[tuple(index)]=value
      except:
         raise TypeError("When reading the results for data item "+repr(name)+", an error occurred in line "+repr(line)+" when assigning the value "+repr(items[-1])+" to index "+repr(tuple(index))+". (Line "+repr(lineCount)+" of file "+repr(resultsfilename)+")")            
      if not name in itemsLoaded:
         itemsLoaded.append(name)
   if len(itemsLoaded)==0:
      print "## No results were loaded into the sheet."
   else:
      print "## Results loaded for data items:", itemsLoaded
else:
   print "## GAMS did not complete (Error Code %d); no solution is available." % exitCode
print "## Done"