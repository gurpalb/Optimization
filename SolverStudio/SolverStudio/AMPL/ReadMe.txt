This code has been repeatedly updated to handle 'edge' cases, including:

** Index value substitutions used by AMPL to reduce output width
Eg:
PC_Flow [*,*]
# $4 = 'Customer 2'
# $6 = 'Customer 6'
:         'Customer 1' $4 
'Plant 1'     107.9 2

** Code was changed from using csv.reader to shlex.split()
 -  csv.reader does not handle the following line appropriately:
"Quantity [*,'EXTRA 05',*]"
as it partitions the space inside the quotations, as it is only able to identify quotation marks when they appear directly after a space. The best solution I have found has been using shlex.split(), this produces the correct partition, but removes the internal quotation marks.

See below for the old CSVParse code.

Change  "display" to "_display"
AJM changed the code to convert all the display XXX > Sheet; to _display XXX > Sheet; 
_display is the safest most-computer-friendly way to output values, and the one we can read back most reliably
This also works around problems with long names for indices which cause unusual wrapping (to do with line wrapping, whcih can be fixed by:  option display_width 80;). Specifically,

I have attached an example where SheetData1.data uses a 77 character index and works normally, and SheetData2.data uses a 78 character index and results in an odd difference.

For 77 characters:

DemandConstr [*] :=
1 2
4 2

For 78 characters:

DemandConstr [*] :=1 2
4 2


-----------


Old CSV Parse Code
def CSVParse(line):
	#
	#	Deals with differing quote characters
	#
	try:
		singleQuote = line.index("'")
	except: singleQuote = 100
	try:
		doubleQuote = line.index('"')
	except: doubleQuote = 101
	if singleQuote < doubleQuote:
		items = csv.reader([line], delimiter=' ', skipinitialspace=True, quotechar="'").next()
	else:
		items = csv.reader([line], delimiter=' ', skipinitialspace=True, quotechar='"').next()
	return items

