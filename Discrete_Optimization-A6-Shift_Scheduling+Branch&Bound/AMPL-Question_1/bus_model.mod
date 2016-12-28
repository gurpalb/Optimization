#data

param Nb_Requirements;	 		# Requirements given in data file
param From{1..Nb_Requirements};  	# start time 
param To{1..Nb_Requirements};    	# end time 
param UBC{1..Nb_Requirements};	    
param SBC{1..Nb_Requirements};	


#model
set SHIFTS;	# set of workshifts, indexed by start time
#set shift_type;
set shift_type;
set company;
#set company := UBC SBC;

#var x{SHIFTS};
var x{i in SHIFTS, j in company, k in shift_type} integer >=0;
# minimize Total : sum{ j in SHIFTS}  Agents_Starting_At[j];
minimize total_labor_hours : sum{ i in SHIFTS, j in company} (9*x[i,j,1]+9*x[i,j,3]+4*x[i,j,2]);	  #not sure if this works

subject to UBC_requirement {r in SHIFTS}:
sum { i in SHIFTS, j in company, k in shift_type: ((i>=5) and (i <= r) and (i > r-4) and (j=1) and (i<=24))
												or ((i>=5) and (i<=r-5) and (i>r-9) and (i<=24) and (j=1) and (k=1))
												or ((i>=5) and (i<=r-5) and (i>r-9) and (i<=24) and (j=2) and (k=3))} 
	x[i,j,k] 
   >=UBC[r-4];   # adjusting index 1 so that it matches timeslot 5-6am

subject to SBC_requirement {r in SHIFTS}:
sum { i in SHIFTS, j in company, k in shift_type: ((i>=5) and (i <= r) and (i > r-4) and (j=2) and (i<=24))
												or ((i>=5) and (i<=r-5) and (i>r-9) and (i<=24) and (j=2) and (k=1))
												or ((i>=5) and (i<=r-5) and (i>r-9) and (i<=24) and (j=1) and (k=3))} 
	x[i,j,k] 
   >=SBC[r-4];

subject to UBC_nd_shift :
sum{i in SHIFTS}
	(x[i,1,1]+x[i,1,3])
>=sum{i in SHIFTS, k in shift_type}
	(0.65*x[i,1,k]);

subject to SBC_nd_shift :
sum{i in SHIFTS}
	(x[i,2,1]+x[i,2,3])
>=sum{i in SHIFTS, k in shift_type}
	(0.65*x[i,2,k]);

 
