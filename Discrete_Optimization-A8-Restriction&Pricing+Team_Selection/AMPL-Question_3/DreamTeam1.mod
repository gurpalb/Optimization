set players := {1..20} ;
set attributes := {'rebound','assist','height','scoring','defense','maxplay'}; 

var p {players} binary;

param stats {players,attributes}; 


### OBJECTIVE 
maximize mean_scoring: (1/12)*( sum{i in players} p[i]*stats[i,'scoring']);

### CONSTRAINTS 

# 1) At least 3 playmakers (players 1-5) 
subject to PM: sum{i in 1..5} p[i] >= 3 ;

# 2) at least 4 shooting guards (players 4 -11) 
subject to SG: sum{i in 4..11} p[i] >= 4; 

# 3) at least 4 Forwards (players 9 - 16) 
subject to F: sum{i in 9..16} p[i] >= 4;

# 4) at least 3 centers (players 16-20) 
subject to C: sum{i in 16..20} p[i] >= 3; 

# 5) 2 NCAA (amateur) players selected
subject to NCAA: p[4] + p[8] + p[15] + p[20] >= 2; 

# 6) at most 1 of players 5 and 9 are selected
subject to p5_or_p9: p[5] + p[9] <= 1 ; 

# 7) players 2 and 9 are selected together, if at all
subject to p2_and_p9: p[2] = p[9] ; 

# 8) at most 3 players from the same professional team 
subject to nba1: p[1] + p[7] + p[12] + p[16] <= 3 ;
subject to nba2: p[2] + p[3] + p[9] + p[19] <= 3 ; 

# 9) mean rebound ability of team is atleast 7
subject to r: (1/12)*sum{i in players} p[i]*stats[i,'rebound'] >= 7 ; 

# 10) mean assist ability of team is atleast 7
subject to a: (1/12)*sum{i in players} p[i]*stats[i,'assist'] >= 7 ; 

# 11) mean defense ability of team is atleast 7
subject to d: (1/12)*sum{i in players} p[i]*stats[i,'defense'] >= 7 ; 

# 12) mean height of team is atleast 1.92 
subject to h: (1/12)*sum{i in players} p[i]*stats[i,'height'] >= 1.92 ;

# 13) exactly 12 people on the basketball team
subject to team_size: sum{i in players} p[i] = 12; 

 








