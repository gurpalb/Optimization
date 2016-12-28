clear all;clc;close all;

% density
n=0.5;

% create coordinates
[xx,yy,zz] = meshgrid(0:n:1,0:n:1,0:n:1);

% boundary condition 1
rr = xx + yy - 2*zz;
% boundary condition 2 
rr2 = xx;
% boundary condition 3
rr3 = yy;

% region
region3= rr<0 & rr2<= 1 & rr2 >= 0 & rr3<= 1 & rr3 >= 0;

p=patch(isosurface(region3,0.5))

set(p,'FaceColor','red','EdgeColor','none');
daspect([1,1,1])
view(3); axis equal
camlight 
lighting gouraud
grid on
