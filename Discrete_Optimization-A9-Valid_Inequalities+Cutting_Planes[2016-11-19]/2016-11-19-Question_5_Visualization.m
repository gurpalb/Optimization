clear all; close all; clc
x = linspace(-0.5,1.5,200);
y = linspace(-0.5,1.5,200);
[X1,Y1] = meshgrid(x,y);

Z1 = 0*X1;

X1(X1 < 1) = NaN;
X1(X1 > 2) = NaN;
surf(X1,Y1,Z1), hold on
shading flat

Z2 = 1+0*X1;

X1(X1 < 1) = NaN;
X1(X1 > 2) = NaN;
surf(X1,Y1,Z2), hold on
shading flat

surf(1,0, 0.5), hold on
shading flat