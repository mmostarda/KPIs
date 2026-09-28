close all
clear variables
clc

disp(' ')
disp('Executable generation - Started')


% Add current folder to path to ensure visibility
addpath(pwd);

t.start = tic;

% Compile the app
mcc -m IC_KPIs_App.mlapp

disp(' ')
t.end=toc(t.start);
disp(['Process duration: ' num2str(t.end, '%.1f') 's'])

disp(' ')
disp('Executable generation - Completed')
