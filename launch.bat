@echo off
%1(start /min cmd.exe /c %0 :&exit)
python.exe ./main.py