REM Compare two excel files specified as parameters to this script
REM Example usage: excomp.bat Book1.xlsx Book2.xlsx
REM Make sure you add to system path folder with SPREADSHEETCOMPARE.EXE
REM C:\Program Files (x86)\Microsoft Office\Office15\DCF\
REM Author: Kamil Niklasinski
REM This script is provided under GNU license - see license file for details.

@ECHO OFF
@ECHO Comparing two files:
@ECHO 1: %1
@ECHO 2: %2

dir %1 /B /S > temp.txt
dir %2 /B /S >> temp.txt

SPREADSHEETCOMPARE temp.txt
