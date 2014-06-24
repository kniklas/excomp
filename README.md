# excomp

Excel compare script - batch file that allows execution of Excel 2013 SPREADSHEETCOMPARE tool from command line with two files as argument.

### Installation

Download ```excomp.bat``` and place it either in folder from which you execute the script or configure Environment Variables PATH so it points to folder where you placed this file (see further part of this README for instructions how to set Environment Variables).

### Example usage

Simply execute from command line following:

```
excomp.bat Book1.xlsx Book2.xlsx
```

### Pre-conditions

Make sure you satisfy below pre-conditions that are required to run this tool.

#### Verify you have proper Microsoft Office software

Make sure your Microsoft Office installation has this tool (Office 2013 or greater), to verify this follow below steps:

1. Select 'Start'  and 'All Programs'
2. Select 'Microsoft 2013' and further 'Office 2013 Tools'
3. Within the last selection you should find application called: 'Spreadsheet Compare 2013'

#### Environment Variables setting

For this script to work correctly you have to add to system path folder with SPREADSHEETCOMPARE.EXE executable.
Usually the executable file is here: C:\Program Files (x86)\Microsoft Office\Office15\DCF\

1. Select 'Start' 
2. Right click on 'Computer', select 'Properties'
3. Select 'Advanced system settings' (note: administrator privileges is required!)
4. Make sure you are in 'Advanced' tab of system properties and choose 'Environment Variables'
5. Add to your PATH variable path you want to add separated by semicolon sign ';'
6. Confirm changes - note you need to restart your shell/command line to apply the changes in path variables

#### Access rights

This script creates a temporary text file that is later deleted by SPREADSHEETCOMPARE tool. If user that executes the script does not have privileges to create a new file, the script will fail.
