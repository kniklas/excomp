# excomp

Excel compare script - batch file that allows execution of Excel 2013 SPREADSHEETCOMPARE tool from command line with two files as argument.


### Example usage
Simply execute from command line following:

```
excomp.bat Book1.xlsx Book2.xlsx
```

### Pre-conditions

Make sure your Microsoft Office installation has this tool (Office 2013 or greater), to verify this follow below steps:

1. Select 'Start'  and 'All Programs'
2. Select 'Microsoft 2013' and further 'Office 2013 Tools'
3. Within the last selection you should find application called: 'Spreadsheet Compare 2013'

For this script to work correctly you have to add to system path folder with SPREADSHEETCOMPARE.EXE executable.
Usually the executable file is here: C:\Program Files (x86)\Microsoft Office\Office15\DCF\

1. Select 'Start' 
2. Right click on 'Computer', select 'Properties'
3. Select 'Advanced system settings' (note: administrator privileges is required!)
4. Make sure you are in 'Advanced' tab of system properties and choose 'Environment Variables'
5. Add to your PATH variable path you want to add separated by semicolon sign ';'
6. Confirm changes - note you need to restart your shell/command line to apply the changes in path variables
