Set objFSO = CreateObject("Scripting.FileSystemObject")
Set pathToTxt = objFSO.OpenTextFile(objFSO.GetParentFolderName(WScript.ScriptFullName) &"\path_pst.txt" , ForReading) 
 
Const ForReading = 1 
 
Dim arrFileLines()
i = 0 
Do Until objFile.AtEndOfStream 
Redim Preserve arrFileLines(i) 
arrFileLines(i) = objFile.ReadLine 
i = i + 1 
Loop 
objFile.Close 
 
'Then you can iterate it like this 
 
For Each strLine in arrFileLines 
WScript.Echo strLine 

Set objOutlook = CreateObject("Outlook.Application")
Set objNameSpace = objOutlook.GetNamespace("MAPI")
'Set oShell = CreateObject("WScript.Shell")
'oShell.Run "outlook.exe", 1, False
objNameSpace.AddStore strLine
Next
objFSO.Deletefile(objFSO.GetParentFolderName(WScript.ScriptFullName) &"\path_pst.txt" , true) 