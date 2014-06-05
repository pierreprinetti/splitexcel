If WScript.Arguments.Count = 0 then
    WScript.Echo "Name the Excel file to open."
Else

Dim strFilename  
Dim objFSO  
Set objFSO = CreateObject("scripting.filesystemobject")  
strFilename = objFSO.GetAbsolutePathName(WScript.Arguments(0))  
If objFSO.fileexists(strFilename) Then  
  Call Writefile(strFilename)  
Else  
  wscript.echo "Error: could not find the file."  
End If  
Set objFSO = Nothing  

Sub Writefile(ByVal strFilename)  
Dim objExcel  
Dim objWB  
Dim objws  

Set objExcel = CreateObject("Excel.Application")  
Set objWB = objExcel.Workbooks.Open(strFilename)  

For Each objws In objWB.Sheets  
  objws.Copy  
  objExcel.ActiveWorkbook.SaveAs objWB.Path & "\" & objFSO.GetBaseName(strFilename) & "-" & objws.Name
  objExcel.ActiveWorkbook.Close False  
Next 

objWB.Close False  
objExcel.Quit  
Set objExcel = Nothing  
End Sub  
End If
