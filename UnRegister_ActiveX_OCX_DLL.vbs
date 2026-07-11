''Keo ActiveX DLL or OCX tha vao
With CreateObject("Scripting.FileSystemObject")
   If WScript.Arguments.Count = 0 Then WScript.Quit
   Dim strFile: strFile = .GetFile(WScript.Arguments.Item(0)).Path
   Dim RegSvr:  RegSvr  = .GetSpecialFolder(0).Path & "\SysWOW64\regsvr32.exe" 
   If Not .FileExists(strFile) Then MsgBox "Error reading file " & strFile & vbLF & .GetParentFolderName(WScript.ScriptFullName): WScript.Quit
   If Not .FileExists(RegSvr)  Then RegSvr = .GetSpecialFolder(0).Path & "\System32\regsvr32.exe"    
   If WScript.Arguments.Count = 1 Then     
      CreateObject("Shell.Application").ShellExecute RegSvr, "/u " & Chr(34) & strFile & Chr(34), "", "runas", 1
   End if
End With