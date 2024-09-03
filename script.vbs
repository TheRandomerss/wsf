' Check if script is running as admin
If Not WScript.Arguments.Named.Exists("elevated") Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
    WScript.Quit
End If

' Now running as admin, proceed with the script
Dim url, downloadPath, objXMLHTTP, objFSO, objFile, objShell
url = "http://yourserver.com/goService.exe"
downloadPath = "C:\path\to\save\goService.exe"

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

' Install and start the service
Set objShell = CreateObject("WScript.Shell")
objShell.Run "goService.exe install", 0, True
objShell.Run "goService.exe start", 0, True
