'##wavier
Dim Auth
Auth = MsgBox("By continuing, you agree all the conditions stated in https://github.com/lemonscripting/chromeExtensionBlocker/blob/main/LICENSE and https://github.com/lemonscripting/chromeExtensionBlocker/blob/main/README.md" & Extension_ID, vbYesNo, "@lemonscripting | chromeExtensionBlocker")
Select Case Auth     
    Case vbNo
        MsgBox "Execution Failed: Blocked by client"
        WScript.Quit
End Select

'##logging
Dim currentDate
currentDate = Now
formattedDate = "[" & Day(currentDate) & "/" & Month(currentDate) & "/" & Year(currentDate) & " " & FormatDateTime(currentDate, 4) & "]"
ForAppending = formattedDate
Set objFSO = CreateObject("Scripting.FileSystemObject")
scriptPath = WScript.ScriptFullName
scriptFolder = objFSO.GetParentFolderName(scriptPath)
filePath = objFSO.BuildPath(scriptFolder, "log.txt")
If objFSO.FileExists(filePath) Then
    Set objFile = objFSO.OpenTextFile(filePath, 8)
Else
    Set objFile = objFSO.CreateTextFile(filePath, True)
End If
objFile.WriteLine ForAppending
objFile.Close
'WScript.Echo "Text file updated successfully at: " & filePath
'###End Region

'##source control
Dim url, xmlhttp, response, password
'replace with json link
url = "https://example.json"
Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
xmlhttp.Open "GET", url, False
xmlhttp.setRequestHeader "Content-Type", "application/json"
xmlhttp.send
response = xmlhttp.responseText
Dim startPos, endPos
startPos = InStr(response, """password"": """) + Len("""password"": """)
endPos = InStr(startPos, response, """")
password = Mid(response, startPos, endPos - startPos)
Dim enteredPassword
enteredPassword = InputBox("Password Required:")
If enteredPassword = password Then
    MsgBox "Password correct. Access granted!", vbInformation, "Authentication Successful"
Else
    MsgBox "Incorrect password. Access denied.", vbExclamation, "Authentication Failed"
    WScript.Quit
End If

'##target extension
Dim Extension_ID
Extension_ID = InputBox("Enter target Extension ID:", "@lemonscripting | chromeExtensionBlocker")

'##gather required data
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")
Username = wshShell.ExpandEnvironmentStrings("%USERNAME%")
Dim Header
Header = "C:\Users\"
profileNum = 9
Do Until objFSO.FolderExists(Header & Username & "\AppData\Local\Google\Chrome\User Data\Profile " & profileNum)
    profileNum = profileNum - 1
Loop
'WScript.Echo "Chrome Profile Number: " & profileNum
Dim Root_Path
Root_Path = Header & Username & "\AppData\Local\Google\Chrome\User Data\Profile " & profileNum & "\Extensions\" & Extension_ID & "\"
'WScript.Echo Root_Path

'##get last modified file in folder
Function GetRecentFile(path)
  Dim fso, file
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set GetRecentFile = Nothing
  For Each file in fso.GetFolder(path).Files
    If GetRecentFile is Nothing Then
      Set GetRecentFile = file
    ElseIf file.DateLastModified > GetRecentFile.DateLastModified Then
      Set GetRecentFile = file
    End If
  Next
End Function
Function GetRecentFolder(path)
  Dim fso, folder
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set GetRecentFolder = Nothing
  For Each folder in fso.GetFolder(path).SubFolders
    If GetRecentFolder is Nothing Then
      Set GetRecentFolder = folder
    ElseIf folder.DateLastModified > GetRecentFolder.DateLastModified Then
      Set GetRecentFolder = folder
    End If
  Next
End Function
Dim recentFile
Set recentFile = GetRecentFolder(Root_Path)
If recentFile is Nothing Then
  WScript.Echo "target file has been deleted or moved"
Else
  'WScript.Echo "Recent file is " & recentFile.Name & " " & recentFile.DateLastModified
  Dim Trigger
  Trigger = Root_Path & CStr(recentFile.Name)
  'WScript.Echo Trigger
  'WScript.Echo recentFile.Name
End If

'##trigger server to post an update
If objFSO.FileExists(Trigger) Then
    Set objFile = objFSO.OpenTextFile(Trigger, 2, True)
    objFile.Close
    'WScript.Echo "Content deleted successfully from: " & Trigger
Else
    WScript.Echo "target file has been deleted or moved"
End If

'##main loop
Dim objFSO, objFolder, strBasePath, i
Dim characters
strBasePath = Root_Path
Set objFSO = CreateObject("Scripting.FileSystemObject")
characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
Do
    Set objFolder = Nothing
    On Error Resume Next
    Set objFolder = objFSO.GetFolder(strBasePath)
    On Error GoTo 0
    If Not objFolder Is Nothing Then
        RenameFolders objFolder
    End If
    WScript.Sleep 20
Loop
Sub RenameFolders(folder)
    Dim subFolder, strRandomName, i
    For Each subFolder In folder.SubFolders
        Do
            strRandomName = ""
            For i = 1 To 5
                strRandomName = strRandomName & Mid(characters, Int((Len(characters) * Rnd) + 1), 1)
            Next
        Loop While FolderExists(strBasePath & "\" & strRandomName)
        On Error Resume Next
        subFolder.Name = strRandomName
        If Err.Number <> 0 Then
            Err.Clear
        End If
        On Error GoTo 0
    Next
End Sub
Function FolderExists(folderPath)
    FolderExists = objFSO.FolderExists(folderPath)
End Function