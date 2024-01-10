Dim url, xmlhttp, response, password
url = "https://example.com/api/password.json"
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
enteredPassword = InputBox("Please enter the password:")
If enteredPassword = password Then
    MsgBox "Password correct. Access granted!", vbInformation, "Authentication Successful"
Else
    MsgBox "Incorrect password. Access denied.", vbExclamation, "Authentication Failed"
End If
