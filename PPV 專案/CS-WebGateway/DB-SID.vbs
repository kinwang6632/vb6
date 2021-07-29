Dim xInfo
Dim xObj
xInfo=InputBox("½Ð¿é¤J DSN/UID/PWD ? " & VBcrlf & "¨Ò : GICMIS/GICMIS/MAY", "¸ß°Ý")
IF xInfo <> "" THEN 
	Set xObj = CreateObject("Encryption.Password")
	xInfo=xObj.Encrypt(xInfo)
	CreateObject("Wscript.Shell").RegWrite "HKCU\xInfo",xInfo
'	Msgbox xObj.Decrypt(xInfo)
	Set xObj=Nothing
End If

