Dim xInfo
Dim xObj
xInfo=InputBox("�п�J DSN/UID/PWD ? " & VBcrlf & "�� : GICMIS/GICMIS/MAY", "�߰�")
IF xInfo <> "" THEN 
	Set xObj = CreateObject("Encryption.Password")
	xInfo=xObj.Encrypt(xInfo)
	CreateObject("Wscript.Shell").RegWrite "HKCU\xInfo",xInfo
'	Msgbox xObj.Decrypt(xInfo)
	Set xObj=Nothing
End If

