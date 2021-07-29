Attribute VB_Name = "mod_RcdSys"
' PR : Hammer
' Start date : 2004/12/22
' Last Modify : 2004/12/22

Option Explicit

Private Const HWND_TOP = 0
Private Const HWND_BOTTOM = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOREDRAW = &H8

Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long

Public strCurPath As String ' �Ҧb��m�ؿ�
Public strINIfile As String  ' INI �ɮצW�� (������|+�ɦW)
Public strCrLf  As String ' Double vbCrlf
Public blnPasswordOK As Boolean ' Password OK Flag
Public strPassword As String ' Login Password
Public strCheckCode As String ' ChkCode

Private Sub Main() ' �Ұʵ{��
  On Error GoTo ChkErr
    If Not SysInitialize Then End
    With frmWellcome
        .Show
         Delay 2
        .Hide
    End With
    If Not Check_Startup_Password Then
        Unload frmWellcome
        End
    End If
    If Len(strPassword) > 0 Then
        frmLogin.Show
    Else
        frmMain.Show
    End If
    Unload frmWellcome
  Exit Sub
ChkErr:
    ErrHandle "mod_RcdSys", "Subroutine : Main"
End Sub

Private Function Check_Startup_Password() As Boolean ' �ˬd�O�_�ҥαK�X�n���t��
  On Error GoTo ChkErr
    strCheckCode = Decrypt(Read_From_INI("Permission", "ChkCode"))
    strPassword = Decrypt(Read_From_INI("Permission", "KeyWord"))
    If strCheckCode = "ChouYinDer" Then
        If Len(strPassword) > 0 Then
            Check_Startup_Password = True
        Else
            MsgBox "�t�ΰѼ��� [RCDpara.ini] �K�X��Q�H�y��L , �t�ΧY�N���� ! " & strCrLf & "�Ь��}�լ�ޤ��q .. ����!"
            Check_Startup_Password = False
        End If
    Else
        If IsDate(strCheckCode) Then
            If Len(strPassword) > 0 Then
                MsgBox "�t�ΰѼ��� [RCDpara.ini] �K�X��Q�H�y��L , �t�ΧY�N���� ! " & strCrLf & "�Ь��}�լ�ޤ��q .. ����!"
                Check_Startup_Password = False
            Else
                Check_Startup_Password = True
            End If
        Else
            MsgBox "�t�ΰѼ��� [RCDpara.ini] �K�X�Ұʶ}���Q�H�y��L , �t�ΧY�N���� ! " & strCrLf & "�Ь��}�լ�ޤ��q .. ����!"
            Check_Startup_Password = False
        End If
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_RcdSys", "Function : Check_Startup_Password"
End Function

Private Function SysInitialize() As Boolean
  On Error GoTo ChkErr
    strCrLf = vbCrLf & vbCrLf
    strCurPath = App.Path & "\"
    strINIfile = strCurPath & "RCDpara.ini"
    If Len(Dir(strINIfile)) = 0 Then
        SysInitialize = False
        MsgBox "�t�ΰѼ��� [RCDpara.ini] ���s�b, �t�αN�L�k�Ұ� ! " & strCrLf & "�Ь��}�լ�ޤ��q .. ����!", vbInformation, "�T��"
    Else
        SysInitialize = True
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_RcdSys", "Function : SysInitialize"
End Function

Public Function Read_From_INI(Section As String, key As String) As String ' INI�ɸ��Ū��
  On Error GoTo ChkErr
    Dim strValue As String
    Dim lngLen As Long
    strValue = String(1024, 0)
    lngLen = GetPrivateProfileString(Section, key, "", strValue, Len(strValue), strINIfile)
    Read_From_INI = Left(strValue, lngLen)
  Exit Function
ChkErr:
    ErrHandle "mod_RcdSys", "Function : Read_From_INI"
End Function

Public Function Write_2_INI(Section As String, key As String, KeyValue As String) As String ' ��Ʀ^�gINI��
  On Error GoTo ChkErr
    Write_2_INI = WritePrivateProfileString(Section, key, KeyValue, strINIfile)
  Exit Function
ChkErr:
    ErrHandle "mod_RcdSys", "Function : Write_2_INI"
End Function

Public Function Get_Computer_Name() As String ' ���o�q���W��
  On Error GoTo ChkErr
    Dim lngLen As Long
    Dim strName As String
'    lngLen = MAX_COMPUTERNAME_LENGTH + 1
    lngLen = 32
    strName = String(lngLen, Chr(0))
    GetComputerName strName, lngLen
    Get_Computer_Name = Left(strName, lngLen)
    If Len(Get_Computer_Name) = 0 Then Get_Computer_Name = CreateObject("WScript.Network").ComputerName
  Exit Function
ChkErr:
    ErrHandle "mod_RcdSys", "Function : Get_Computer_Name"
End Function

Public Function Get_Local_IP(Optional ByVal strHostName As String = "") As String ' ���o�q��IP
  On Error GoTo ChkErr
    Dim clsIP As New clsGetIP
    Get_Local_IP = clsIP.GetIPAddress(strHostName)
    If Len(Get_Local_IP) = 0 Then Get_Local_IP = "127.0.0.1"
    Set clsIP = Nothing
  Exit Function
ChkErr:
    ErrHandle "mod_RcdSys", "Function : Get_Local_IP"
End Function

Public Sub Form_On_Top(obj As Object) ' �]�w����̤W�h
  On Error Resume Next
    SetWindowPos obj.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub

Public Sub Form_No_Top(obj As Object) ' �٭��椣��̤W�h
  On Error Resume Next
    SetWindowPos obj.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub

Public Sub Delay(DelayTime As Single) ' �]�w����
  On Error Resume Next
    Dim varTimer As Variant
    varTimer = Timer
    Do Until Timer - varTimer > DelayTime
        DoEvents
    Loop
End Sub

Public Sub Rlx(ByRef var As Variant) ' �����ܼƩΪ���
  On Error Resume Next
    Select Case TypeName(var)
             Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Date", "Boolean"
                     var = 0
                     ZeroMemory var, Len(var)
             Case "Object", "Recordset", "Connection", "IOraDynaset", "IOraDatabase"
                     var.Close
             Case "String"
                     var = vbNullString
                     ZeroMemory var, Len(var)
             Case "Collection"
                     ZeroMemory var, var.Count
             Case "Error", "Empty", "Null", "Nothing", "����"
                     ZeroMemory var, Len(var)
             Case "ORADC"
                     var.Recordset.Close
                     Set var.Recordset = Nothing
             Case "IOraSession"

    End Select
    If IsArray(var) Then Erase var
    Set var = Nothing
End Sub

Public Sub ErrHandle(Optional FormName As String = "frmMain", Optional ProcedureName As String = "") ' ���~�x���B�z
    Dim strErr As String
    Dim ErrNum As Variant
    Dim ErrDesc As Variant
    Dim ErrSrc As Variant
    ErrNum = Err.Number:      ErrDesc = Err.Description:     ErrSrc = Err.Source
    
    On Error Resume Next
    Screen.MousePointer = vbDefault
    DoEvents
    Screen.ActiveForm.Enabled = True
    strErr = "�o�Ϳ��~�ɶ� : " & CStr(Now) & strCrLf & _
                "�o�Ϳ��~�M�� : " & CStr(ErrSrc) & strCrLf & _
                "�o�Ϳ��~��� : " & FormName & strCrLf & _
                "�o�Ϳ��~�{�� : " & ProcedureName & strCrLf & _
                "���~�N�X : " & CStr(ErrNum) & strCrLf & _
                "���~��] : " & CStr(ErrDesc)
    
    MsgBox strErr, vbOKOnly + vbInformation, "�t�ΰ�����~!"
End Sub


