Attribute VB_Name = "mod_SysLib"
' PR : Hammer
' Start date : 2004/12/22
' Last Modify : 2005/08/01

Option Explicit

Private Const HWND_TOP = 0
Private Const HWND_BOTTOM = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOREDRAW = &H8

Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long

Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Private sConnType As String * 255

Public OnLine As Boolean

Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectoryEmpty Lib "shlwapi.dll" Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Long
Private Declare Function PathIsLFNFileSpec Lib "shlwapi.dll" Alias "PathIsLFNFileSpecA" (ByVal lpName As String) As Long
Private Declare Function PathIsNetworkPath Lib "shlwapi.dll" Alias "PathIsNetworkPathA" (ByVal pszPath As String) As Long
Private Declare Function PathIsPrefix Lib "shlwapi.dll" Alias "PathIsPrefixA" (ByVal pszPrefix As String, ByVal pszPath As String) As Long
Private Declare Function PathIsRelative Lib "shlwapi.dll" Alias "PathIsRelativeA" (ByVal pszPath As String) As Long
Private Declare Function PathIsRoot Lib "shlwapi.dll" Alias "PathIsRootA" (ByVal pszPath As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Private Const OFS_MAXPATHNAME = 128
'Const OF_CREATE = &H1000
'Const OF_READ = &H0
'Const OF_WRITE = &H1

'Private Type OFSTRUCT
'        cBytes As Byte
'        fFixedDisk As Byte
'        nErrCode As Integer
'        Reserved1 As Integer
'        Reserved2 As Integer
'        szPathName(OFS_MAXPATHNAME) As Byte
'End Type

'Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
'Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
'Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
'Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
'Private Declare Function CopyLZFile Lib "lz32" (ByVal n1 As Long, ByVal n2 As Long) As Long

'Const CREATE_NEW = 1
'Const OPEN_EXISTING = 3
'Const FILE_SHARE_WRITE = &H2
'Const GENERIC_WRITE = &H40000000

Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
'Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public strCurPath As String ' 所在位置目錄
Public strINIfile As String  ' INI 檔案名稱 (完整路徑+檔名)
Public strCrLf  As String ' Double vbCrlf
Public blnPasswordOK As Boolean ' Password OK Flag
Public strPassword As String ' Login Password
Public strCheckCode As String ' ChkCode

Public cnCOMM As New ADODB.Connection ' 共用資料區之 [連線物件]
Public rsVirtual As New ADODB.Recordset ' 虛擬記憶體資料錄 , 供 XML 檔案排序用

Private Sub Main() ' 啟動程序
  On Error GoTo ChkErr
    If Not SysInitialize Then End
    Dim lngRet As Long
    OnLine = (InternetGetConnectedStateEx(lngRet, sConnType, 254, 0) = 1)
    With frmWellcome
        .Show
         Load frmLogin
         Delay 1
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
    ErrHandle "mod_SysLib", "Subroutine : Main"
End Sub

Private Function SysInitialize() As Boolean ' 系統初始化
  On Error GoTo ChkErr
    strCrLf = vbCrLf & vbCrLf
    strCurPath = App.Path & "\"
    strINIfile = strCurPath & "PRGpara.ini"
    If Len(Dir(strINIfile)) = 0 Then
        SysInitialize = False
        MsgBox "系統參數檔 [PRGpara.ini] 不存在, 系統將無法啟動 ! " & strCrLf & "請洽開博科技公司 .. 謝謝!", vbInformation, "訊息"
    Else
        SysInitialize = True
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "Function : SysInitialize"
End Function

Private Function Check_Startup_Password() As Boolean ' 檢查是否啟用密碼登錄系統
  On Error GoTo ChkErr
    strCheckCode = Decrypt(Read_From_INI("Permission", "ChkCode"))
    strPassword = Decrypt(Read_From_INI("Permission", "KeyWord"))
    If strCheckCode = "ChouYinDer" Then
        If Len(strPassword) > 0 Then
            Check_Startup_Password = True
        Else
            MsgBox "系統參數檔 [PRGpara.ini] 密碼欄被人篡改過 , 系統即將關閉 ! " & strCrLf & "請洽開博科技公司 .. 謝謝!"
            Check_Startup_Password = False
        End If
    Else
        If IsDate(strCheckCode) Then
            If Len(strPassword) > 0 Then
                MsgBox "系統參數檔 [PRGpara.ini] 密碼欄被人篡改過 , 系統即將關閉 ! " & strCrLf & "請洽開博科技公司 .. 謝謝!"
                Check_Startup_Password = False
            Else
                Check_Startup_Password = True
            End If
        Else
            MsgBox "系統參數檔 [PRGpara.ini] 密碼啟動開關被人篡改過 , 系統即將關閉 ! " & strCrLf & "請洽開博科技公司 .. 謝謝!"
            Check_Startup_Password = False
        End If
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "Function : Check_Startup_Password"
End Function

Public Function Chk_HM_OK(ByRef objMsk As Object) As Boolean ' 檢核 24HH:MM
  On Error GoTo ChkErr
    Dim strValue As String
    Chk_HM_OK = True
    With objMsk
        strValue = .ClipText
        If Val(Left(strValue, 2)) > 23 Then
            MsgBox "[時段定義] 欄位為 24 小時制, 時間不得超過 23 ! 請修正 !", vbInformation, "訊息"
            Chk_HM_OK = False
            .SelStart = 0
            .SelLength = 2
        End If
        If Val(Right(strValue, 2)) > 59 Then
            MsgBox "[時段定義] 欄位裡的 [分] 輸入不正確 ! 大於 59 請修正 !", vbInformation, "訊息"
            Chk_HM_OK = False
            .SelStart = 3
            .SelLength = 4
        End If
    End With
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "Function : Chk_HM_OK"
End Function

Public Function SecToHMS(Sec As Single) As String ' 將秒轉成時分秒字串
  On Error GoTo ChkErr
    Dim strHH As String, strMM As String, strSS As String
    Sec = CLng(Sec)
    If Sec > 60 Then
        strMM = Sec \ 60
        strSS = Sec Mod 60
        SecToHMS = strMM & " 分 " & IIf(strSS = 0, "", strSS & " 秒 ")
        If strMM > 60 Then
            strHH = strMM \ 60
            strMM = strMM Mod 60
            SecToHMS = strHH & " 時 " & IIf(strMM = 0, strSS & " 秒 ", strMM & " 分 " & strSS & " 秒 ")
        End If
    Else
        SecToHMS = Sec & " 秒 "
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "Function : SecToHMS"
End Function

Public Sub objSelAll(ByRef obj As Object) ' 設定欄位值全選
   On Error Resume Next
    With obj
        .SelStart = 0
        .SelLength = Len(.Text) + 6
    End With
End Sub

Public Function Read_From_INI(Section As String, key As String) As String ' INI檔資料讀取
  On Error GoTo ChkErr
    Dim strValue As String
    Dim lngLen As Long
    strValue = String(1024, 0)
    lngLen = GetPrivateProfileString(Section, key, "", strValue, Len(strValue), strINIfile)
    Read_From_INI = Left(strValue, lngLen)
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "Function : Read_From_INI"
End Function

Public Function Write_2_INI(Section As String, key As String, KeyValue As String) As String ' 資料回寫INI檔
  On Error GoTo ChkErr
    Write_2_INI = WritePrivateProfileString(Section, key, KeyValue, strINIfile)
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "Function : Write_2_INI"
End Function

Public Function File_ReadOnly(strFileName As String) As Boolean ' 檢核檔案是否唯讀
  On Error GoTo ChkErr
    File_ReadOnly = CBool(GetAttr(strFileName) And vbReadOnly)
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "File_ReadOnly"
End Function

' 設定唯讀檔案為一般 ( Normal ) 模式 ( 非系統 , 非隱藏 , 非唯讀 )
Public Sub Set_FileRead(strFileName As String)
  On Error GoTo ChkErr
     SetAttr strFileName, vbNormal
  Exit Sub
ChkErr:
    ErrHandle "mod_SysLib", "Set_FileRead"
End Sub

' 檢核檔案是否獨佔
Public Function Chk_File_Exclusive(strFileName As String) As Boolean
  On Error Resume Next
    Open strFileName For Binary As #256
    Chk_File_Exclusive = (Err.Number <> 0)
    Close #256
End Function

' 檢核路徑是否為資料夾
Public Function DirExists(strPath As String) As Boolean
  On Error GoTo ChkErr
    DirExists = CBool(PathIsDirectory(strPath))
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "DirExists"
End Function

' 檢核是否為空資料夾 ( 目錄中無檔案 )
Public Function EmptyDir(strPath As String) As Boolean
  On Error GoTo ChkErr
    EmptyDir = CBool(PathIsDirectoryEmpty(strPath))
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "EmptyDir"
End Function

Public Function Get_Computer_Name() As String ' 取得電腦名稱
  On Error GoTo ChkErr
    Dim lngLen As Long
    Dim StrName As String
    lngLen = 32
    StrName = String(lngLen, Chr(0))
    GetComputerName StrName, lngLen
    Get_Computer_Name = Left(StrName, lngLen)
    If Len(Get_Computer_Name) = 0 Then Get_Computer_Name = Environ("COMPUTERNAME")
    If Len(Get_Computer_Name) = 0 Then Get_Computer_Name = CreateObject("WScript.Network").ComputerName
    Get_Computer_Name = Replace(Get_Computer_Name, Chr(0), "", 1)
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "Get_Computer_Name"
End Function

Public Function CP_Str(strP1 As String, strP2 As String) As Boolean ' 比較字串是否相等 (不分大小寫)
  On Error GoTo ChkErr
    CP_Str = (LCase(strP1) = LCase(strP2))
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "CP_Str"
End Function

Public Sub Form_On_Top(obj As Object) ' 設定表單於最上層
  On Error Resume Next
    SetWindowPos obj.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub

Public Sub Form_No_Top(obj As Object) ' 還原表單不於最上層
  On Error Resume Next
    SetWindowPos obj.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub

Public Sub Delay(DelayTime As Single) ' 設定延遲
  On Error Resume Next
    Dim varTimer As Variant
    varTimer = Timer
    Do Until Timer - varTimer > DelayTime
        DoEvents
    Loop
End Sub

Public Sub Rlx(ByRef var As Variant) ' 釋放變數或物件
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
                     ZeroMemory var, var.count
             Case "Error", "Empty", "Null", "Nothing", "未知"
                     ZeroMemory var, Len(var)
             Case "ORADC"
                     var.Recordset.Close
                     Set var.Recordset = Nothing
             Case "IOraSession"

    End Select
    If IsArray(var) Then Erase var
    Set var = Nothing
End Sub

Public Sub ErrHandle(Optional FormName As String = "frmMain", Optional ProcedureName As String = "") ' 錯誤掌控處理
    Dim strErr As String
    Dim ErrNum As Variant
    Dim ErrDesc As Variant
    Dim ErrSrc As Variant
    ErrNum = Err.Number:      ErrDesc = Err.Description:     ErrSrc = Err.Source
    
    On Error Resume Next
    Screen.MousePointer = vbDefault
    DoEvents
    Screen.ActiveForm.Enabled = True
    strErr = "發生錯誤時間 : " & CStr(Now) & strCrLf & _
                "發生錯誤專案 : " & CStr(ErrSrc) & strCrLf & _
                "發生錯誤表單 : " & FormName & strCrLf & _
                "發生錯誤程序 : " & ProcedureName & strCrLf & _
                "錯誤代碼 : " & CStr(ErrNum) & strCrLf & _
                "錯誤原因 : " & CStr(ErrDesc)
    
    MsgBox strErr, vbOKOnly + vbInformation, "系統執行錯誤!"
End Sub


