Attribute VB_Name = "mduVodTools"
Option Explicit
Public strRecordProcedureFile As String
Public objFileSystem As Scripting.FileSystemObject
Public objRecordWrite As TextStream
Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Function ReadGICMIS1(strKey As String) As String
  On Error GoTo chkErr
    Dim strINI As String
    strINI = GaryGi(10)
    If strINI = "" Then
        ReadGICMIS1 = Empty
        Exit Function
    End If
    ReadGICMIS1 = ReadFROMINI(strINI, "GICMISPath", strKey)
  Exit Function
chkErr:
    Call ErrSub(FormName, "Function ReadGICMIS1")
End Function
Public Function ReadFROMINI(IniFileName As String, Section As String, Key As String, Optional DecryptFlag As Boolean) As String
  On Error GoTo chkErr
    If Not ChkDirExist(IniFileName) Then MsgBox "INI之檔名或路徑不正確!!", vbInformation, "訊息..": Exit Function
    Dim s As String, length As Long
    s = String(1024, 0)
    length = GetPrivateProfileString(Section, Key, "", s, Len(s), IniFileName)
    ReadFROMINI = IIf(DecryptFlag, Decrypt(Left(s, length)), Left(s, length))
  Exit Function
chkErr:
    Call ErrSub(FormName, "Function ReadFROMINI")
End Function
'判斷是否要記錄LOG
Public Function IsRecordProcedure() As Boolean
  On Error GoTo chkErr
    Dim aFile As String
    gErrLogPath = GaryGi(10)
    
    If gErrLogPath = Empty Then
        IsRecordProcedure = False
        Exit Function
        'gErrLogPath = ReadGICMIS1("ErrLogPath")
    End If
    
    aFile = gErrLogPath
    If Right(aFile, 1) <> "\" Then
        aFile = gErrLogPath & "\"
    End If
    'aFile = aFile & "VOD_Record\"
    aFile = aFile & "VODRecord_Kin.TXT"
    'aFile = aFile & GetGUID & "_" & Format(Now, "yyyymmddhhmmss") & ".TXT"
    If Dir(aFile) = "" Then
        IsRecordProcedure = False
        Exit Function
    End If
    Set objFileSystem = New Scripting.FileSystemObject
    strRecordProcedureFile = gErrLogPath
    
    If Right(strRecordProcedureFile, 1) <> "\" Then
        strRecordProcedureFile = strRecordProcedureFile & "\"
    End If
    'strRecordProcedureFile = strRecordProcedureFile & "VOD_Record\"
    If Not objFileSystem.FolderExists(strRecordProcedureFile) Then
        objFileSystem.CreateFolder (strRecordProcedureFile)
    End If
    strRecordProcedureFile = strRecordProcedureFile & "VOD_" & Format(Now, "yyyymmddhhmmss") & "_" & _
            GetGUID & ".TXT"
    
    Set objRecordWrite = objFileSystem.CreateTextFile(strRecordProcedureFile, True)
    IsRecordProcedure = True
    Exit Function
chkErr:
  IsRecordProcedure = False
End Function
'將程式訊息寫入記錄檔
Public Sub WriteRecordVodProcedure(ByVal aMsg As String)
  On Error Resume Next
    If Not blnRecordProcedure Then Exit Sub
    Dim strMsg As String
    strMsg = Now & " > " & aMsg
    objRecordWrite.WriteLine strMsg
    
End Sub








Public Function GetGUID() As String
  On Error Resume Next
    Dim udtGUID As GUID
    If (CoCreateGuid(udtGUID) = 0) Then
        GetGUID = _
        String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
        String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
        String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
        IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
        IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
        IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
        IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
        IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
        IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
        IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
        IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If
End Function


