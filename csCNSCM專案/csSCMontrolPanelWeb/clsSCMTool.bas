Attribute VB_Name = "clsSCMTool"
Option Explicit
Public objFileSystem As New Scripting.FileSystemObject
Public Const strRecordName = "SCM_RecordLog.TXT"
Public strRecordProcedureFile As String
Public objRecordWrite As Scripting.TextStream
Public blnRecordProcedure As Boolean
Public frmProcessOk As Boolean
Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long
Public Const strMoveDir As String = "SucessSCM"
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Function ReadGICMIS1(strKey As String) As String
  On Error GoTo ChkErr
    Dim strINI As String
    strINI = GaryGi(10)
    If strINI = "" Then
        ReadGICMIS1 = Empty
        Exit Function
    End If
    
    ReadGICMIS1 = ReadFROMINI(strINI, "GICMISPath", strKey)
  Exit Function
ChkErr:
    Call ErrSub("clsSCMTool", "Function ReadGICMIS1")
End Function
Public Function ReadFROMINI(IniFileName As String, Section As String, Key As String, Optional DecryptFlag As Boolean) As String
  On Error GoTo ChkErr
    If Not ChkDirExist(IniFileName) Then MsgBox "INI之檔名或路徑不正確!!", vbInformation, "訊息..": Exit Function
    Dim s As String, length As Long
    s = String(1024, 0)
    length = GetPrivateProfileString(Section, Key, "", s, Len(s), IniFileName)
    ReadFROMINI = IIf(DecryptFlag, Decrypt(Left(s, length)), Left(s, length))
  Exit Function
ChkErr:
    Call ErrSub("clsSCMTool", "Function ReadFROMINI")
End Function


'判斷是否要記錄LOG
Public Function IsRecordProcedure() As Boolean
  On Error GoTo ChkErr
    Dim aFile As String
    Dim aPath As String
    If gErrLogPath = Empty Then
        If GaryGi(10) = Empty Then
            gErrLogPath = App.Path
        Else
            gErrLogPath = ReadGICMIS1("ErrLogPath")
        End If
    End If
    If gErrLogPath = Empty Then
        IsRecordProcedure = False
        Exit Function
        'gErrLogPath = ReadGICMIS1("ErrLogPath")
    End If
    aPath = objFileSystem.GetFolder(gErrLogPath)
    
    If Right(aPath, 1) <> "\" Then
        aPath = aPath & "\"
    End If
    'aFile = aFile & "VOD_Record\"
    aFile = aPath & strRecordName
    'aFile = aFile & GetGUID & "_" & Format(Now, "yyyymmddhhmmss") & ".TXT"
    If Dir(aFile) = "" Then
        IsRecordProcedure = False
        Exit Function
    End If
    'Set objFileSystem = New Scripting.FileSystemObject
    strRecordProcedureFile = aPath
    
    If Right(strRecordProcedureFile, 1) <> "\" Then
        strRecordProcedureFile = strRecordProcedureFile & "\"
    End If
    If Not objFileSystem.FolderExists(strRecordProcedureFile) Then
        objFileSystem.CreateFolder (strRecordProcedureFile)
    End If
    strRecordProcedureFile = strRecordProcedureFile & "SCM_" & Format(Now, "yyyymmddhhmmss") & "_" & _
            GetGUID & ".TXT"
    
    Set objRecordWrite = objFileSystem.CreateTextFile(strRecordProcedureFile, True)
    IsRecordProcedure = True
    Exit Function
ChkErr:
  IsRecordProcedure = False
End Function

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

'將程式訊息寫入記錄檔
Public Sub WriteRecordVodProcedure(ByVal aMsg As String, Optional ByVal aBlnMoveDir As Boolean = False)
  On Error Resume Next
    If Not blnRecordProcedure Then Exit Sub
    Dim strMsg As String
    strMsg = Now & " > " & aMsg
    objRecordWrite.WriteLine strMsg
    If aBlnMoveDir Then
        
        objRecordWrite.Close
        Dim aMoveFolder As String
        Dim aMoveFileName As String
        aMoveFolder = ReadGICMIS1("ErrLogPath")
        If Right(aMoveFolder, 1) <> "\" Then
            aMoveFolder = aMoveFolder & "\"
        End If
        aMoveFileName = objFileSystem.GetFileName(strRecordProcedureFile)
        If Not objFileSystem.FolderExists(aMoveFolder & strMoveDir) Then
            objFileSystem.CreateFolder aMoveFolder & strMoveDir
            
        End If
        objFileSystem.MoveFile strRecordProcedureFile, aMoveFolder & strMoveDir & "\" & aMoveFileName
    End If
    
End Sub
