Attribute VB_Name = "mod_Common"
Option Explicit

' ����O����
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)
Public Const strCrLf = vbCrLf & vbCrLf

' �s�� INI ����
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'Public objGcn As Object
'Public cn As Object
'Public objCn As Object
'Public strCn As String

Public Enum ChangeDB
    giAccessDb = 0 'Access Table
    giOracle = 1    'Oracle Table
End Enum

Public Enum enumFldType ' ��ƫ��A�C�|��
    FldNum = 0
    FldChar = 1
    FldDate = 2
    FldTime = 3
End Enum

Public Enum giNargaProcess
    intSTBCreDold = 0
    intOrdPPVPro = 1
    intCancelPPVPro = 2
    intPPVRight = 3
    intIPPVRight = 4
End Enum

Public Enum giVType
    giLongV = 0
    giStringV = 1
    giDateV = 2
End Enum

Public strErr As String ' �r���ܼ� , �ΨӦs����~�T��
Public bytComp As Byte ' ���q�O
Public objCrypt As Object ' �[�K����
'Public str_INI_Comp As String ' INI ���ҩw�q�����x���q�O
'Public str_INI_DB_Link As String ' INI ���ҩw�q�����x��Ƴs��
Public garyGi(20) As String '--------- Dim Array --------------

Private Const strINIfile As String = "C:\SMS Web Gateway\ThreadPath.ini" ' INI �ɮ׸��|

Public Function Minus2str(ByVal str1 As String, ByVal str2 As String) As String
  On Error GoTo ChkErr
    Dim vAry As Variant, vElement As Variant
    Dim blnFlag As Boolean
    If str1 = str2 Then
        Minus2str = ""
    Else
        blnFlag = InStr(1, str1, "'") > 0
        If blnFlag Then
            str1 = Replace(str1, "'", "", 1)
            str2 = Replace(str2, "'", "", 1)
        End If
        str1 = str1 & ","
        vAry = Split(str2, ",")
        For Each vElement In vAry
            str1 = Replace(str1, vElement & ",", "", 1, , 1)
        Next
        Minus2str = Left(str1, Len(str1) - 1)
        If blnFlag Then Minus2str = "'" & Replace(Minus2str, ",", "','", 1) & "'"
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_Common", "Minus2str"
End Function

' �B�z��ƫ��O
Public Function PrcType(ByRef FldValue As Variant, _
                                        Optional DataType As enumFldType = FldChar, _
                                        Optional MustBe As Boolean = False) As String
  On Error GoTo ChkErr
    Select Case DataType
                Case FldNum
                         If MustBe Or Not IsNull(FldValue) Then
                            PrcType = Val(FldValue & Empty)
                         Else
                            PrcType = "NULL"
                         End If
                Case FldChar
                         If MustBe Or Not IsNull(FldValue) Then
                            PrcType = "'" & FldValue & Empty & "'"
                         Else
                            PrcType = "NULL"
                         End If
                Case FldDate
                         If MustBe Or Not IsNull(FldValue) Then
                            If InStr(1, FldValue, "/") Then
                                PrcType = "TO_DATE('" & FldValue & "','YYYY/MM/DD')"
                            Else
                                PrcType = "TO_DATE('" & FldValue & "','YYYYMMDD')"
                            End If
                         Else
                            PrcType = "NULL"
                         End If
                Case FldTime
                         If MustBe Or Not IsNull(FldValue) Then
                            If InStr(1, FldValue, "/") Then
                                PrcType = "TO_DATE('" & FldValue & "','YYYY/MM/DD HH24:MI:SS')"
                            Else
                                PrcType = "TO_DATE('" & FldValue & "','YYYYMMDD HH24MISS')"
                            End If
                         Else
                            PrcType = "NULL"
                         End If
    End Select
  Exit Function
ChkErr:
    ErrHandle "mod_Common", "PrcType"
End Function

'   �q��渹����: �褸�~��+ Sequence Object S_SO105E_OrderNo Ex: '2005030000001'
Public Function GetOrderNo(ByRef cn As Object, _
                                            ByVal strOwner As String, _
                                            ByRef strOrderNo As String, _
                                            ByRef strErrMsg As String) As Boolean
  On Error Resume Next
    strOrderNo = Trim(GetValue(cn, "SELECT " & strOwner & "S_SO105E_ORDERNO.NEXTVAL FROM DUAL") & "")
    GetOrderNo = (Len(strOrderNo) > 0)
    If GetOrderNo Then
        strOrderNo = Format(RightDate(cn), "YYYYMM") & Right("0000000" & strOrderNo, 7)
    Else
        strErrMsg = "-99,�q��s�� [S_SO105E_ORDERNO] Sequence ����o�Ϳ��~ ! "
    End If
End Function

Public Function GetBillNo(ByRef cn As Object, ByVal strOwner As String, ByRef strErrMsg As String) As String
  On Error Resume Next
    GetBillNo = "SELECT '" & Format(RightDate(cn), "YYYYMM") & "TC' || Ltrim(To_Char(" & strOwner & "S_SO033_CATV.NextVal, '0999999')) FROM DUAL"
    GetBillNo = GetValue(cn, GetBillNo)
    If Err.Number <> 0 Then
        Err.Clear
        GetBillNo = ""
        strErrMsg = "�L�k���o [��ڽs��] BillNo!"
    End If
End Function

Public Function GetOwner(ByRef cn As Object) As String ' ���o Owner
  On Error GoTo ChkErr
    If Val(bytComp) = 0 Then Exit Function
    GetOwner = RPxx(cn.Execute("SELECT TABLEOWNER FROM CD039" & Get_DB_Link(cn) & _
                                                    " WHERE CODENO=" & bytComp).GetString(2, 1, "", "", "") & "")
    If Len(GetOwner) > 0 Then GetOwner = GetOwner & "."
  Exit Function
ChkErr:
    ErrHandle "mod_Common", "GetOwner"
End Function

Private Function Read_From_INI(Section As String, Key As String) As String ' INI�ɸ��Ū��
  On Error GoTo ChkErr
    Dim strValue As String
    Dim lngLen As Long
    strValue = String(1024, 0)
    lngLen = GetPrivateProfileString(Section, Key, "", strValue, Len(strValue), strINIfile)
    Read_From_INI = Left(strValue, lngLen)
  Exit Function
ChkErr:
    ErrHandle "mod_Common", "Read_From_INI"
End Function

Public Function RightNow(ByRef objCN As Object) As Date
  On Error Resume Next
    RightNow = RPxx(objCN.Execute("SELECT TO_CHAR(SYSDATE,'YYYY/MM/DD HH24:MI:SS') FROM DUAL").GetString & "")
    If Err.Number <> 0 Then RightNow = Now
End Function

Public Function RightDate(ByRef objCN As Object) As Date
  On Error Resume Next
    RightDate = RPxx(objCN.Execute("SELECT TO_CHAR(SYSDATE,'YYYY/MM/DD') FROM DUAL").GetString & "")
    If Err.Number <> 0 Then RightDate = Date
End Function

Public Function NoZero(strVal As Variant, Optional ZeroFlag As Boolean = False) As Variant
  On Error GoTo ChkErr
    strVal = strVal & ""
    If Trim(strVal) = Empty Then
        NoZero = IIf(ZeroFlag, 0, Null)
    Else
        NoZero = strVal
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_Common", "NoZero"
End Function

Public Function GetDTString(dtm As String) As String
  On Error Resume Next
    GetDTString = Format(dtm, "ee/mm/dd hh:mm:ss")
End Function

'    Dim objCrypt As Object
'    Set objCrypt = CreateObject("Crypt32DLL.Crypt32")
'    Dim varResult As Variant
'    objCrypt.BValue "1,2,3,4,5", varResult, "CsMisk"
'    MsgBox varResult
'    objCrypt.DValue varResult, varResult, "CsMisk"
'    MsgBox varResult
'    ' [MutiThread]
'    ' C= Compnay
'    ' D= DB Link

'select * from gicmis3.cd001@myself

Public Function Get_DB_Link(ByRef cn As Object) As String ' ���o Database link
  On Error GoTo ChkErr
    If Val(bytComp) = 0 Then Exit Function
    Get_DB_Link = GetValue(cn, "SELECT LINK FROM COM.SO507 WHERE CODENO=" & bytComp)
    If Get_DB_Link <> Empty Then Get_DB_Link = "@" & Get_DB_Link
    Get_DB_Link = Get_DB_Link & " "
  Exit Function
ChkErr:
    ErrHandle "mod_Common", "Get_DB_Link(cn)"
End Function

'Public Function Get_DB_Link(cn)() As String ' ���o Database link
'  On Error GoTo ChkErr
'    Dim varResult As Variant
'    If Val(bytComp) = 0 Then Exit Function
'    If Dir(strINIfile) = Empty Then Exit Function
'    If objCrypt Is Nothing Then Set objCrypt = CreateObject("Crypt32DLL.Crypt32")
'    If str_INI_Comp = Empty Then
'        objCrypt.DValue Read_From_INI("MutiThread", "C"), varResult, "CsMisk"
'        str_INI_Comp = "," & CStr(varResult) & ","
'    End If
'    If str_INI_DB_Link = Empty Then
'        objCrypt.DValue Read_From_INI("MutiThread", "D"), varResult, "CsMisk"
'        str_INI_DB_Link = CStr(varResult)
'    End If
'    If InStr(1, str_INI_Comp, "," & bytComp & ",") Then
'        Get_DB_Link(cn) = "@" & RPxx(str_INI_DB_Link) & " "
'    Else
'        Get_DB_Link(cn) = " "
'    End If
'  Exit Function
'ChkErr:
'    ErrHandle "mod_Common", "Get_DB_Link(cn)"
'End Function

''---- CursorTypeEnum Values ----
'Const adOpenForwardOnly = 0
'Const adOpenKeyset = 1
'Const adOpenDynamic = 2
'Const adOpenStatic = 3
'
''---- LockTypeEnum Values ----
'Const adLockReadOnly = 1
'Const adLockPessimistic = 2
'Const adLockOptimistic = 3
'Const adLockBatchOptimistic = 4
'
''---- CursorLocationEnum Values ----
'Const adUseServer = 2
'Const adUseClient = 3

Public Function GetRS(ByRef cn As Object, ByRef rs As Object, strQry As String) As Boolean ' �ھڶǨӤ� SQL ����ƿ���
  On Error GoTo ChkErr
    Set rs = CreateObject("ADODB.Recordset")
    With rs
            If .State > 0 Then
                On Error Resume Next
                .Close
            End If
            On Error GoTo ChkErr
            .CursorLocation = 3
            .Open strQry, cn, 1, 3
            If .State <= 0 Then
                Set rs = Nothing
            Else
                GetRS = True
            End If
    End With
  Exit Function
ChkErr:
    ErrHandle "mod_Common", "GetRS"
    GetRS = False
End Function

Public Function GetValue(ByRef cn As Object, strQry As String) As String ' �ھڶǨӤ� SQL ���Y�@����
  On Error Resume Next
    GetValue = RPxx(cn.Execute(strQry).GetString(2, 1, "", "", ""))
  Exit Function
ChkErr:
    If Err.Number = 3021 Then ' EOF or BOF , �d�L��ƪ��B�z�ʧ@
        Err.Clear
        GetValue = Empty
    ElseIf Err.Number <> 0 Then
        ErrHandle "mod_Common", "GetValue"
    End If
End Function
 
Public Function ExpXML(ByRef rs As Object, Optional strOnlyRow As Boolean = False, Optional FilterSymbol As Boolean = True) As String ' �N Recordset �צ� XML ���榡
  On Error GoTo ChkErr
    If rs.EOF Then
        ExpXML = Empty
    Else
        Dim objPHdom As New PHxDOM ' �I�s���� XML ���������O
        ExpXML = objPHdom.RScvtXML(rs, "DataSet", "", "", "DataTable", "", "DataRow", strOnlyRow, FilterSymbol)
        Set objPHdom = Nothing
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_Common", "ExpXML"
End Function

Public Function Get_General_CN() As Object ' ���o Connection ���� (�����|�إ� Connection ����, �ós�u��Ʈw)
  On Error GoTo ChkErr
    Dim strINIfile As String ' INI �� (���|+�ɦW)
    Dim strCN As String ' �s�u�r��
    strINIfile = Environ("WINDIR") & IIf(IsNTserial, "\SYSTEM32\DBInfo.ini", "\SYSTEM\DBInfo.ini") ' ���o INI �ɸ��|+�ɦW
    If Len(Dir(strINIfile)) > 0 Then ' �P�_ INI �ɬO�_���T�s�b
        strCN = RPxx(Decrypt(GetAscFile(strINIfile), "CS")) ' �ѱK
        Set Get_General_CN = CreateObject("ADODB.Connection") ' �إ� ADODB �s�u����
        With Get_General_CN
            .CursorLocation = 3 ' �]�w���Ц�m�� adUseClient
            .Open strCN ' �}�Ҹ�Ʈw�s�u
        End With
    Else ' INI �ɤ��s�b���ʧ@
        strErr = "INI �t�ΰѼ��ɳ]�w�����T���ɮפ��s�b ! �Ь��}�իȤ�A�ȤH�� .. ���� !"
        MsgBox strErr, vbInformation, "�T��"
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_Common", "Get_General_CN"
End Function

'*************************************************************corey �s�W���s�u�Ѽ�*********************************************************************************************************
'�M���w��쥻��CONNECT POOL �קאּMSDAO Ū���}�ժ�INI�ɮ榡�A
'�m���ɮ׸��|+�ɦW  c:\windows\system32\dbinfo2.ini
'windown64�줸���|��c:\windows\WOW64\dbinfo2.ini
Public Function Get_General_CN2(ByVal intComp As Integer) As Object
On Error GoTo ChkErr
    Dim strINI2 As String ' INI �� (���|+�ɦW)
    Dim strCN2 As String ' �s�u�r��
    Dim strDSN As String
    Dim strUID As String
    Dim strPWD As String
    Dim cn As Object
    
      Set cn = CreateObject("ADODB.Connection") ' �إ� ADODB �s�u����
    strINI2 = Environ("WINDIR") & IIf(IsNTserial, "\SYSTEM32\DBInfo2.ini", "\SYSTEM\DBInfo2.ini") ' ���o INI �ɸ��|+

    If Len(Dir(strINI2)) > 0 Then ' �P�_ INI �ɬO�_���T�s�b
         strDSN = ReadFromINI(strINI2, CStr(intComp), "1", True)
         strUID = ReadFromINI(strINI2, CStr(intComp), "2", True)
         strPWD = ReadFromINI(strINI2, CStr(intComp), "3", True)
         strCN2 = "Provider=MSDAORA.1;" & _
                    "Password=" & strPWD & ";" & _
                    "User ID=" & strUID & ";" & _
                    "Data Source=" & strDSN & ";" & _
                    "Persist Security Info=True"
         cn.CursorLocation = 3 ' �]�w���Ц�m�� adUseClient
         cn.Open strCN2 ' �}�Ҹ�Ʈw�s�u
         If cn.State > 0 Then
            Set Get_General_CN2 = cn
         Else
            strErr = strDSN & " ( " & strUID & " ) �s�u�����\!"
         End If
    Else ' INI �ɤ��s�b���ʧ@
      strErr = "INI �t�ΰѼ��ɳ]�w�����T���ɮפ��s�b ! �Ь��}�իȤ�A�ȤH�� .. ���� !"
    End If
    
    Exit Function
ChkErr:
    ErrHandle "mod_Common", "Get_General_CN2"
End Function

Private Function ReadFromINI(IniFileName As String, Section As String, Key As String, Optional DecryptFlag As Boolean) As String
  On Error GoTo ChkErr

    If Not ChkDirExist(IniFileName) Then MsgBox "INI���ɦW�θ��|�����T!!", vbInformation, "�T��..": Exit Function
    Dim S As String, length As Long
    S = String(1024, 0)
    length = GetPrivateProfileString(Section, Key, "", S, Len(S), IniFileName)
    ReadFromINI = IIf(DecryptFlag, Decrypt2(Left(S, length)), Left(S, length))

  Exit Function
ChkErr:
    ErrHandle "mod_Common", "ReadFromINI"
End Function

Private Function Encrypt(OriginalString As String) As String
  On Error GoTo ChkErr

    Encrypt = CreateObject("DevPowerEncrypt.EnCrypt").Encrypt(OriginalString)

  Exit Function
ChkErr:
    ErrHandle "mod_Common", "Encrypt"
End Function

Private Function Decrypt2(EncryptionString As String) As String
  On Error GoTo ChkErr

    Decrypt2 = CreateObject("DevPowerEncrypt.EnCrypt").Decrypt(EncryptionString)

  Exit Function
ChkErr:
    ErrHandle "mod_Common", "Decrypt2"
End Function

Private Function ChkDirExist(FileName As String) As Boolean
  On Error GoTo ChkErr

    Dim varValue As Variant
    If Right(FileName, 1) = "\" Then FileName = Left(FileName, Len(FileName) - 1)
    FileName = UCase(FileName)
    varValue = Split(FileName, "\")
    If UCase(Dir(FileName, vbDirectory)) = varValue(UBound(varValue)) Then
        ChkDirExist = True
    Else
        If Dir(Trim(FileName)) = "" Then
            ChkDirExist = False
        Else
            ChkDirExist = True
        End If
    End If

  Exit Function
ChkErr:
    ErrHandle "mod_Common", "ChkDirExist"
End Function
'***********************************************************************************************************************************************************************

Private Function GetAscFile(ByVal strFileName As String) As String ' Ū����r��
  On Error GoTo ChkErr
    Dim FileBuffer() As Byte ' �ŧi�줸�}�C
    Dim FileHandle As Integer
    ReDim FileBuffer(1 To FileLen(strFileName)) As Byte
    FileHandle = FreeFile ' ���o�i�Τ��ɮ�ê�N�X
    Open strFileName For Binary As FileHandle ' �}�Ҥ�r�� , �� 2 �i��}��
    Get #FileHandle, 1, FileBuffer ' ���o�ɮ׸�Ƥ��e
    Close FileHandle ' �����ɮ�
    GetAscFile = FileBuffer ' �N�줸�}�C�ন��y��r�榡�r��
    GetAscFile = StrConv(GetAscFile, vbUnicode) ' �B�z Unicode
  Exit Function
ChkErr:
    ErrHandle "mod_Common", "GetAscFile"
End Function

Private Function IsNTserial() As Boolean ' �P�O�O�_�� NT �t�C�@�~�t�� (NT,Win2K,XP,2003��)
  On Error Resume Next
    IsNTserial = Len(Environ("OS")) > 0
End Function

Public Function RPxx(strValue As String) As String ' �h�� CHR(13) , CHR(10) �� CHR(9)
  On Error Resume Next
    strValue = Trim(Replace(strValue, Chr(13), "", 1, , vbTextCompare))
    strValue = Trim(Replace(strValue, Chr(10), "", 1, , vbTextCompare))
    RPxx = Trim(Replace(strValue, Chr(9), "", 1, , vbTextCompare))
End Function

Public Function InsertToOracle(ByRef cn As Object, _
                                                    ByVal strTable As String, _
                                                    rsSox As Object, _
                                                    Optional strWhere As String, _
                                                    Optional lngAffected As Long = 0, _
                                                    Optional blnAddNew As Boolean = True) As Boolean
  On Error GoTo ChkErr
  
    Dim intLoop As Integer
    Dim rsTmp As Object
    Dim strField As String
    Dim strValue As String
    Dim strSQL As String
    Dim strFullField As String
    
    If rsSox.State = 0 Then Exit Function
    
    If strWhere = "" Then strWhere = " WHERE 1 = 0 "
    
    If Not GetRS(cn, rsTmp, "SELECT * FROM " & strTable & " " & strWhere) Then Exit Function
    
    Call ChkHaveField(strFullField, "", rsSox)
    
    With rsTmp
        If .RecordCount = 0 And blnAddNew Then
            For intLoop = 0 To .Fields.Count - 1
                If ChkHaveField(strFullField, .Fields(intLoop).Name, rsSox) Then '�ˬd�O���������
                    If Len(rsSox(.Fields(intLoop).Name) & "") > 0 Then
                        strField = strField & "," & .Fields(intLoop).Name
                        If .Fields(intLoop).Type = 135 Then
                            strValue = strValue & "," & GetNullString(rsSox(.Fields(intLoop).Name).Value, giDateV, giOracle, True)
                        Else
                            strValue = strValue & "," & Replace(GetNullString(rsSox(.Fields(intLoop).Name).Value, giStringV, giOracle), ",", "'||chr(44)||'")
                        End If
                    End If
                End If
            Next
            strValue = "INSERT INTO " & strTable & " (" & Mid(strField, 2) & ") VALUES (" & Mid(strValue, 2) & ")"
        Else
            For intLoop = 0 To .Fields.Count - 1
                If ChkHaveField(strFullField, .Fields(intLoop).Name, rsSox) Then
                    If .Fields(intLoop).Type = 135 Then
                        strField = strField & "," & .Fields(intLoop).Name & "=" & GetNullString(rsSox(.Fields(intLoop).Name).Value, giDateV, giOracle, True)
                    Else
                        strField = strField & "," & .Fields(intLoop).Name & "=" & Replace(GetNullString(rsSox(.Fields(intLoop).Name).Value, giStringV, giOracle), ",", "'||chr(44)||'")
                    End If
                End If
            Next
            strValue = "UPDATE " & strTable & " SET " & Mid(strField, 2) & " " & strWhere
        End If
    End With
    
    cn.Execute strValue, lngAffected
    Call CloseRecordset(rsTmp)
    InsertToOracle = True
    
  Exit Function
ChkErr:
    ErrHandle "mod_Common", "InsertToOracle"
End Function

Public Function GetNullString(varValue As Variant, _
                                                Optional lngVType As giVType = giStringV, _
                                                Optional DBFlag As ChangeDB = giOracle, _
                                                Optional blnAddTime As Boolean = False) As Variant
  On Error GoTo ChkErr
    If lngVType = giDateV Then
        If Len(varValue & "") = 0 Then
            GetNullString = "NULL"
        Else
            If DBFlag = giOracle Then
                If blnAddTime Then
                   GetNullString = "To_Date('" & Format(varValue, "YYYYMMDDHHMMSS") & "','YYYYMMDDHH24MISS')"
                Else
                   GetNullString = "To_Date('" & Format(varValue, "YYYYMMDD") & "','YYYYMMDD')"
                End If
            Else
                If blnAddTime Then
                    GetNullString = "#" & Format(varValue, "YYYY/MM/DD HH:MM:SS") & "#"
                Else
                    GetNullString = "#" & Format(varValue, "YYYY/MM/DD") & "#"
                End If
            End If
        End If
    ElseIf lngVType = giLongV Then
        If Len(varValue & "") = 0 Then
            GetNullString = "NULL"
        Else
            GetNullString = varValue
        End If
    ElseIf lngVType = giStringV Then
        If Len(varValue & "") = 0 Then
            GetNullString = "NULL"
        Else
            varValue = Replace(varValue, "'", "''")
            GetNullString = "'" & varValue & "'"
        End If
    End If
  Exit Function
ChkErr:
  ErrHandle "mod_Common", "GetNullString"
End Function

Public Function ChkHaveField(ByRef strFullField As String, ByVal strField As String, rsSource As Object) As Boolean
  On Error GoTo ChkErr
    Dim intLoop As Integer
    If strFullField = Empty Then
        With rsSource
            For intLoop = 0 To .Fields.Count - 1
                strFullField = strFullField & ",'" & .Fields(intLoop).Name & "'"
            Next
        End With
        strFullField = Mid(strFullField, 2)
    Else
        If strField = Empty Then Exit Function
        If InStr(1, strFullField, "'" & strField & "'") = 0 Then Exit Function
    End If
    ChkHaveField = True
  Exit Function
ChkErr:
    ChkHaveField = False
    ErrHandle "mod_Common", "ChkHaveField"
End Function

'   ���O�覡�N�X, ���O�覡�W��: select CodeNo,Description from cd031 where RefNo=4 and StopFlag=0 and rownum=1
'#3237(2) �w�]���O�覡�N�X,���O�覡�W��-��Ū��select CodeNo,Description from cd031 where RefNo=4 and StopFlag=0 and Kind=2 and rownum=1 By Kin 2007/06/01
Public Function Get_Def_CM(ByRef cn As Object, ByVal strOwner As String, ByRef strCMcode As String, ByRef strCMname As String) As Boolean
  On Error Resume Next
    Dim rs As Object
    If GetRS(cn, rs, "SELECT CODENO,DESCRIPTION FROM " & strOwner & "CD031" & Get_DB_Link(cn) & _
                                " WHERE REFNO=4 AND STOPFLAG=0 AND Kind=2 AND ROWNUM=1") Then
        strCMcode = rs("CODENO") & ""
        strCMname = rs("DESCRIPTION") & ""
        Get_Def_CM = True
    Else
        strCMcode = ""
        strCMname = ""
    End If
End Function

'1�D �I�ں����A�{�W��O��select CodeNo,Description from cd032 where RefNo=4 and StopFlag=0 and rownum=1
'��CD032�ݤ������N�X�A�G�L�ϥΨ�ѦҼơA�G����W���CD032.CodeNo=4���C
' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'   �I�ں����N�X, �I�ں����W��: select CodeNo,Description from cd032 where RefNo=4 and StopFlag=0 and rownum=1
' XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Public Function Get_Def_PT(ByRef cn As Object, ByVal strOwner As String, ByRef strPTcode As String, ByRef strPTname As String) As Boolean
  On Error Resume Next
    Dim rs As Object
    If GetRS(cn, rs, "SELECT CODENO,DESCRIPTION FROM " & strOwner & "CD032" & Get_DB_Link(cn) & _
                                    " WHERE CODENO=4 AND STOPFLAG=0 AND ROWNUM=1") Then
        strPTcode = rs("CODENO") & ""
        strPTname = rs("DESCRIPTION") & ""
        Get_Def_PT = True
    Else
        strPTcode = ""
        strPTname = ""
    End If
End Function

Public Function RPC(strValue) As String
  On Error Resume Next
    If IsNull(strValue) Then RPC = "": Exit Function
    strValue = Replace(strValue, "&", "&amp;")
    strValue = Replace(strValue, "<", "&lt;")
    strValue = Replace(strValue, ">", "&gt;")
    strValue = Replace(strValue, """", "&quot;")
    strValue = Replace(strValue, "'", "&apos;")
    RPC = strValue
End Function

Public Sub CloseRecordset(rs As Object)
  On Error Resume Next
    With rs
        .CancelUpdate
        If .State > 0 Then .Close
    End With
    Set rs = Nothing
End Sub

Public Function Decrypt(ByVal Encrypted As String, Optional sEncKey As String) As String ' �ѱK
  On Error GoTo ChkErr
    Dim intWord As Integer
    Dim strLetter As String
    Dim strKeyNum As String
    Dim strEncNum As String
    Dim lngEncbuffer As Long
    Dim strDecrypted As String
    Dim lngMath As Long
    
    If sEncKey = "" Then sEncKey = "PowerHammer"
    ReDim encKEY(1 To Len(sEncKey))
    
    For intWord = 1 To Len(sEncKey)
        strKeyNum = Mid(sEncKey, intWord, 1)
        encKEY(intWord) = Asc(strKeyNum)
        If intWord = 1 Then lngMath = encKEY(intWord): GoTo Nextone
        If intWord >= 2 And encKEY(intWord) >= lngMath And encKEY(intWord) <= encKEY(intWord - 1) Then lngMath = lngMath - encKEY(intWord)
        If intWord >= 2 And encKEY(intWord) <= lngMath And encKEY(intWord) <= encKEY(intWord - 1) Then lngMath = lngMath - encKEY(intWord)
        If intWord >= 2 And encKEY(intWord) >= lngMath And encKEY(intWord) >= encKEY(intWord - 1) Then lngMath = lngMath + encKEY(intWord)
        If intWord >= 2 And encKEY(intWord) <= lngMath And encKEY(intWord) >= encKEY(intWord - 1) Then lngMath = lngMath + encKEY(intWord)
Nextone:
    Next intWord
    
    For intWord = 1 To Len(Encrypted)
        strLetter = Mid(Encrypted, intWord, 1)
        strEncNum = strEncNum & strLetter
        If strLetter = " " Then
            lngEncbuffer = CLng(Mid(strEncNum, 1, Len(strEncNum) - 1))
            strDecrypted = strDecrypted & Chr(lngEncbuffer - lngMath)
            strEncNum = ""
        End If
    Next intWord
    Decrypt = strDecrypted
    
  Exit Function
ChkErr:
    ErrHandle "mod_Common", "Decrypt"
End Function

Public Sub Rlx(ByRef var As Variant) ' �����ܼƩΪ���귽
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

Public Sub ErrHandle(Optional FormName As String = "CS_Web.Gateway", Optional ProcedureName As String = "") ' ���~�x���B�z
    Dim ErrNum As Variant, ErrDesc As Variant, ErrSrc As Variant
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSrc = Err.Source
    On Error Resume Next
    Screen.MousePointer = vbDefault
    DoEvents
    strErr = "�o�Ϳ��~�ɶ� : " & CStr(Now) & strCrLf & _
                "�o�Ϳ��~�M�� : " & CStr(ErrSrc) & strCrLf & _
                "�o�Ϳ��~��� : " & FormName & strCrLf & _
                "�o�Ϳ��~�{�� : " & ProcedureName & strCrLf & _
                "���~�N�X : " & CStr(ErrNum) & strCrLf & _
                "���~��] : " & CStr(ErrDesc)
    Debug.Print strErr
    strErr = Replace(strErr, ",", "..")
End Sub

