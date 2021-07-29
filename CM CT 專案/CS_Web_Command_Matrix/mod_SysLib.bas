Attribute VB_Name = "mod_SysLib"
Option Explicit

Private Declare Function GetComputerName Lib "Kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Declare Function GetTickCount Lib "Kernel32" () As Long
Private Const strCrLf = vbCrLf & vbCrLf
Public strErrorMessage As String

'Public Enum CmdCodeEnu
'    CM_Reg = 2
'    Reset_CM = 12
'    Ping_CM = 20
'    Clear_CPE_IP = 25
'    CM_Status_Query = 13
'    CM_Cus_Account = 24
'    HFC_Type_Query = 21
'    CM_Conn_Log = 22
'    CM_QA_Log = 23
'    CM_PC_Log = 26
'    CP_IP_Reg = 27
'End Enum

'Private cnObj As Object
'Private rsObj As Object

'Private strOwner As String
Public strOwner As String
Public intCompCode As Integer
Private strErr As String

Public Function CP_Str(strP1 As String, strP2 As String) As Boolean ' ����r��O�_�۵� (�����j�p�g)
  On Error GoTo ChkErr
    CP_Str = (LCase(strP1) = LCase(strP2))
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "CP_Str"
End Function

Public Function RPxx(strValue As String) As String
  On Error Resume Next
    strValue = Replace(strValue, Chr(0), "", 1)
    strValue = Trim(Replace(strValue, Chr(13), "", 1, , vbTextCompare))
    strValue = Trim(Replace(strValue, Chr(10), "", 1, , vbTextCompare))
    RPxx = Trim(Replace(strValue, Chr(9), "", 1, , vbTextCompare))
End Function

Public Function RightNow(ByRef cnObj As Object) As Date
  On Error GoTo ChkErr
    If Val(cnObj.Execute("SELECT SYSTIME FROM " & strOwner & "SO041 WHERE ROWNUM <=1 ORDER BY COMPCODE").GetString(2, , , , 1) & "") = 1 Then
        RightNow = cnObj.Execute("SELECT TO_CHAR(SYSDATE,'YYYY/MM/DD') FROM DUAL").GetString & ""
        RightNow = RPxx(CStr(RightNow)) & " " & RPxx(cnObj.Execute("SELECT TO_CHAR(SYSDATE,'HH24:MI:SS') FROM DUAL").GetString & "")
    Else
        RightNow = Now
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "RightNow"
End Function

Public Function Get_Cmd_Seq_No(ByRef cnObj As Object, strReserve As String) As String
  On Error GoTo ChkErr
'    V0.8 (�W�[��(4)�p�I����)
'    (4) �]CM G/W�|�̾�SO311.CmdSeqNo�����X���ǳB�z�A
'         �G��妸"����"�@�~�A���F�w���ɶ������n���ȥ~�A��cmdseqno�s�X�W�h��follow�U�C��hinsert ��SO311�K
'    A.  �L�U�w���ɶ������O�y������ (2�X���q�O +�t�Φ褸����@8�X + 8�Xs_cmdseqno�y����)
'    B.  �Ъ`�N�I����CM�����A��CmdSeqNo����������A�����w���ɶ�������s�X�C
'    C.  �p�G�󤵤�2006/12/28���]�w�A�w����2007/01/01�����A��cmdseqno���e��092007010199999999�C
    Get_Cmd_Seq_No = RPxx(cnObj.Execute("SELECT " & strOwner & "S_CMDSEQNO.NEXTVAL FROM DUAL").GetString & "")
    Get_Cmd_Seq_No = Format(Get_Cmd_Seq_No, "00000000")
    If Len(strReserve) > 0 Then
        Get_Cmd_Seq_No = Format(strReserve, "yyyymmdd") & Get_Cmd_Seq_No
    Else
        Get_Cmd_Seq_No = Format(RightDate(cnObj), "yyyymmdd") & Get_Cmd_Seq_No
    End If
    Get_Cmd_Seq_No = Format(intCompCode, "00") & Get_Cmd_Seq_No
'   Jacky: Format(RightDate, "yyyymmdd") & GetInvoiceNo3("CmdSeqNo")
'   yyyymmdd00000000 �~���+8�X�y���Ǹ�
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "Get_Cmd_Seq_No"
End Function

Public Function RightDate(ByRef cnObj As Object) As Date
  On Error GoTo ChkErr
    If Val(cnObj.Execute("SELECT SYSTIME FROM " & strOwner & "SO041 WHERE ROWNUM <=1 ORDER BY COMPCODE").GetString(2, , , , 1) & "") = 1 Then
        RightDate = RPxx(cnObj.Execute("SELECT TO_CHAR(SYSDATE,'YYYY/MM/DD') FROM DUAL").GetString & "")
    Else
        RightDate = Date
    End If
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "RightDate"
End Function

Public Function GetDTString(dtm As String) As String
  On Error GoTo ChkErr
    GetDTString = Format(dtm, "ee/mm/dd hh:mm:ss")
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "Function GetDTString"
End Function

Public Function Get_Computer_Name() As String ' ���o�q���W�� 2005/09/07 By Hammer
  On Error GoTo ChkErr
    Dim lngLen As Long
    Dim strName As String
    If Len(Get_Computer_Name) = 0 Then Get_Computer_Name = CreateObject("WScript.Network").ComputerName
    If Len(Get_Computer_Name) = 0 Then Get_Computer_Name = Environ("COMPUTERNAME")
    If Len(Get_Computer_Name) = 0 Then
        lngLen = 32
        strName = String(lngLen, Chr(0))
        GetComputerName strName, lngLen
        Get_Computer_Name = Left(strName, lngLen)
        Get_Computer_Name = Replace(Get_Computer_Name, Chr(0), "", 1)
    End If
    Get_Computer_Name = (Get_Computer_Name)
  Exit Function
ChkErr:
    ErrHandle "mod_SysLib", "Get_Computer_Name"
End Function

'Public Sub Rlx(ByRef var As Variant)
'  On Error Resume Next
'    Select Case TypeName(var)
'             Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Date", "Boolean"
'                     var = 0
'                     ZeroMemory var, Len(var)
'             Case "Object", "Recordset", "Connection", "IOraDynaset", "IOraDatabase"
'                     var.Close
'             Case "String"
'                     var = vbNullString
'                     ZeroMemory var, Len(var)
'             Case "Collection"
'                     ZeroMemory var, var.Count
'             Case "Error", "Empty", "Null", "Nothing", "����"
'                     ZeroMemory var, Len(var)
'             Case "ORADC"
'                     var.Recordset.Close
'                     Set var.Recordset = Nothing
'             Case "IOraSession"
'
'    End Select
'    If IsArray(var) Then Erase var
'    Set var = Nothing
'End Sub

Public Sub ErrHandle(FormName As String, ProcedureName As String)     ' ���~�x���B�z
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
    strErrorMessage = strErr
End Sub




