VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmConvert 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "SO106A����ഫ"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   8565
   StartUpPosition =   2  '�ù�����
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   4830
      TabIndex        =   8
      Top             =   840
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FilterStop      =   -1  'True
   End
   Begin VB.TextBox txtConnect 
      Height          =   375
      Left            =   1170
      ScrollBars      =   1  '�������b
      TabIndex        =   7
      Text            =   "Provider=MSDAORA.1;Password=EMCNCC;User ID=EMCNCC;Data Source=RDKNET"
      Top             =   330
      Width           =   7065
   End
   Begin MSComctlLib.ProgressBar prbProcess 
      Height          =   465
      Left            =   120
      TabIndex        =   5
      Top             =   1410
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   820
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtCustId2 
      Alignment       =   1  '�a�k���
      Height          =   315
      Left            =   3090
      TabIndex        =   3
      Top             =   780
      Width           =   1455
   End
   Begin VB.TextBox txtCustId1 
      Alignment       =   1  '�a�k���
      Height          =   315
      Left            =   1140
      TabIndex        =   2
      Top             =   780
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "����"
      Height          =   405
      Left            =   150
      TabIndex        =   0
      Top             =   2010
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "�s�u�r��"
      Height          =   225
      Left            =   150
      TabIndex        =   6
      Top             =   390
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "~"
      Height          =   285
      Left            =   2700
      TabIndex        =   4
      Top             =   840
      Width           =   315
   End
   Begin VB.Label Label1 
      Caption         =   "�Ȥ�s��"
      Height          =   225
      Left            =   150
      TabIndex        =   1
      Top             =   870
      Width           =   945
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strCn As String
Private rsSO106 As New ADODB.Recordset
Private rsSO106A As New ADODB.Recordset
Private cn As New ADODB.Connection
Private strQry As String
Private intCompCode As Integer
Private strOwnerName As String
Private strCreateTime As String
Private strUpdTime As String
Private objFile As New Scripting.FileSystemObject
Private objLog As TextStream
Const gDTfmt = "EE/MM/DD HH:MM:SS"
Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    If txtCustId1.Text <> "" Then Call subChoose(strQry, "CustId>=" & txtCustId1)
    If txtCustId2.Text <> "" Then Call subChoose(strQry, "CustId<=" & txtCustId2)
    If gilCompCode.GetCodeNo <> "" Then Call subChoose(strQry, "CompCode=" & gilCompCode.GetCodeNo)
    Call subChoose(strQry, "ACHTNO is Not Null")
    Call subChoose(strQry, "ACHTDESC is Not Null")
    rsSO106.CursorLocation = adUseClient
    rsSO106.Open strQry, cn, adOpenKeyset, adLockOptimistic
    If rsSO106.RecordCount > 0 Then
        prbProcess.Min = 0
        prbProcess.Max = rsSO106.RecordCount
        rsSO106.MoveFirst
        cn.BeginTrans
        Do While Not rsSO106.EOF
            If IsNull(rsSO106("MasterId")) Then
                rsSO106("MasterId") = Get_MasterID_Seq
                rsSO106.Update
            End If
            With rsSO106
                If IsNull(.Fields("AuthorizeStatus")) Then
                    If IsNull(.Fields("StopDate")) And (.Fields("Stopflag") & "" <> "1") Then
                        If IsNull(.Fields("Senddate")) Then
                            Call ChkSO106A(0) '�ݴ��X
                        Else
                            Call ChkSO106A(3) '�w���X�������v���\
                        End If
                    End If
                Else
                    Select Case .Fields("AuthorizeStatus") & ""
                        Case "1"
                            If (Not IsNull(.Fields("Senddate"))) And (.Fields("Stopflag") & "" <> "1") And (IsNull(.Fields("StopDate"))) Then
                                Call ChkSO106A(1) '���v���\
                            ElseIf (Not IsNull(.Fields("SendDate")) And (.Fields("StopFlag") & "" = "1") And (Not IsNull(.Fields("StopDate")))) Then
                                Call ChkSO106A(4) '�������v�ݴ��X
                            End If
                        Case "2"
                            If (Not IsNull(.Fields("Senddate"))) And (.Fields("Stopflag") & "" <> "0") And (Not IsNull(.Fields("StopDate"))) Then
                                Call ChkSO106A(2) '�������v
                            End If
                    
                    End Select
                End If
            End With
            rsSO106.MoveNext
            prbProcess.Value = prbProcess.Value + 1
NextLoop:
        Loop
        cn.CommitTrans
    Else
        MsgBox "�d�L���!", vbExclamation, "�T��"
    End If
    MsgBox "����ഫ�����I", vbInformation, "�T��"
    On Error Resume Next
    rsSO106.Close
    rsSO106A.Close
  Exit Sub
ChkErr:
  objLog.WriteLine ("�Ƚs:" & rsSO106("CustId") & " �ഫ����")
  cn.RollbackTrans
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    strCn = "Provider=MSDAORA.1;Password=EMCNCC;User ID=EMCNCC;Data Source=RDKNET"
    cn.ConnectionString = strCn
    cn.CursorLocation = adUseClient
    cn.Open
    intCompCode = 5
    Call Getowner
    strQry = "Select RowId,A.* From " & strOwnerName & "SO106 A"
    strUpdTime = GetDTString(RightNow)
    strCreateTime = Format(RightNow, "yyyy/mm/dd hh:mm:ss")
    Set objLog = objFile.CreateTextFile(App.Path & "\ConverErr.Log", True)
    SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12
    gilCompCode.ListIndex = 1
    Exit Sub
ChkErr:
  MsgBox "Form_Load����", vbCritical, "�T��"
End Sub
Private Sub Getowner()
  On Error GoTo ChkErr
    Dim strQry As String
    Dim strOwner As String
    strQry = "Select TableOwner From CD039 Where CodeNo=" & intCompCode
    strOwner = cn.Execute(strQry)(0)
    If strOwner <> Empty Then
      strOwner = strOwner & "."
    End If
    strOwnerName = strOwner
    Exit Sub
ChkErr:
    MsgBox "���oOwnerName����", vbCritical, "�T��"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'  On Error Resume Next
'    rsSO106.Close
'    cn.Close
'    Set rsSO106 = Nothing
'    Set cn = Nothing
End Sub
Private Sub subChoose(sQry As String, ByVal strWhere As String)
  On Error GoTo ChkErr
    sQry = sQry & IIf(InStr(1, sQry, "Where") > 0, " And " & strWhere, " Where " & strWhere)
  Exit Sub
ChkErr:
    MsgBox "�����ɿ��~", vbCritical, "�T��"
End Sub
Private Function InsertSO106A(ByVal intType As Integer, ByRef strAry() As String) As Boolean
    Dim strSQL As String
    Dim i As Integer
    For i = LBound(strAry) To UBound(strAry)
        Select Case intType
            '�ݴ��X���
            Case 0
                strSQL = "Insert into " & strOwnerName & "SO106A " & _
                                "(MasterRowID,ACHTNO,CitemCodeStr,CitemNameStr," & _
                                "UpdEn,UpdTime,CreateTime,CreateEn,RecordType," & _
                                "AuthorizeStatus,ACHDesc,MasterId)" & _
                                " Values ('" & rsSO106("RowId") & "'," & _
                                          "'" & strAry(i, 0) & "'," & _
                                          IIf(strAry(i, 3) = Empty, "Null,", "'" & strAry(i, 3) & "',") & _
                                          IIf(strAry(i, 4) = Empty, "Null,", "'" & strAry(i, 4) & "',") & _
                                          "'CableSoft'," & _
                                          "'" & strUpdTime & "'," & _
                                          "TO_DATE('" & Format(rsSO106("AcceptTime"), "yyyy/mm/dd hh:mm:ss") & "','YYYY/MM/DD HH24:MI:SS')," & _
                                          "'CableSoft'," & _
                                          "0,4,'" & strAry(i, 1) & "'," & rsSO106("MasterId") & ")"
                cn.Execute strSQL
                '���v���\
            Case 1
                strSQL = "Insert into " & strOwnerName & "SO106A " & _
                                "(MasterRowID,ACHTNO,CitemCodeStr,CitemNameStr," & _
                                "UpdEn,UpdTime,CreateTime,CreateEn,RecordType," & _
                                "AuthorizeStatus,ACHDesc,MasterId)" & _
                                " Values ('" & rsSO106("RowId") & "'," & _
                                          "'" & strAry(i, 0) & "'," & _
                                          IIf(strAry(i, 3) = Empty, "Null,", "'" & strAry(i, 3) & "',") & _
                                          IIf(strAry(i, 4) = Empty, "Null,", "'" & strAry(i, 4) & "',") & _
                                          "'CableSoft'," & _
                                          "'" & strUpdTime & "'," & _
                                          "TO_DATE('" & Format(rsSO106("AcceptTime"), "yyyy/mm/dd hh:mm:ss") & "','YYYY/MM/DD HH24:MI:SS')," & _
                                          "'CableSoft'," & _
                                          "0,1,'" & strAry(i, 1) & "'," & rsSO106("MasterId") & ")"
                cn.Execute strSQL
                '�������v���
            Case 2
                strSQL = "Insert into " & strOwnerName & "SO106A " & _
                                "(MasterRowID,ACHTNO,CitemCodeStr,CitemNameStr," & _
                                "UpdEn,UpdTime,CreateTime,CreateEn,RecordType," & _
                                "AuthorizeStatus,ACHDesc,MasterId)" & _
                                " Values ('" & rsSO106("RowId") & "'," & _
                                          "'" & strAry(i, 0) & "'," & _
                                          IIf(strAry(i, 3) = Empty, "Null,", "'" & strAry(i, 3) & "',") & _
                                          IIf(strAry(i, 4) = Empty, "Null,", "'" & strAry(i, 4) & "',") & _
                                          "'CableSoft'," & _
                                          "'" & strUpdTime & "'," & _
                                          "TO_DATE('" & Format(rsSO106("AcceptTime"), "yyyy/mm/dd hh:mm:ss") & "','YYYY/MM/DD HH24:MI:SS')," & _
                                          "'CableSoft'," & _
                                          "1,2,'" & strAry(i, 1) & "'," & rsSO106("MasterId") & ")"
                cn.Execute strSQL
                '�w���X�������v���\
            Case 3
                strSQL = "Insert into " & strOwnerName & "SO106A " & _
                                "(MasterRowID,ACHTNO,CitemCodeStr,CitemNameStr," & _
                                "UpdEn,UpdTime,CreateTime,CreateEn,RecordType," & _
                                "AuthorizeStatus,ACHDesc,MasterId)" & _
                                " Values ('" & rsSO106("RowId") & "'," & _
                                          "'" & strAry(i, 0) & "'," & _
                                          IIf(strAry(i, 3) = Empty, "Null,", "'" & strAry(i, 3) & "',") & _
                                          IIf(strAry(i, 4) = Empty, "Null,", "'" & strAry(i, 4) & "',") & _
                                          "'CableSoft'," & _
                                          "'" & strUpdTime & "'," & _
                                          "TO_DATE('" & Format(rsSO106("AcceptTime"), "yyyy/mm/dd hh:mm:ss") & "','YYYY/MM/DD HH24:MI:SS')," & _
                                          "'CableSoft'," & _
                                          "0,NULL,'" & strAry(i, 1) & "'," & rsSO106("MasterId") & ")"
                cn.Execute strSQL
                '�������v�ݴ��X
            Case 4
                strSQL = "Insert into " & strOwnerName & "SO106A " & _
                                    "(MasterRowID,ACHTNO,CitemCodeStr,CitemNameStr," & _
                                    "UpdEn,UpdTime,CreateTime,CreateEn,RecordType," & _
                                    "AuthorizeStatus,ACHDesc,MasterId)" & _
                                    " Values ('" & rsSO106("RowId") & "'," & _
                                              "'" & strAry(i, 0) & "'," & _
                                              IIf(strAry(i, 3) = Empty, "Null,", "'" & strAry(i, 3) & "',") & _
                                              IIf(strAry(i, 4) = Empty, "Null,", "'" & strAry(i, 4) & "',") & _
                                              "'CableSoft'," & _
                                              "'" & strUpdTime & "'," & _
                                              "TO_DATE('" & Format(rsSO106("AcceptTime"), "yyyy/mm/dd hh:mm:ss") & "','YYYY/MM/DD HH24:MI:SS')," & _
                                              "'CableSoft'," & _
                                              "1,4,'" & strAry(i, 1) & "'," & rsSO106("MasterId") & ")"
                cn.Execute strSQL
        End Select
    Next i
    
    InsertSO106A = True
    Exit Function
ChkErr:
    objLog.WriteLine ("�Ƚs:" & rsSO106("CustId") & " �s�WSO106A����")
End Function

Private Function ChkSO106A(ByVal intType As Integer) As Boolean
  On Error GoTo ChkErr:
    Dim strQry106A As String
    Dim aryAdd() As String
    Dim strCitemCode As String
    Dim strCitemName As String
    strQry106A = "Select Count(*) From " & strOwnerName & "SO106A"
    Call subChoose(strQry106A, "MasterId=" & rsSO106("MasterId"))
    Select Case intType
        Case 0 '�ݴ��X
            Call subChoose(strQry106A, "AuthorizeStatus=4")
            Call subChoose(strQry106A, "RecordType=0")
            If cn.Execute(strQry106A)(0) = 0 Then
                If Not GetCitemName(rsSO106("CitemStr") & "", strCitemCode, strCitemName) Then Exit Function
                If Not GetCitemCode(rsSO106("ACHTNO") & "", rsSO106("ACHTDESC") & "", strCitemCode, strCitemName, True, aryAdd) Then Exit Function
                If Not InsertSO106A(0, aryAdd) Then Exit Function
            End If
        Case 1 '�w���v���\
            Call subChoose(strQry106A, "AuthorizeStatus=1")
            Call subChoose(strQry106A, "RecordType=0")
            If cn.Execute(strQry106A)(0) = 0 Then
                If Not GetCitemName(rsSO106("CitemStr") & "", strCitemCode, strCitemName) Then Exit Function
                If Not GetCitemCode(rsSO106("ACHTNO") & "", rsSO106("ACHTDESC") & "", strCitemCode, strCitemName, True, aryAdd) Then Exit Function
                If Not InsertSO106A(1, aryAdd) Then Exit Function
            End If
        Case 2 '�������v
            Call subChoose(strQry106A, "AuthorizeStatus=2")
            Call subChoose(strQry106A, "RecordType=1")
            If cn.Execute(strQry106A)(0) = 0 Then
                If Not GetCitemName(rsSO106("CitemStr") & "", strCitemCode, strCitemName) Then Exit Function
                If Not GetCitemCode(rsSO106("ACHTNO") & "", rsSO106("ACHTDESC") & "", strCitemCode, strCitemName, True, aryAdd) Then Exit Function
                If Not InsertSO106A(2, aryAdd) Then Exit Function
            End If
        Case 3 '�w���v�����^
            Call subChoose(strQry106A, "AuthorizeStatus is Null")
            Call subChoose(strQry106A, "RecordType=0")
            If cn.Execute(strQry106A)(0) = 0 Then
                If Not GetCitemName(rsSO106("CitemStr") & "", strCitemCode, strCitemName) Then Exit Function
                If Not GetCitemCode(rsSO106("ACHTNO") & "", rsSO106("ACHTDESC") & "", strCitemCode, strCitemName, True, aryAdd) Then Exit Function
                If Not InsertSO106A(3, aryAdd) Then Exit Function
            End If
        Case 4 '�������v�ݴ��X
            Call subChoose(strQry106A, "AuthorizeStatus=4")
            Call subChoose(strQry106A, "RecordType=1")
            If cn.Execute(strQry106A)(0) = 0 Then
                If Not GetCitemName(rsSO106("CitemStr") & "", strCitemCode, strCitemName) Then Exit Function
                If Not GetCitemCode(rsSO106("ACHTNO") & "", rsSO106("ACHTDESC") & "", strCitemCode, strCitemName, True, aryAdd) Then Exit Function
                If Not InsertSO106A(4, aryAdd) Then Exit Function
            End If
    End Select
    ChkSO106A = True
    Exit Function
ChkErr:
    objLog.WriteLine ("�Ƚs:" & rsSO106("Custid") & " �ˬdSO106A����")
End Function
Private Function GetCitemName(ByVal strSeqNo As String, ByRef strCitemCode As String, ByRef strCitemName As String) As Boolean
  On Error GoTo ChkErr
    Dim strTmp As String
    Dim strSQL As String
    Dim strSQLName As String
    Dim strSQLCode As String
    GetCitemName = False
    strTmp = Empty
    strSQLName = "Select CitemName From " & strOwnerName & "SO003 A," & strOwnerName & "SO004 B " & _
                 " Where A.SeqNo In(" & Replace(strSeqNo, "'", "") & ")" & _
                 " And A.CustId=" & rsSO106("CustId") & " And A.CompCode=" & rsSO106("CompCode") & _
                 " And A.CustId=B.CustId(+)" & _
                 " And A.FaciSeqNo=B.SeqNo(+)" & _
                 " And A.CompCode=B.CompCode(+)"
    strSQLCode = "Select CitemCode From " & strOwnerName & "SO003 A," & strOwnerName & "SO004 B " & _
                 " Where A.SeqNo In(" & Replace(strSeqNo, "'", "") & ")" & _
                 " And A.CustId=" & rsSO106("CustId") & " And A.CompCode=" & rsSO106("CompCode") & _
                 " And A.CustId=B.CustId(+)" & _
                 " And A.FaciSeqNo=B.SeqNo(+)" & _
                 " And A.CompCode=B.CompCode(+)"
    If strSeqNo <> "" Then
'        strTmp = ""
'        strTmp = cn.Execute("Select CitemCode From " & strOwnerName & "SO003 Where SeqNo In(" & Replace(strSeqNo, "'", "") & ")" & _
'                            " And CustId=" & rsSO106("CustId") & " And CompCode=" & rsSO106("CompCode")).GetString(, , , ",")
        strTmp = cn.Execute(strSQLCode).GetString(, , , ",")
        If Right(strTmp, 1) = "," Then strCitemCode = Left(strTmp, Len(strTmp) - 1)
'        strTmp = ""
'        strTmp = cn.Execute("Select CitemName From " & strOwnerName & "SO003 Where SeqNo In(" & Replace(strSeqNo, "'", "") & ")" & _
'                            " And CustId=" & rsSO106("CustId") & " And CompCode=" & rsSO106("CompCode")).GetString(, , , ",")
        strTmp = cn.Execute(strSQLName).GetString(, , , ",")
        If Right(strTmp, 1) = "," Then strCitemName = Left(strTmp, Len(strTmp) - 1)
    End If
    GetCitemName = True
    Exit Function
ChkErr:
    objLog.WriteLine ("�Ƚs:" & rsSO106("CustId") & " ���o���O���ئW�٥���")
End Function
Private Function GetCitemCode(ByVal strACHTNO As String, _
                        ByVal strACHTDESC As String, _
                        ByVal strCitemCode As String, _
                        ByVal strCitemName As String, _
                        ByVal blnAdd As Boolean, _
                        ByRef aryCitem() As String) As Boolean
  
    On Error GoTo ChkErr
    '0:�N��n�s�W�@�����v,1:�N��n�s�W�@���������v 2:�N��nUpdate�쥻�����
    Dim aryACHTNO() As String
    Dim aryACHTDESC() As String
    Dim aryCitemCode() As String
    Dim aryCitemName() As String
    Dim rs As New ADODB.Recordset
    Dim i As Long, j As Long
    Dim lngACHTNO As Long
    Dim lngUbound As Long
    
    '�s�W���
    If blnAdd Then
        If strACHTNO = Empty Then ReDim aryCitem(0, 4): GoTo Fin    '�p�G�S�����ACH,���i����
        aryACHTDESC = Split(Replace(strACHTDESC, "'", ""), ",")
        lngUbound = UBound(aryACHTDESC)
        ReDim aryCitem(lngUbound, 4)
        ReDim aryACHTNO(UBound(aryACHTDESC), 1)
        For i = LBound(aryACHTDESC) To UBound(aryACHTDESC)
            aryACHTNO(i, 0) = Split(Replace(strACHTNO, "'", ""), ",")(i)
            aryACHTNO(i, 1) = aryACHTDESC(i)
            aryCitem(i, 0) = aryACHTNO(i, 0)
            aryCitem(i, 1) = aryACHTNO(i, 1)
            aryCitem(i, 2) = "0"    '0�N��s�W
        Next i
        aryCitemCode = Split(Replace(strCitemCode, "'", ""), ",")
        aryCitemName = Split(Replace(strCitemName, "'", ""), ",")
    '�s����
    Else
        With rsSO106
            Dim aryOldACHTDESC() As String  'SO106.CitemStr
            Dim strQry1 As String
            Dim strQry2 As String
            Dim blnDiff As Boolean
            Dim lngAry As Long
            Dim strOldCitemCode As String
            Dim strOldCitemName As String
            Dim blnUpdCitem As Boolean
            Dim aryUpdCitemCode() As String
            Dim aryUpdCitemName() As String
            Dim blnAddNew As Boolean
            '**************************************************************************
            '�s��Ҧ��p�G��SO106�S��ACH�ӵe���]�S��ACH �h���i����
            '#3728 ���դ�OK,.Fields("CitemStr")�n�令"ACHTNO" By Kin 2008/04/09
            If strACHTNO = Empty And IsNull(.Fields("ACHTNO").OriginalValue) Then
                ReDim aryCitem(0, 4)
                GoTo Fin
            End If
            '*************************************************************************
            aryACHTDESC = Split(Replace(strACHTDESC, "'", ""), ",") '�e���W�ҿ����ACH���v����O
            aryOldACHTDESC = Split(Replace(.Fields("ACHTDESC").OriginalValue & "", "'", ""), ",")
            strQry1 = "Select CitemCode From " & strOwnerName & "SO003 Where SeqNo in(" & _
                        Replace(rsSO106("CitemStr").OriginalValue & "", "'", "") & ")" & _
                        " And Custid=" & rsSO106("CustId")
            strQry2 = "Select CitemName From " & strOwnerName & "SO003 Where SeqNo in(" & _
                        Replace(rsSO106("CitemStr").OriginalValue & "", "'", "") & ")" & _
                        " And Custid=" & rsSO106("CustId")
            'ACH�e���ҿ���PSO106�ۦP��ܥu�����O���ئ�����
            If UBound(aryOldACHTDESC) = UBound(aryACHTDESC) Then
                lngUbound = UBound(aryACHTDESC)
                ReDim aryCitem(lngUbound, 4)
                ReDim aryACHTNO(UBound(aryACHTDESC), 1)
                aryCitemCode = Split(Replace(strCitemCode, "'", ""), ",")
                aryCitemName = Split(Replace(strCitemName, "'", ""), ",")
                For i = LBound(aryACHTDESC) To UBound(aryACHTDESC)
                    aryACHTNO(i, 0) = Split(Replace(strACHTNO, "'", ""), ",")(i)
                    aryACHTNO(i, 1) = aryACHTDESC(i)
                    aryCitem(i, 0) = aryACHTNO(i, 0)
                    aryCitem(i, 1) = aryACHTNO(i, 1)
                    aryCitem(i, 2) = "2"            '2�N��O�s��
                Next i
            Else
                blnUpdCitem = True  '�N��ACH������,���O���إi��]������,�ҥH�]�n�����䥦��ACH���O����
                aryUpdCitemCode = Split(Replace(strCitemCode, "'", ""), ",")
                aryUpdCitemName = Split(Replace(strCitemName, "'", ""), ",")
                If Not IsNull(rsSO106("CitemStr").OriginalValue) Then
                    strOldCitemCode = cn.Execute(strQry1).GetString(, , , ",")
                    strOldCitemName = cn.Execute(strQry2).GetString(, , , ",")
                End If
                '***********************************************************************************
                '���W�[�δ�֪�ACH�H�̦h�����O���ج���,�M��Ӳ��ʹLACH�����O����
                If Right(strOldCitemCode, 1) = "," Then
                    strOldCitemCode = Mid(strOldCitemCode, 1, Len(strOldCitemCode) - 1)
                End If
                If Len(strOldCitemCode) > Len(Replace(strCitemCode, "'", "")) Then
                    aryCitemCode = Split(strOldCitemCode, ",")
                    aryCitemName = Split(strOldCitemName, ",")
                Else
                    aryCitemCode = Split(Replace(strCitemCode, "'", ""), ",")
                    aryCitemName = Split(Replace(strCitemName, "'", ""), ",")
                End If
                '**********************************************************************************
                
                'SO106��ACH���ؤj��e���W�ҬD��ACH����
                If UBound(aryOldACHTDESC) > UBound(aryACHTDESC) Then
                    lngUbound = UBound(aryOldACHTDESC)
                    ReDim aryCitem(lngUbound, 4)
                    ReDim aryACHTNO(lngUbound, 1)
                    For i = LBound(aryOldACHTDESC) To UBound(aryOldACHTDESC)
                        blnDiff = True
                        '��J�S���ʹL��ACH
                        For j = LBound(aryACHTDESC) To UBound(aryACHTDESC)
                            If aryOldACHTDESC(i) = aryACHTDESC(j) Then
                                aryCitem(lngAry, 0) = Split(Replace(strACHTNO, "'", ""), ",")(j)
                                aryCitem(lngAry, 1) = aryACHTDESC(j)
                                aryCitem(lngAry, 2) = "2"
                                lngAry = lngAry + 1
                                blnDiff = False
                                Exit For
                            End If
                        Next j
                        
                        '��J���ʹL��ACH
                        If blnDiff Then
                            aryACHTNO(lngAry, 0) = Split(Replace(.Fields("ACHTNO").OriginalValue, "'", ""), ",")(i)
                            aryACHTNO(lngAry, 1) = aryOldACHTDESC(i)
                            aryCitem(lngAry, 0) = aryACHTNO(lngAry, 0)
                            aryCitem(lngAry, 1) = aryACHTNO(lngAry, 1)
                            aryCitem(lngAry, 2) = "1"   '1�N��s�W�@���������v
                            lngAry = lngAry + 1
                        End If
                    Next i
                Else
                    lngUbound = UBound(aryACHTDESC)
                    ReDim aryCitem(lngUbound, 4)
                    ReDim aryACHTNO(lngUbound, 1)
                    For i = LBound(aryACHTDESC) To UBound(aryACHTDESC)
                        blnDiff = True
                        For j = LBound(aryOldACHTDESC) To UBound(aryOldACHTDESC)
                            If aryACHTDESC(i) = aryOldACHTDESC(j) Then
                                aryCitem(lngAry, 0) = Split(Replace(strACHTNO, "'", ""), ",")(i)
                                aryCitem(lngAry, 1) = aryOldACHTDESC(i)
                                aryCitem(lngAry, 2) = "2"
                                lngAry = lngAry + 1
                                blnDiff = False
                                Exit For
                            End If
                        Next j
                        If blnDiff Then
                            aryACHTNO(lngAry, 0) = Split(Replace(strACHTNO, "'", ""), ",")(i)
                            aryACHTNO(lngAry, 1) = aryACHTDESC(i)
                            aryCitem(lngAry, 0) = aryACHTNO(lngAry, 0)
                            aryCitem(lngAry, 1) = aryACHTNO(lngAry, 1)
                            aryCitem(lngAry, 2) = "0"
                            blnAddNew = True
                            lngAry = lngAry + 1
                        End If
                    Next i
                End If
            End If
        End With
    End If
    '���ʹLACH,���e���W��ACH�P���������O����
    If blnUpdCitem Then
        For i = LBound(aryUpdCitemCode) To UBound(aryUpdCitemCode)
            Set rs = cn.Execute("Select * From " & strOwnerName & "CD068 Where " & _
                                    "InStr(',' || CitemCodeStr || ',','," & aryUpdCitemCode(i) & ",')>0")
            For j = LBound(aryCitem) To UBound(aryCitem)
                If aryCitem(j, 0) = rs("ACHTNO") And aryCitem(j, 1) = rs("ACHTDESC") Then
                    aryCitem(j, 3) = IIf(aryCitem(j, 3) = Empty, aryUpdCitemCode(i), aryCitem(j, 3) & "," & aryUpdCitemCode(i))
                    aryCitem(j, 4) = IIf(aryCitem(j, 4) = Empty, aryUpdCitemName(i), aryCitem(j, 4) & "," & aryUpdCitemName(i))
                    Exit For
                
                End If
            Next j
        Next i
    
    End If
    '��JACH�ҹ��������O����,�p�G�O���ACH,�h�n�M��O"1"������,�Ϥ��n�M�䬰0��
    For i = LBound(aryCitemCode) To UBound(aryCitemCode)
        Set rs = cn.Execute("Select * From " & strOwnerName & "CD068 Where " & _
                                    "InStr(',' || CitemCodeStr || ',','," & aryCitemCode(i) & ",')>0")
        If Not rs.EOF Then
            For j = LBound(aryCitem) To UBound(aryCitem)
                If aryCitem(j, 0) = rs("ACHTNO") And aryCitem(j, 1) = rs("ACHTDESC") Then
                    If blnUpdCitem Then
                        If blnAddNew Then
'                            If aryCitem(j, 2) = "0" Then
'                                aryCitem(j, 3) = IIf(aryCitem(j, 3) = Empty, aryCitemCode(i), aryCitem(j, 3) & "," & aryCitemCode(i))
'                                aryCitem(j, 4) = IIf(aryCitem(j, 4) = Empty, aryCitemName(i), aryCitem(j, 4) & "," & aryCitemName(i))
'                                Exit For
'                            End If
                        Else
                            If aryCitem(j, 2) = "1" Then
                                aryCitem(j, 3) = IIf(aryCitem(j, 3) = Empty, aryCitemCode(i), aryCitem(j, 3) & "," & aryCitemCode(i))
                                aryCitem(j, 4) = IIf(aryCitem(j, 4) = Empty, aryCitemName(i), aryCitem(j, 4) & "," & aryCitemName(i))
                                Exit For
                            End If
                        End If
                    Else
                        aryCitem(j, 3) = IIf(aryCitem(j, 3) = Empty, aryCitemCode(i), aryCitem(j, 3) & "," & aryCitemCode(i))
                        aryCitem(j, 4) = IIf(aryCitem(j, 4) = Empty, aryCitemName(i), aryCitem(j, 4) & "," & aryCitemName(i))
                        Exit For
                    End If
                End If
            Next j

        End If
    Next i
Fin:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Erase aryACHTDESC
    Erase aryACHTNO
    Erase aryCitemCode
    Erase aryCitemName
    Erase aryOldACHTDESC
    Erase aryUpdCitemCode
    Erase aryUpdCitemName
    GetCitemCode = True
  Exit Function
ChkErr:
    objLog.WriteLine ("�Ƚs:" & rsSO106("CustId") & " ���o���O���إN�X����")
End Function
'#3946 ���oSO106.MasterId��Sequence By Kin 2008/05/30
Private Function Get_MasterID_Seq() As String
  On Error GoTo ChkErr
    Get_MasterID_Seq = RPxx(cn.Execute("SELECT " & strOwnerName & "S_SO106_MasterId.NEXTVAL FROM DUAL").GetString & "")
  Exit Function
ChkErr:
    MsgBox "���oSO106A.Seqrenec����!", vbCritical, "�T��"
End Function

Private Function RightNow() As Date
  On Error GoTo ChkErr
    If Val(cn.Execute("SELECT SYSTIME FROM " & strOwnerName & "SO041 WHERE ROWNUM <=1 ORDER BY COMPCODE").GetString(adClipString, , , , 1) & "") = 1 Then
        RightNow = cn.Execute("SELECT TO_CHAR(SYSDATE,'YYYY/MM/DD') FROM DUAL").GetString & ""
        RightNow = RPxx(CStr(RightNow)) & " " & RPxx(cn.Execute("SELECT TO_CHAR(SYSDATE,'HH24:MI:SS') FROM DUAL").GetString & "")
    Else
        RightNow = Now
    End If
  Exit Function
ChkErr:
    RightNow = Now
End Function
Private Function GetDTString(dtm As String) As String
  On Error Resume Next
    GetDTString = Format(dtm, gDTfmt)
  Exit Function
End Function
Private Function RPxx(strValue As String) As String
  On Error Resume Next
    strValue = Trim(Replace(strValue, Chr(13), "", 1, , vbTextCompare))
    strValue = Trim(Replace(strValue, Chr(10), "", 1, , vbTextCompare))
    RPxx = Trim(Replace(strValue, Chr(9), "", 1, , vbTextCompare))
End Function
Private Sub SetgiList(objGilist As Object, strFldName1 As String, strFldName2 As String, _
            StrTableName As String, Optional lngTop As Long, Optional lngLeft As Long, Optional lngWidth1 As Long, _
            Optional lngWidth2 As Long, Optional lngFldLen1 As Long, Optional lngFldLen2 As Long, _
            Optional strWhere As String, Optional blnStopFlag As Boolean = False, Optional gicn As ADODB.Connection)
 '���1,���2,���W��,Top,Left,Width1,Width2,FldLen1,FldLen2
  On Error GoTo ChkErr
    If UCase(StrTableName) <> "CD039" Then objGilist.Clear
    objGilist.SetFldName1 (strFldName1)
    objGilist.SetFldName2 (strFldName2)
    'objGilist.SetTableName (StrTableName)
    objGilist.SetTableName strOwnerName & StrTableName
    objGilist.Filter = strWhere
    If lngTop > 0 Then objGilist.Top = lngTop
    If lngLeft > 0 Then objGilist.Left = lngLeft
    If lngWidth1 > 0 Then objGilist.FldWidth1 = lngWidth1
    If lngWidth2 > 0 Then objGilist.FldWidth2 = lngWidth2
    If lngFldLen1 > 0 Then objGilist.FldLen1 = lngFldLen1
    If lngFldLen2 > 0 Then objGilist.FldLen2 = lngFldLen2
    objGilist.FilterStop = blnStopFlag
    If gicn Is Nothing Then Set gicn = cn
    Call objGilist.SendConn(gicn)
  Exit Sub
ChkErr:
    MsgBox "�]�w���q�O������~!", vbCritical, "ĵ�i"
End Sub

