VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1100U 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�{�d�ɳ�@�~ [SO1100U]"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   Icon            =   "SO1100U.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   11025
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton cmdNoUI 
      Caption         =   "�LUI"
      Height          =   345
      Left            =   3630
      TabIndex        =   17
      Top             =   5490
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Caption         =   "�s�ʱ���"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8970
      TabIndex        =   3
      Top             =   60
      Width           =   2025
      Begin VB.CheckBox chkFaciSNO 
         Caption         =   "�]�ƧǸ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   870
         Width           =   1125
      End
      Begin VB.CheckBox chkOrder 
         Caption         =   "�q��渹"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   630
         Width           =   1125
      End
      Begin VB.CheckBox chkAccountNo 
         Caption         =   "���ڱb��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   390
         Width           =   1125
      End
      Begin VB.CheckBox chkDeclarantName 
         Caption         =   "�ӽФH"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdMedia 
      Caption         =   "F3.��岣�ʹC��渹"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1470
      TabIndex        =   7
      Top             =   5490
      Width           =   1995
   End
   Begin VB.CheckBox chkGenOverdue 
      Caption         =   "&1.���ͤ�O�Ȥ�󥻴������O���"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      TabIndex        =   1
      Top             =   210
      Width           =   3225
   End
   Begin prjNumber.GiNumber ginPeriod 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   270
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      AllowZero       =   0   'False
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "����(&X)"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9840
      TabIndex        =   6
      Top             =   5490
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�T�w"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   300
      TabIndex        =   5
      Top             =   5490
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   4125
      Left            =   90
      TabIndex        =   8
      Top             =   1290
      Width           =   10875
      Begin prjGiGridR.GiGridR ggrData 
         Height          =   3750
         Left            =   180
         TabIndex        =   4
         Top             =   150
         Width           =   10530
         _ExtentX        =   18574
         _ExtentY        =   6615
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseCellForeColor=   -1  'True
      End
   End
   Begin Gi_Date.GiDate gdaClctStopDate 
      Height          =   315
      Left            =   6720
      TabIndex        =   2
      Top             =   270
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblClctStopDate 
      Caption         =   "���O�I���"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5610
      TabIndex        =   16
      Top             =   300
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "�Ȥ���w�O�ΥH����w���ƥX��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   450
      TabIndex        =   11
      Top             =   900
      Width           =   2835
   End
   Begin VB.Label lblMemo 
      Caption         =   "�ťիh�H����Ʋ���"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   450
      TabIndex        =   10
      Top             =   630
      Width           =   1875
   End
   Begin VB.Label lblPeriod 
      Caption         =   "���ʹ���"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Top             =   300
      Width           =   825
   End
End
Attribute VB_Name = "frmSO1100U"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#4030 �W�[���W�� By Kin 2008/08/14
Option Explicit
Private strCustid As String
Private WithEvents rsTmp As ADODB.Recordset
Attribute rsTmp.VB_VarHelpID = -1
Private rsSO0304 As New ADODB.Recordset
Private strRowIds As String
Private strShouldDate1 As String
Private strShouldDate2 As String
Private strServiceTypes As String
Private strCitemCodes As String
Private strAryName As String
Private strAryName2 As String
Private strTmpSO032 As String
Private blnOK As Boolean
Private rsCopyClone As Recordset
Private blnCloseRelate As Boolean
Private blnNoDefine As Boolean
Private blnTrans As Boolean
Private blnSuccess As Boolean
Private strRetBillNo As String
Private strCloseDate As String
Private strFaciSeqNo As String
Private rsRetSO032 As New ADODB.Recordset
'Private obj As csBillMend.clsExeCommand
'Private obj As New csBillMend.clsExeCommand
Private obj As Object
Private strPeriod As String
Private Sub GetShouldDate()
  On Error GoTo ChkErr
    Dim rsDate As New ADODB.Recordset
    Dim strQry As String
    strQry = "Select Min(ClctDate) ShouldDate1,Max(ClctDate) ShouldDate2" & _
            " From " & GetOwner & "SO003" & _
            " Where RowId In (" & strRowIds & ")"
    If Not GetRS(rsDate, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    
    strShouldDate1 = Format(rsDate("ShouldDate1"), "yyyymmdd")
    strShouldDate2 = Format(rsDate("ShouldDate2"), "yyyymmdd")
    
    On Error Resume Next
    Call Close3Recordset(rs)
    Exit Sub
ChkErr:
    ErrSub Me.Name, "GetShouldDate"
End Sub
Private Sub GetTmpFiles(ByVal intTables As Integer)
  On Error GoTo ChkErr
    Dim i As Long
    Dim strTmp As String
    strAryName = Empty
    strTmp = Empty
    strTmpSO032 = Empty
    For i = 0 To intTables - 1
        Select Case i
            Case 5
                strTmp = "TMP017"
            Case 8, 9, 10, 11
                strTmp = "TMP0" & i + 5
            Case Else
                strTmp = GetTmpViewName
        End Select
        
        If i = 5 Then strTmpSO032 = strTmp
        strAryName = IIf(strAryName = Empty, strTmp, strAryName & "," & strTmp)
    Next i
    Exit Sub
ChkErr:
    ErrSub Me.Name, "GetTmpFiles"
End Sub
'#4142 �I�s�L�J��ݭn�h�W3��Table By Kin 2008/12/04 For Lawrence
Private Sub GetTmpFiles2(ByVal intTables As Integer)
  On Error GoTo ChkErr
    Dim i As Integer
    Dim strTmp As String
    strAryName2 = Empty
    strTmp = Empty
    For i = 0 To intTables - 1
        strTmp = GetTmpViewName
        strAryName2 = IIf(strAryName2 = Empty, strTmp, strAryName2 & "," & strTmp)
    Next i
  Exit Sub
ChkErr:
    ErrSub Me.Name, "GetTmpFiles2"
End Sub
Private Sub GetCitemCodes()
  On Error GoTo ChkErr
    Dim strQry As String
    strQry = "Select CitemCode From " & GetOwner & "SO003" & _
            " Where RowId In (" & strRowIds & ")"
    strCitemCodes = gcnGi.Execute(strQry).GetString(, , , ",")
    strCitemCodes = Mid(strCitemCodes, 1, Len(strCitemCodes) - 1)
    Exit Sub
ChkErr:
    ErrSub Me.Name, "GetCitemCode"
End Sub
Private Sub GetUseDate()
  On Error GoTo ChkErr
    Dim rsDate As New ADODB.Recordset
    Dim strSO033SQL As String
    
    strSO033SQL = "Select Max(RealStopDate) MaxDate,Count(*) Cnt From " & GetOwner & "SO033" & _
                " Where CustId=" & gCustId & _
                " And UCcode is Not Null " & _
                " And CitemCode In(" & strCitemCodes & ")"
                
    If Not GetRS(rsDate, strSO033SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If Not rsDate.EOF Then
        If Val(rsDate("Cnt")) > 0 Then
            If CDate(rsDate(0).Value) > strCloseDate Then
                strShouldDate2 = Format(strCloseDate & "", "yyyymmdd")
            Else
                strShouldDate2 = Format(CDate(rsDate(0) + 1) & "", "yyyymmdd")
            End If
        End If
    End If
    On Error Resume Next
    Close3Recordset rsDate
    Exit Sub
ChkErr:
    ErrSub Me.Name, "GetRightDate"
End Sub
Private Sub GetServiceTypes()
   On Error GoTo ChkErr
    Dim strQry As String
    strQry = "Select Min(ServiceType) ServiceType From " & GetOwner & "SO003" & _
            " Where RowId In (" & strRowIds & ")" & _
            " Group By ServiceType"
    strServiceTypes = gcnGi.Execute(strQry).GetString(, , , ",")
    strServiceTypes = Mid(strServiceTypes, 1, Len(strServiceTypes) - 1)
    Exit Sub
ChkErr:
    ErrSub Me.Name, "GetServiceType"
End Sub

Private Sub chkAccountNo_Click()
 On Error Resume Next
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            rsTmp("Choice") = "0"
            rsTmp.MoveNext
        Loop
        rsTmp.MoveFirst
    End If
End Sub

Private Sub chkDeclarantName_Click()
  On Error Resume Next
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            rsTmp("Choice") = "0"
            rsTmp.MoveNext
        Loop
        rsTmp.MoveFirst
    End If
End Sub

Private Sub chkOrder_Click()
 On Error Resume Next
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            rsTmp("Choice") = "0"
            rsTmp.MoveNext
        Loop
        rsTmp.MoveFirst
    End If
End Sub

Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
End Sub
Property Get uSuccess() As Boolean
  On Error Resume Next
  uSuccess = blnSuccess
End Property

Property Let uOK(ByVal vData As Boolean)
  On Error Resume Next
  blnOK = vData
End Property
Private Sub cmdMedia_Click()
  On Error Resume Next
  frmSO1100T.Show 1
End Sub
Private Sub EXEMeth()
    Dim aMsg As String
    Dim Ret As Boolean
    Dim aTmpSO032 As String
    If Not blnTrans Then gcnGi.BeginTrans
    '�n�^�ǵ����M��RecordSet �����@�ӪŪ����,�קK��Noting�����
    If Not GetRS(rsRetSO032, "SELECT * FROM " & GetOwner & "SO032 WHERE 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Ret = obj.Exe_UI_SF_Account(rsTmp, aMsg, aTmpSO032, chkGenOverdue.Value, _
            gdaClctStopDate.GetValue(True), ginPeriod.Text)
    
    If aMsg <> Empty Then
        MsgBox aMsg, vbInformation, "�T��"
    End If
    If Not Ret Then
        
        If Not blnTrans Then gcnGi.RollbackTrans
        '���ѭn��쥻������ Update�^��
        obj.FinallProcess
        blnSuccess = False
    Else
        aMsg = Empty
        If obj.uShowForm Then
            If obj.GetShowSO1100U2(aTmpSO032) Then
                '�^�ǵ��M���GRecordSet
                If Not GetRS(rsRetSO032, "SELECT * FROM " & GetOwner & aTmpSO032, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                '���M�i�Ӫ����ݭnShow From
                If strCloseDate = "" Then
                    With frmSO1100U2
                        .uTmpFile = aTmpSO032
                        .Show vbModal
                    End With
                    If aMsg <> Empty Then MsgBox aMsg, vbInformation, "�T��"
                    aMsg = Empty
                    If blnOK Then
                        If Not obj.Exe_UI_SF_TMPTOC1(aMsg) Then
                            If Not blnTrans Then gcnGi.RollbackTrans
                            If aMsg <> "" Then
                                MsgBox aMsg, vbInformation, "�T��"
                            End If
                            '���ѭn��쥻������UPDATE�^��
                            obj.FinallProcess
                            blnSuccess = False
                            Exit Sub
                        End If
                        blnSuccess = True
                    End If
                Else
                    blnSuccess = True
                End If
            Else
                blnSuccess = True
            End If
        End If
        Set rsRetSO032.ActiveConnection = Nothing
        If aMsg <> Empty Then
            MsgBox aMsg, vbInformation, "�T��"
        End If
        obj.FinallProcess
        If Not blnTrans Then gcnGi.CommitTrans
    End If
    If strCloseDate <> Empty Then blnNoDefine = True
    If blnNoDefine Then
        If InStr(1, aMsg, "�п�ܸ��") > 0 Then
            blnNoDefine = False
            Exit Sub
        End If
        Unload Me
    End If
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ExeMeth")
End Sub

Private Sub cmdNoUI_Click()
'    Dim aMsg As String
'    Dim obj2 As New csBillMend.clsExeCommand
'
'    Dim obj3 As New csBillMend.clsExeCommand
'    obj.uCompCode = gCompCode
'    obj.uCustId = gCustId
'    Set obj.uConn = gcnGi
'    obj.uErrPath = ReadGICMIS1("ErrLogPath")
'    obj.ugaryGi = PutGlobal
'    'obj3.EXE_NOUI_COMMAND(amsg,gCustId,gCompCode,,"���O����",,"���M���",,,"�^�Ǫ�RecordSet")
    Dim aMsg As String
    Dim aRec As New ADODB.Recordset
    Dim obj2 As Object
    Dim rsT As New ADODB.Recordset
    Set obj2 = CreateObject("csBillMend.clsExeCommand")
    obj2.uCompCode = gCompCode
    obj2.uCustId = gCustId
    Set obj2.uConn = gcnGi
    obj2.uErrPath = ReadGICMIS1("ErrLogPath")
    obj2.ugaryGi = PutGlobal
    
    gcnGi.BeginTrans
    'Dim obj3 As New csBillMend.clsExeCommand
    
    Call obj2.EXE_NOUI_COMMAND(aMsg, gCustId, gCompCode, , , , "2010/09/05", "2010/09/05", , , aRec)
    obj2.EXE_DropTmpTable
    If aMsg <> Empty Then
        MsgBox aMsg
    End If
    gcnGi.RollbackTrans
    
End Sub

Private Sub cmdOK_Click()
    EXEMeth
    Exit Sub
  On Error GoTo ChkErr
    
    Dim rsSO003 As New ADODB.Recordset
    Dim rsClone As New ADODB.Recordset
    Dim strQry As String
    Dim lngRetCode As Long
    Dim strRetMsg As String
    Dim strTmpFiles() As String
    Dim i As Long
    Dim j As Long
    
    blnOK = False
    blnSuccess = False
    If rsTmp.RecordCount = 0 Then Exit Sub
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    strRowIds = Empty
    strRetMsg = Empty
    lngRetCode = 0
    strFaciSeqNo = Empty
    Set rsClone = rsTmp.Clone
    rsClone.MoveFirst
    Do While Not rsClone.EOF
        If rsClone("Choice") = "1" Then
            strRowIds = IIf(strRowIds <> "", strRowIds & ",'" & rsClone("RowId") & "'", "'" & rsClone("RowId") & "'")
            '#5569 �n��]�ƬD�X�Ӷǵ���� By Kin 2010/03/03
            If rsClone("FaciSeqNo") & "" <> "" Then
                If strFaciSeqNo = Empty Then
                    strFaciSeqNo = "'" & rsClone("FaciSeqNo") & "'"
                Else
                    If InStr(1, strFaciSeqNo, rsClone("FaciSeqNo") & "") <= 0 Then
                        strFaciSeqNo = strFaciSeqNo & ",'" & rsClone("FaciSeqNo") & "'"
                    End If
                End If
            End If
        End If
    
        rsClone.MoveNext
    Loop
    '#5569 �n��]�ƬD�X�Ӷǵ���� By Kin 2010/03/03
    If strFaciSeqNo <> Empty Then
        If InStr(1, strFaciSeqNo, ",") > 0 Then
            strFaciSeqNo = "In(" & strFaciSeqNo & ")"
        Else
            strFaciSeqNo = "=" & strFaciSeqNo
        End If
    End If
    If strRowIds = "" Then MsgBox "�п�ܸ�ơI", vbInformation, "�T��": GoTo Fin
    strQry = "Select Period,Amount,CustAllot,CLCTSTOPDATE,RowId" & _
            " From " & GetOwner & "SO003" & _
            " Where RowId In (" & strRowIds & ")"
    If Not GetRS(rsSO003, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Call GetShouldDate '�I��������O���ت�ClctDate�̤j�ȻP�̤p��,�a�J��ݪ�ShouldDate1�PShouldDate2
    Call GetServiceTypes '�I��������O���ت��A�ȧO
    Call GetCitemCodes  '�I����������O����
    '****************************************************************
    '#4306 �P�_�I��O�n�ε��M��B�ݦ���ΤU���� By Kin 2009/03/25
    If strCloseDate <> Empty Then
        GetUseDate
    End If
    '****************************************************************
    If Not blnTrans Then gcnGi.BeginTrans
    '�p�G����J����,�n���i��UPDATE,����ݰ�����A�N���UPDATE�^�� By Kin 2008/08/06
    '#5220 �W�[�@�Ӧ��O�I��� By Kin 2009/10/06
    If ginPeriod.Text <> "" Or gdaClctStopDate.GetValue <> "" Then
        rsSO003.MoveFirst
        Do While Not rsSO003.EOF
            If ginPeriod.Text <> "" Then
                If Val(rsSO003("CustAllot")) <> 1 Then
                    If Val(rsSO003("Amount")) <> 0 Then
                        rsSO003("Amount") = (Val(rsSO003("Amount")) / Val(rsSO003("Period"))) * Val(ginPeriod.Text)
                    End If
                    rsSO003("Period") = ginPeriod.Text
                    'rsSO003.Update
                End If
            End If
            If Len(gdaClctStopDate.GetValue) > 0 Then
                rsSO003("CLCTSTOPDATE") = gdaClctStopDate.GetValue(True)
                'rsSO003.Update
            End If
            rsSO003.Update
            rsSO003.MoveNext
        Loop
    End If
    '�I�s�X�b���
    Call GetTmpFiles(20)
    '*****************************************
    '#4142 �L�J�n�h�[�T��Table By Kin 2008/12/04 For Lawrence
    Call GetTmpFiles2(3)
    '*****************************************
    lngRetCode = SF_ACCOUNTING(gcnGi, Format(RightNow, "yyyymm"), Val(strShouldDate1), Val(strShouldDate2), _
                             " In (" & gCompCode & " )", strCitemCodes, "", "", "", "", "", "", 1 _
                             , Val(chkGenOverdue.Value), 0, 0, str(gCustId), garyGi(0), garyGi(1), _
                             strServiceTypes, "", "", 0, strAryName, strFaciSeqNo, strRetMsg)
                         
    
    If lngRetCode < 0 Then
        If Not blnTrans Then gcnGi.RollbackTrans
        MsgBox strRetMsg, vbCritical, giMsgWarning
        GoTo Fin
    Else
        lngRetCode = 0
        If gcnGi.Execute("Select Count(*) From " & strTmpSO032)(0) = 0 Then
            MsgBox strRetMsg
        Else
            Call ReplaceT
            With frmSO1100U2
              .uTmpFile = strTmpSO032
              .Show vbModal
            End With
            '�I�s�L�J���
            If blnOK Then
                lngRetCode = 0
                strRetMsg = Empty
                
                lngRetCode = SF_TMPTOC1(gcnGi, garyGi(1), gCompCode, "", strTmpSO032 & "," & strAryName2, strRetMsg)
                If lngRetCode < 0 Then
                    gcnGi.RollbackTrans
                    MsgBox strRetMsg, vbInformation, "�T���I"
                    GoTo Fin
                Else
                    MsgBox strRetMsg, vbInformation, "�T���I"
                End If
            End If
        End If
        '�N����UPDATE�^��
        '#5220 �W�[�@�Ӧ��O�I��� By Kin 2009/10/06
        If ginPeriod.Text <> "" Or gdaClctStopDate.GetValue <> "" Then
            rsSO003.MoveFirst
            Do While Not rsSO003.EOF
                rsSO0304.Find "ROWID='" & rsSO003("RowId") & "'", , adSearchForward, 1
                If ginPeriod.Text <> "" Then
                    rsSO003("Period") = rsSO0304("Period")
                    rsSO003("Amount") = rsSO0304("Amount")
                End If
                If Len(gdaClctStopDate.GetValue) > 0 Then
                    rsSO003("CLCTSTOPDATE") = rsSO0304("CLCTSTOPDATE")
                End If
                rsSO003.Update
                rsSO003.MoveNext
            Loop
    
        End If
    End If
    
    If Not blnTrans Then gcnGi.CommitTrans
Fin:
    On Error Resume Next
   
    
    
    Close3Recordset rsSO003
    Close3Recordset rsClone
    'strTmpFiles = Split(strAryName, ",")
    '�n�s�L�J���ʺA�ɦW�@�_�R�� By Kin 2008/12/04
    For j = 0 To 1
        Select Case j
            Case 0
                strTmpFiles = Split(strAryName, ",")
            Case 1
                strTmpFiles = Split(strAryName2, ",")
        End Select
        For i = LBound(strTmpFiles) To UBound(strTmpFiles)
          On Error Resume Next
            Select Case UCase(strTmpFiles(i))
                Case "TMP017", "TMP013", "TMP014", "TMP015", "TMP016"
                Case Else
                    gcnGi.Execute "Drop Table " & strTmpFiles(i)
            End Select
'            If UCase(strTmpFiles(i)) <> "TMP017" Then
'                gcnGi.Execute "Drop Table " & strTmpFiles(i)
'            End If
        Next i
    Next j
    blnSuccess = True
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    If strCloseDate <> Empty Then blnNoDefine = True
    If blnNoDefine Then Unload Me
    Exit Sub
ChkErr:
    blnOK = False
    blnSuccess = False
    Me.Enabled = True
    gcnGi.RollbackTrans
    Screen.MousePointer = vbDefault
    ErrSub Me.Name, "cmdOK_Click"
End Sub
Private Sub ReplaceT()
    On Error GoTo ChkErr
    Dim rsBill As New ADODB.Recordset
    Dim rsServiceType As New ADODB.Recordset
    Dim strSeqNo As String
    Dim strBillNo As String
    strRetBillNo = Empty
    '******************************************************************
    '#4158 �P�@�ӪA�ȧO���ͥX�Ӫ�T��n�@�� By Kin 2008/10/21
    If Not GetRS(rsServiceType, "Select ServiceType From " & strTmpSO032 & " Group By ServiceType", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Do While Not rsServiceType.EOF
        strBillNo = Empty
        strBillNo = GetInvoiceNo("T", rsServiceType("ServiceType") & "")
        If Not GetRS(rsBill, "Select * From " & strTmpSO032 & " Where ServiceType='" & rsServiceType("ServiceType") & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        Do While Not rsBill.EOF
            rsBill("BillNo") = strBillNo
            rsBill.Update
            rsBill.MoveNext
        Loop
        strRetBillNo = IIf(strRetBillNo = Empty, "'" & strBillNo & "'", strRetBillNo & ",'" & strBillNo & "'")
        rsServiceType.MoveNext
    Loop
    '******************************************************************
'    If Not GetRS(rsBill, "Select * From " & strTmpSO032, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
'    Do While Not rsBill.EOF
''        strSeqNo = GetInvoiceNo("T", "C")
''        strSeqNo = GetBillSeqNo("C")
'        'rsBill("BillNo") = Left(rsBill("BillNo"), 6) & "T" & rsBill("ServiceType") & strSeqNo
'        rsBill("BillNo") = GetInvoiceNo("T", rsBill("ServiceType") & "")
'        rsBill.Update
'        rsBill.MoveNext
'    Loop
    On Error Resume Next
    CloseRecordset rsBill
    CloseRecordset rsServiceType
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "ReplaceT")
End Sub
Private Function GetBillSeqNo(ByVal strServiceType) As String
  On Error GoTo ChkErr
    Select Case strServiceType
        Case "C"
            GetBillSeqNo = Format(RPxx(gcnGi.Execute("SELECT " & GetOwner & "S_SO033_CATV.NEXTVAL FROM DUAL").GetString & ""), "0000000")
    End Select

  Exit Function
ChkErr:
    ErrSub Me.Name, "GetSeqNo"
End Function
'�I�s�X�b��ݵ{�� By Kin 2008/08/06
'#5569 �W�[�ǤJ�]�ƧǸ� By Kin 2010/03/03
Private Function SF_ACCOUNTING(ByVal cnConn As ADODB.Connection, ByVal p_YM As String, ByVal P_DAY1 As Long, _
ByVal P_DAY2 As Long, ByVal p_CompSQL As String, ByVal p_CitemSQL As String, _
ByVal p_ClassSQL As String, ByVal P_AREASQL As String, ByVal p_ServSQL As String, _
ByVal p_ClctAreaSQL As String, ByVal P_MDUSQL As String, ByVal p_StrtSQL As String, _
ByVal P_ADDRTYPE As Long, ByVal p_GENOVERDUE As Long, ByVal p_GENPRCUST As Long, _
ByVal P_TOENDDATE As Long, ByVal P_CUSTIDLIST As String, ByVal P_UPDEN As String, _
ByVal P_UPDNAME As String, ByVal p_SERVICETYPE As String, ByVal P_BANKCODE As String, _
ByVal p_CMSQL As String, ByVal p_StopPRCust As Long, ByVal p_FileName As String, _
ByVal p_FaciSeqNoSQL, ByRef P_RETMSG As String) As Integer
On Error GoTo ChkErr
        Dim ComSF_ACCOUNTING As New ADODB.Command
        If cnConn Is Nothing Then
                'MsgBox vmsgSendConn, vbCritical, vmsgPrompt
                Exit Function
        End If

        With ComSF_ACCOUNTING
                .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
                .Parameters.Append .CreateParameter("P_YM", adVarChar, adParamInput, 2000, p_YM)
                .Parameters.Append .CreateParameter("P_DAY1", adVarNumeric, adParamInput, , P_DAY1)
                .Parameters.Append .CreateParameter("P_DAY2", adVarNumeric, adParamInput, , P_DAY2)
                .Parameters.Append .CreateParameter("P_COMPSQL", adVarChar, adParamInput, 2000, p_CompSQL)
                .Parameters.Append .CreateParameter("P_CITEMSQL", adVarChar, adParamInput, 2000, p_CitemSQL)
                .Parameters.Append .CreateParameter("P_CLASSSQL", adVarChar, adParamInput, 2000, p_ClassSQL)
                .Parameters.Append .CreateParameter("P_AREASQL", adVarChar, adParamInput, 2000, P_AREASQL)
                .Parameters.Append .CreateParameter("P_SERVSQL", adVarChar, adParamInput, 2000, p_ServSQL)
                .Parameters.Append .CreateParameter("P_CLCTAREASQL", adVarChar, adParamInput, 2000, p_ClctAreaSQL)
                .Parameters.Append .CreateParameter("P_MDUSQL", adVarChar, adParamInput, 2000, P_MDUSQL)
                .Parameters.Append .CreateParameter("P_STRTSQL", adVarChar, adParamInput, 2000, p_StrtSQL)
                .Parameters.Append .CreateParameter("P_ADDRTYPE", adVarNumeric, adParamInput, , P_ADDRTYPE)
                .Parameters.Append .CreateParameter("P_GENOVERDUE", adVarNumeric, adParamInput, , p_GENOVERDUE)
                .Parameters.Append .CreateParameter("P_GENPRCUST", adVarNumeric, adParamInput, , p_GENPRCUST)
                .Parameters.Append .CreateParameter("P_TOENDDATE", adVarNumeric, adParamInput, , P_TOENDDATE)
                .Parameters.Append .CreateParameter("P_CUSTIDLIST", adVarChar, adParamInput, 2000, P_CUSTIDLIST)
                .Parameters.Append .CreateParameter("P_UPDEN", adVarChar, adParamInput, 2000, P_UPDEN)
                .Parameters.Append .CreateParameter("P_UPDNAME", adVarChar, adParamInput, 2000, P_UPDNAME)
                .Parameters.Append .CreateParameter("P_SERVICETYPE", adVarChar, adParamInput, 2000, p_SERVICETYPE)
                .Parameters.Append .CreateParameter("P_BANKCODE", adVarChar, adParamInput, 2000, P_BANKCODE)
                .Parameters.Append .CreateParameter("P_CMSQL", adVarChar, adParamInput, 2000, p_CMSQL)
                .Parameters.Append .CreateParameter("P_STOPPRCUST", adNumeric, adParamInput, , p_StopPRCust)
                .Parameters.Append .CreateParameter("p_FileName", adVarChar, adParamInput, 2000, p_FileName)
                .Parameters.Append .CreateParameter("p_FaciSeqNoSQL", adVarChar, adParamInput, 4000, p_FaciSeqNoSQL)
                .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, P_RETMSG)
                Set .ActiveConnection = cnConn
                .CommandText = "SF_ACCOUNTING"
                .CommandType = adCmdStoredProc
                .Execute
                P_RETMSG = .Parameters("P_RETMSG").Value & ""
                SF_ACCOUNTING = Val(.Parameters("Return_value").Value & "")
        End With
        Exit Function
ChkErr:
    ErrSub Me.Name, "SF_ACCOUNTING"
End Function
Private Function SF_TMPTOC1(ByVal cnConn As ADODB.Connection, ByVal p_UserId As String, _
    ByVal p_CompCode As Long, ByVal p_SERVICETYPE As String, ByVal p_FileName As String, ByRef P_RETMSG As String) As Long
On Error GoTo ChkErr
        Dim ComSF_TMPTOC1 As New ADODB.Command
        If cnConn Is Nothing Then
                'MsgBox vmsgSendConn, vbCritical, vmsgPrompt
                Exit Function
        End If

        With ComSF_TMPTOC1
                .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
                .Parameters.Append .CreateParameter("P_USERID", adVarChar, adParamInput, 2000, p_UserId)
                .Parameters.Append .CreateParameter("P_COMPCODE", adVarNumeric, adParamInput, 2000, p_CompCode)
                .Parameters.Append .CreateParameter("P_SERVICETYPE", adVarChar, adParamInput, 2000, p_SERVICETYPE)
                .Parameters.Append .CreateParameter("p_FileName", adVarChar, adParamInput, 2000, p_FileName)
                .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, P_RETMSG)
                
                Set .ActiveConnection = cnConn
                .CommandText = "SF_TMPTOC1"
                .CommandType = adCmdStoredProc
                .Execute
                P_RETMSG = .Parameters("P_RETMSG").Value & ""
                SF_TMPTOC1 = Val(.Parameters("Return_value").Value & "")
        End With
        Exit Function
ChkErr:
        ErrSub "clsStoreFunction", "SF_TMPTOC1"
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If Not ChkGiList(KeyCode) Then Exit Sub
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
        Case vbKeyF2
            If cmdOK.Enabled Then cmdOK.Value = True
        Case vbKeyF3
            If cmdMedia.Enabled Then cmdMedia.Value = True
    End Select
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    
    If Not Alfa Then
        cmdNoUI.Visible = False
    Else
        cmdNoUI.Visible = True
    End If
    Set obj = CreateObject("csBillMend.clsExeCommand")
    Set rsTmp = New ADODB.Recordset
    obj.uCompCode = gCompCode
    obj.uCustId = gCustId
    Set obj.uConn = gcnGi
    obj.uErrPath = ReadGICMIS1("ErrLogPath")
    obj.ugaryGi = PutGlobal
    If Not rsCopyClone Is Nothing Then
        Set obj.uRsCopy = rsCopyClone
    End If
    Set rsTmp = obj.GetDefineRs
    If strCloseDate <> "" Then
        obj.uCloseDate = strCloseDate
        gdaClctStopDate.SetValue strCloseDate
    End If
    If strPeriod <> "" Then
        ginPeriod.Text = strPeriod
    End If
    If strCloseDate <> "" Then
        lblPeriod.Visible = False
        ginPeriod.Visible = False
        'gdaClctStopDate.Enabled = False
    End If
    'If Not DefineRs Then Unload Me
    Call subGrid
    cmdMedia.Enabled = GetUserPriv("SO1100", "(SO1100T)")
    '#5220 �W�[�@�Ӧ��O�I��� By Kin 2009/10/06
    gdaClctStopDate.Enabled = GetUserPriv("SO1100U", "(SO1100U1)") And strCloseDate = ""
'    If Alfa Then
'        blnCloseRelate = True
'    End If
    If blnCloseRelate Then
        cmdMedia.Enabled = False
        cmdMedia.Visible = False
        Frame2.Enabled = False
        Frame2.Visible = False
        lblClctStopDate.Visible = True
        gdaClctStopDate.Visible = True
    End If
    If Alfa Then
        'strCloseDate = Format("2009/02/28", "yyyy/mm/dd")
    End If
    If strCloseDate <> Empty Then
        chkGenOverdue.Value = 1
    End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub
Property Let uCloseRelate(ByVal vData As Boolean)
  On Error GoTo ChkErr
  blnCloseRelate = vData
  Exit Property
ChkErr:
  ErrSub Me.Name, "uCloseRelate"
End Property
Property Let uRS(ByRef vData As Recordset)
  On Error GoTo ChkErr
    If Not vData Is Nothing Then
        Set rsCopyClone = vData
        blnNoDefine = True
    End If
     
  Exit Property
ChkErr:
  ErrSub Me.Name, "uRs"
End Property
Property Let uTrans(ByVal vData As Boolean)
  On Error GoTo ChkErr
    blnTrans = vData
  Exit Property
ChkErr:
  ErrSub Me.Name, "uTrans"
End Property
Property Get uSO032RecordSet() As ADODB.Recordset
  On Error Resume Next
    Set uSO032RecordSet = rsRetSO032
End Property


Property Let uCustId(ByVal vData As String)
  On Error GoTo ChkErr
    strCustid = vData
    Exit Property
ChkErr:
  ErrSub Me.Name, "uCustId"
End Property
Private Function DefineRs() As Boolean
  On Error GoTo ChkErr
    Dim strQry As String
    Dim i As Long
    Dim strWhere As String
    strWhere = " SO003.CUSTID=" & gCustId & _
             " And SO003.COMPCODE=" & gCompCode & _
             " And SO003.CUSTID=SO004.CUSTID(+) And SO003.COMPCODE=SO004.COMPCODE(+)" & _
             " And SO003.FaciSeqNo=SO004.SeqNo(+)" & _
             " And SO003.StopFlag<>1  And SO003.CitemCode=CD019.CodeNo And CD019.NoShowCitem=0"
             
    strQry = "SELECT SO003.OrderNo,SO004.DeclarantName,SO003.CITEMCODE,SO003.CITEMNAME,SO003.PERIOD,SO003.AMOUNT," & _
             "SO003.ACCOUNTNO,SO003.STARTDATE,SO003.StopDate,SO003.CLCTSTOPDATE," & _
             "SO003.SEQNO,SO003.FaciSNo,SO004.FaciName,FaciSeqNo,SO003.RowId,SO003.CustId FROM " & _
             GetOwner & "SO003," & GetOwner & "SO004, " & GetOwner & "CD019"
    If Not blnNoDefine Then
        If Not GetRS(rsSO0304, strQry & " Where 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    Else
        Set rsSO0304 = rsCopyClone
    End If
    With rsTmp
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Fields.Append ("Choice"), adBSTR, 1, adFldIsNullable
        For i = 0 To rsSO0304.Fields.Count - 1
            .Fields.Append rsSO0304.Fields(i).Name, adBSTR, rsSO0304.Fields(i).DefinedSize, adFldIsNullable
        Next i
    End With
    rsTmp.Open
    If Not blnNoDefine Then
        If Not GetRS(rsSO0304, strQry & " Where " & strWhere, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    End If
    If rsSO0304.RecordCount > 0 Then
        rsSO0304.MoveFirst
    End If
    Do While Not rsSO0304.EOF
        rsTmp.AddNew
        For i = 0 To rsSO0304.Fields.Count - 1
            If Not IsNull(rsSO0304.Fields(i).Value) Then
                rsTmp.Fields(rsSO0304.Fields(i).Name).Value = rsSO0304.Fields(i).Value & ""
            End If
        Next i
        If blnNoDefine Then rsTmp("Choice").Value = "1"
        
        
        rsTmp.Update
        rsSO0304.MoveNext
    Loop
    DefineRs = True
    On Error Resume Next
    
    Exit Function
ChkErr:
  ErrSub Me.Name, "OpenData"
End Function

Private Sub subGrid()
    On Error Resume Next
    Dim mFlds As New GiGridFlds
    With mFlds
        .Add "Choice", , , , , "���"
        .Add "OrderNo", , , , , "�q��渹" & Space(10)
        .Add "AccountNo", , , , , "�b��" & Space(20)
        .Add "DeclarantName", , , , , "�ӽФH�m�W" & Space(5)
        .Add "FaciSNO", , , , , "�]�ƧǸ�" & Space(8)
        .Add "CitemCode", , , , , "���O�N�X  "
        .Add "CitemName", , , , , "���O����             "
        .Add "Period", , , , , "����"
        .Add "AMOUNT", , , , , "�������B" & Space(5)
        .Add "StartDate", , , , , "���O�_��" & Space(5)
        .Add "StopDate", , , , , "���O����" & Space(5)
    End With
    ggrData.AllFields = mFlds
    ggrData.SetHead
    Set ggrData.Recordset = rsTmp
    ggrData.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
    Call Close3Recordset(rsSO0304)
    Call Close3Recordset(rsCopyClone)
    blnNoDefine = False
    blnOK = False
'    If strCloseDate <> Empty Then
'        frmSO1100O.uBillNo = strRetBillNo
'        frmSO1100O.uRowId = strRowIds
'    End If
    strRetBillNo = Empty
    strRowIds = Empty
    strCloseDate = Empty
    strPeriod = Empty
    blnCloseRelate = False
    blnTrans = False
    
'    strCustid = ""
'    strCloseDate = ""
'    strRowIds = ""
'    strShouldDate1 = ""
'    strShouldDate2 = ""
'    strServiceTypes = ""
'    strCitemCodes = ""
'    strAryName2 = ""
'    blnNoDefine = False
'    blnSuccess = False
'    strTmpSO032 = ""
    Set obj = Nothing
    ReleaseCOM Me
End Sub

Private Sub ggrData_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If UCase(giFld.FieldName) = "CHOICE" Then
        If Value = 1 Then Value = vbRed
    End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "ggrData_ColorCellData"
End Sub

Private Sub ggrData_DblClick()
  On Error GoTo ChkErr
    Dim bm As Variant
    Dim strQry As String
    Dim blnCancel As Boolean
    Dim strReturnCode As String
    Dim rsSO003A As New ADODB.Recordset
    Dim rsReturnCode As New ADODB.Recordset
    Dim strDeclarantName As String
    Dim strAccountNo As String
    Dim strOrderNo As String
    Dim strFaciSNo As String
    Dim strFindWhere As String
    Dim arySame() As String
    Dim i As Integer
    Dim blnSame As Boolean
    If rsTmp.RecordCount = 0 Then Exit Sub
    If blnNoDefine Then Exit Sub
    bm = rsTmp.Bookmark
    If rsTmp("Choice").Value = "1" Then
        rsTmp("Choice") = "0"
        blnCancel = True
    Else
        rsTmp("Choice") = "1"
        blnCancel = False
    End If
    
    DoEvents
    strReturnCode = ""
    strQry = "Select LinkKey,StepNo From " & GetOwner & "SO003A" & _
            " Where Custid=" & rsTmp("CustId") & _
            " And CompCode=" & gCompCode & _
            " And CitemCode=" & rsTmp("CitemCode") & _
            " And FaciSeqNo='" & rsTmp("FaciSeqNo") & "'" & _
            " And StopFlag<>1"
    If Not GetRS(rsReturnCode, strQry, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    '���M�䥿�t���O����
    
    Do While Not rsReturnCode.EOF
        rsTmp.Bookmark = bm
        
        If rsReturnCode("LinkKey") & "" <> "" Or rsReturnCode("StepNo") & "" <> "" Then
            If rsReturnCode("LinkKey") & "" = "" Then
                strQry = "Select CitemCode,FaciSeqNo From " & GetOwner & "SO003A" & _
                        " Where Custid=" & gCustId & _
                        " And CompCode=" & gCompCode & _
                        " And FaciSeqNo='" & rsTmp("FaciSeqNo") & "'" & _
                        " And StopFlag<>1" & _
                        " And LinkKey=" & rsReturnCode("StepNo")
            Else
                strQry = "Select CitemCode,FaciSeqNo From " & GetOwner & "SO003A" & _
                        " Where Custid=" & gCustId & _
                        " And CompCode=" & gCompCode & _
                        " And FaciSeqNo='" & rsTmp("FaciSeqNo") & "'" & _
                        " And StopFlag<>1" & _
                        " And StepNo=" & rsReturnCode("LinkKey")
            End If
            Set rsSO003A = gcnGi.Execute(strQry)
            Do While Not rsSO003A.EOF
                rsTmp.Find "CitemCode='" & rsSO003A("CitemCode") & "'", , adSearchForward, 1
                '#5080 �٭n�P�_�]�ƬO�_�@�� For Karen By Kin 2009/05/04
                If (Not rsTmp.EOF) Then
                    If (rsTmp("FaciSeqNo") & "" = rsSO003A("FACISEQNO") & "") Then
                        rsTmp("Choice").Value = blnCancel + 1
                    End If
                End If
                rsSO003A.MoveNext
            Loop
        End If
        rsReturnCode.MoveNext
    Loop
    If rsTmp.RecordCount > 0 Then
        rsTmp.AbsolutePosition = bm
    End If
    '#5533 �W�[�]�ƧǸ��s�ʱ���
    If chkDeclarantName.Value Or chkAccountNo.Value Or chkOrder.Value Or chkFaciSNO.Value Then
        strFindWhere = Empty
        strDeclarantName = rsTmp("DeclarantName") & ""
        strAccountNo = rsTmp("AccountNo") & ""
        strOrderNo = rsTmp("OrderNo") & ""
        strFaciSNo = rsTmp("FaciSNO") & ""
        If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            If chkDeclarantName.Value Then
                strFindWhere = "1"
            End If
            If chkAccountNo.Value Then
                strFindWhere = IIf(strFindWhere = Empty, "2", strFindWhere & "," & "2")
            End If
            If chkOrder.Value Then
                strFindWhere = IIf(strFindWhere = Empty, "3", strFindWhere & "," & "3")
            End If
            If chkFaciSNO.Value Then
                strFindWhere = IIf(strFindWhere = Empty, "4", strFindWhere & "," & "4")
            End If
            blnSame = True
            arySame = Split(strFindWhere, ",")
            For i = 0 To UBound(arySame)
                Select Case arySame(i)
                    Case "1"
                        If rsTmp("DeclarantName") & "" <> strDeclarantName Then
                            blnSame = False
                            Exit For
                        End If
                    Case "2"
                        If rsTmp("AccountNo") & "" <> strAccountNo Then
                            blnSame = False
                            Exit For
                        End If
                    Case "3"
                        If rsTmp("OrderNo") & "" <> strOrderNo Then
                            blnSame = False
                            Exit For
                        End If
                    '#5533 �W�[�]�ƧǸ����s�ʱ��� By Kin 2010/03/08
                    Case "4"
                        If rsTmp("FaciSNO") & "" <> strFaciSNo Then
                            blnSame = False
                            Exit For
                        End If
                End Select
            Next i
            If blnSame Then rsTmp("Choice").Value = blnCancel + 1
            rsTmp.MoveNext
        Loop
        
        
    End If
    rsTmp.AbsolutePosition = bm
    rsTmp.Update
    On Error Resume Next
    Screen.MousePointer = vbDefault
    Close3Recordset rsReturnCode
    Close3Recordset rsSO003A
    Exit Sub
ChkErr:
  ErrSub Me.Name, "ggrData_DblClick"
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If giFld.FieldName = "Choice" Then
        If Value = "1" Then
            Value = "V"
        Else
            Value = ""
        End If
    End If
    Exit Sub
ChkErr:
  ErrSub Me.Name, "ggrData_ShowCellData"
End Sub
Property Let uCloseDate(ByVal vDate As String)
  On Error Resume Next
  strCloseDate = vDate
End Property
Property Let uPeriod(ByVal vData As String)
  On Error Resume Next
    strPeriod = vData
End Property
