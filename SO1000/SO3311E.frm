VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSO3311E 
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "���O��/������ֳt�n�� [SO3311E]"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3311E.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   210
      TabIndex        =   16
      Top             =   1020
      Width           =   5925
      Begin VB.CommandButton cmdOpen 
         Caption         =   "...."
         Height          =   345
         Left            =   5070
         TabIndex        =   18
         Top             =   210
         Width           =   555
      End
      Begin VB.TextBox txtFile 
         Height          =   345
         Left            =   1620
         TabIndex        =   17
         Top             =   210
         Width           =   3405
      End
      Begin VB.Label Label2 
         Caption         =   "�פJEXCEL���"
         Height          =   225
         Left            =   210
         TabIndex        =   19
         Top             =   270
         Width           =   1605
      End
   End
   Begin VB.ComboBox cboBillType 
      Height          =   315
      ItemData        =   "SO3311E.frx":0442
      Left            =   270
      List            =   "SO3311E.frx":044F
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   15
      Top             =   600
      Width           =   1185
   End
   Begin VB.CheckBox chkWarningDup 
      Caption         =   "�w��J��ڬO�_ĵ�i"
      Height          =   195
      Left            =   5520
      TabIndex        =   14
      Top             =   660
      Value           =   1  '�֨�
      Width           =   2085
   End
   Begin VB.CheckBox chkUseOldBillNo 
      Caption         =   "�O�_�ϥ��ª��渹"
      Height          =   195
      Left            =   3450
      TabIndex        =   13
      Top             =   660
      Width           =   1875
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   345
      Left            =   10320
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtBillNo 
      Height          =   315
      Left            =   1455
      MaxLength       =   15
      TabIndex        =   1
      Top             =   600
      Width           =   1665
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "�R���Ȧs���"
      Height          =   345
      Left            =   8640
      TabIndex        =   4
      Top             =   570
      Width           =   1395
   End
   Begin VB.CommandButton cmdchkPrint 
      Caption         =   "�d�ݲ��`���"
      Height          =   345
      Left            =   8640
      TabIndex        =   3
      Top             =   120
      Width           =   1395
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   5955
      Left            =   180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2100
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   10504
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   285
      Left            =   1230
      TabIndex        =   0
      Top             =   120
      Width           =   2430
      _ExtentX        =   4286
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
      Locked          =   -1  'True
      FldWidth1       =   600
      FldWidth2       =   1500
      F5Corresponding =   -1  'True
   End
   Begin MSComDlg.CommonDialog comdPath 
      Left            =   10800
      Top             =   1020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "�n����ڼ�:"
      Height          =   195
      Left            =   3930
      TabIndex        =   12
      Top             =   150
      Width           =   1020
   End
   Begin VB.Label lblBillCnt 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   5190
      TabIndex        =   11
      Top             =   150
      Width           =   90
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      Caption         =   "�n���`���B:"
      Height          =   195
      Left            =   6240
      TabIndex        =   10
      Top             =   150
      Width           =   1020
   End
   Begin VB.Label lblTotalAmt 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   7500
      TabIndex        =   9
      Top             =   150
      Width           =   90
   End
   Begin VB.Label lblBillNo 
      AutoSize        =   -1  'True
      Caption         =   "��ڽs��:"
      Height          =   195
      Left            =   270
      TabIndex        =   8
      Top             =   660
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "(��ڧ�,��ڽs��,�Ȥ�s��,�m�W,�ꦬ���B)"
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   1740
      Visible         =   0   'False
      Width           =   8715
   End
   Begin VB.Label lblComp 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��  �q  �O"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   270
      TabIndex        =   6
      Top             =   180
      Width           =   765
   End
End
Attribute VB_Name = "frmSO3311E"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnMDB As New ADODB.Connection
Private rsSo3311 As New ADODB.Recordset
Private rsSo3311E As New ADODB.Recordset

Private blnCho As Boolean
Private strYM As String '�����W�h�ǨӤ����O�~��
Private strTranDate As String '�����W�h�ǨӤ�TranDate
Private lngPeriod As Long '�����W�h�ǨӤ��w�]����
Private intxPara6 As Integer '�����W�h�ǨӤ�<para6>���Ѽ�
Private strNote As String '�����W�h�ǨӤ��w�]�Ƶ���
Private strClctEn As String '�����W�h�Ǩӹw�]���O�H��
Private strClctName As String '�����W�h�ӹw�]���O�H���m�W
Private strCMCode As String  '�����W�h�Ǩӹw�]���O�覡
Private strCMName As String '�����W�h�Ǩӹw�]���O�覡�W��
Private strPTCode As String '�����W�h�Ǩӹw�]�I�ں���
Private strPTName As String '�����W�h�Ǩӹw�]�I�ں����W��
Private lngRealAmt As String '�����W�h�Ǩӹw�]�ꦬ���B
Private strSTCode As String  '�����W�h�Ǩӹw�]�u����]�N�X
Private strSTName As String '�����W�h�Ǩӹw�]�u����]
Private strRealDate As String '�����W�h�Ǩӹw�]�J�b���
Private blnOption3 As Boolean '�����W�h�ǨӤ��۰ʷs�W�g���ʶ�
Private strManualNo As String '�����W�h����}�渹
Private strCompCode As String
Private intPara23 As Integer
Private rsXls As New ADODB.Recordset
Private intEntryNo As Integer
Private strSQL As String
Private lngRealPeriod As Integer

Private Sub cboBillType_Click()
    On Error Resume Next
        txtBillNo.Text = ""
        Select Case cboBillType.ListIndex
            Case 0
                txtBillNo.MaxLength = 15
            Case 1
                txtBillNo.MaxLength = 12
            Case 2
                txtBillNo.MaxLength = 11
        End Select
        txtBillNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
        Unload Me
End Sub

Private Sub cmdchkPrint_Click()
    On Error GoTo chkErr
    Dim clsAlterCharge As Object
    Dim strSQL As String
    Dim rsSo033 As New ADODB.Recordset
        strSQL = "Select A.*,'" & garyGi(1) & "' UpdEn,'" & GetDTString(RightNow) & "' UpdTime From " & GetOwner & "So077 A," & GetOwner & "So033  B Where A.RcdRowId=B.RowId And B.UCCode Is Null And B.RealDate Is Not Null"
        If Not GetRS(rsSo033, strSQL) Then Exit Sub
        If rsSo033.RecordCount = 0 Then
            MsgNoRcd
        Else
            Set clsAlterCharge = CreateObject("csAlterCharge4.clsAlterCharge")
            With clsAlterCharge
                Set cnn = GetTmpMdbCn
                .uOwnerName = GetOwner
                Set .ucnMDB = cnn
                If Not .InsertChkErrLog(gcnGi, rsSo033) Then Exit Sub
            End With
            Call PrintLogData
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdchkPrint_Click"
End Sub

Private Sub PrintLogData()
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
        If Not GetRS(rsTmp, "Select * From SO07xLog", cnn) Then Exit Sub
        If rsTmp.RecordCount > 0 Then
            With ViewForm
                .uIsSo07x = True
                .uRecordset = rsTmp
                .Show vbModal
            End With
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "PrintLogData"
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo chkErr
    Dim lngRealAmt As Long
        With rsSo3311
            If .BOF Or .EOF Then Exit Sub
            lngRealAmt = .Fields("RealAmt")
            gcnGi.BeginTrans
            gcnGi.Execute "Delete From " & GetOwner & "So077 Where RowId='" & .Fields("RowId") & "'"
            gcnGi.CommitTrans
            Call SetConn
        End With
        lblBillCnt = rsSo3311.RecordCount
        lblTotalAmt = Val(lblTotalAmt) - Val(lngRealAmt)
        Call RefreshGridNo
        MsgBox "�R�����\!!", vbExclamation, gimsgPrompt
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdDelete_Click"
End Sub

Private Sub RefreshGridNo()
    On Error GoTo chkErr
    Dim lngBR As Long
        With rsSo3311
            ggrData.Visible = False
            lngBR = .AbsolutePosition
            
            If Not .EOF Then .MoveFirst
            Do While Not .EOF
                .Fields("EntryNo") = .RecordCount - .AbsolutePosition + 1
                .MoveNext
            Loop
            If lngBR > 0 Then .AbsolutePosition = lngBR
            intEntryNo = .RecordCount
            ggrData.Visible = True
        End With
    Exit Sub
chkErr:
    ggrData.Visible = True
    ErrSub Me.Name, "RefreshGridNo"
End Sub

Private Sub cmdOpen_Click()
On Error GoTo chkErr
    With comdPath
        .DialogTitle = "��ܿ�X���|"
        .Filter = "Microsoft Excel ����ï (*.XLS)|*.XLS"
        .Action = 1
         If .FileName <> "" Then
            txtFile.Text = .FileName
            Call WriteFromExcel
        End If
         
    End With
    Exit Sub
chkErr:
   ErrSub Me.Name, "cmdOpen_Click"
End Sub
'#6449 �W�[Excel�פJ�J�b By Kin 2013/04/30
Private Sub WriteFromExcel()
  On Error GoTo chkErr
    Screen.MousePointer = vbHourglass
    Dim intSucessCount As Integer
    Dim intFailCount As Integer
    Dim intErrCount As Integer
    Dim errMsg As String
    Dim strCustName As String
    Dim strQuerySql As String
    Dim txtErrFile As TextStream
    intSucessCount = 0
    intFailCount = 0
    If Not GetXlsData(txtFile.Text) Then
          MsgBox "�פJ���ѡI", vbInformation, "�T��"
          Screen.MousePointer = vbDefault
    End If
    Do While Not rsXls.EOF
        
        If Not GetRS(rsSo3311E, "Select CustName From " & GetOwner & "So001 Where Custid= " & rsXls("�Ȥ�s��") & _
                " And CompCode = " & strCompCode, gcnGi) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
        If rsSo3311E.EOF Then
            errMsg = IIf(errMsg <> Empty, errMsg & vbCrLf & "(CustId=" & rsXls("�Ȥ�s��") & ") �L����ڤ��Ȥ��ơI", _
                                "(CustId=" & rsXls("�Ȥ�s��") & ") �L����ڤ��Ȥ��ơI")
            intFailCount = intFailCount + 1
            GoTo 66
        Else
            strCustName = rsSo3311E("CustName").Value & ""
            strQuerySql = strSQL & " From " & GetOwner & "So033 Where BillNo='" & rsXls("�u��渹") & _
                                "' And CompCode = " & strCompCode & _
                                " and UCCOde is Not NUll and CancelFlag=0 "
            If Not GetRS(rsSo3311E, strQuerySql, gcnGi) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
        End If
        If rsSo3311E.EOF Then
            If chkWarningDup = 1 Then
                If errMsg <> Empty Then
                    errMsg = errMsg & vbCrLf & "(" & rsXls("�u��渹") & ") �L����ڽs���Φ���ڤw�n���L�A�Юֹ�I"
                Else
                    errMsg = "(" & rsXls("�u��渹") & ") �L����ڽs���Φ���ڤw�n���L�A�Юֹ�I"
                End If
            End If
            intFailCount = intFailCount + 1
            GoTo 66
        End If
         '�ˬd��ڬO�_�w�n���I
        Dim chkSQL As String
        Dim rsTmp As New ADODB.Recordset
        chkSQL = "Select BillNo From " & GetOwner & "So077 Where Billno='" & rsXls("�u��渹") & "'"
            
        If OpenRecordset(rsTmp, chkSQL, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False, False) <> giOK Then
                MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
                Exit Sub
        End If
        
        If Not rsTmp.EOF Then
            If chkWarningDup.Value = 1 Then
                If errMsg = Empty Then
                    errMsg = "(" & rsXls("�u��渹") & " ) ����ڤw�n���L�I�Э��s��J�I"
                Else
                    errMsg = errMsg & vbCrLf & "(" & rsXls("�u��渹") & " ) ����ڤw�n���L�I�Э��s��J�I"
                End If
                intFailCount = intFailCount + 1
                GoTo 66
            End If
        End If
        Dim strExcelDataOK As String
        strExcelDataOK = IsExcelDataOK
        If Len(strExcelDataOK) > 0 Then
            If Len(errMsg) = 0 Then
                errMsg = strExcelDataOK
            Else
                errMsg = errMsg & vbCrLf & strExcelDataOK
            End If
            intFailCount = intFailCount + 1
            GoTo 66
        End If
        '#7314 to check wether the source of stcode is null when the source of default realamt is zero by Kin 2016/10/04
        If (Len(lngRealAmt) > 0) And (lngRealAmt = "0") Then
            If Len(strSTCode) = 0 Then
                errMsg = errMsg & IIf(Len(errMsg) = 0, "", vbCrLf) & "�w�]�ꦬ���B��0�A�����]�w�u����]�I"
                intFailCount = intFailCount + 1
                GoTo 66
            End If
        End If
        'Call GetCustname(rsSo3311E("Custid").Value, strCustName, intCustStatus, strCustStatusname, rsSo3311E("ServiceType"))
    
        'If intCustStatus <> 1 Then If MsgBox("���Ȥ᪬�A�w�D���`������, �O�_�~��?", vbYesNo + vbQuestion, "�T���I") = vbNo Then txtBillNo.Text = "": Exit Sub
        If rsSo3311E.RecordCount > 0 Then rsSo3311E.MoveFirst
        If Not rsSo3311E.EOF Then
            rsSo3311E.MoveFirst
            On Error GoTo lblRollback
            gcnGi.BeginTrans
            intEntryNo = intEntryNo + 1
            Do While Not rsSo3311E.EOF
                With rsSo3311
                    .AddNew
                    .Fields("EntryNO").Value = intEntryNo
                    .Fields("ServiceType").Value = GetFieldValue(rsSo3311E, "ServiceType")
                    .Fields("BillNo").Value = GetFieldValue(rsSo3311E, "BillNO")
                    .Fields("MediaBillNo").Value = GetFieldValue(rsSo3311E, "MediaBillNO")
                    .Fields("Item").Value = GetFieldValue(rsSo3311E, "Item")
                    .Fields("CustName").Value = strCustName
                    .Fields("CustId").Value = GetFieldValue(rsSo3311E, "Custid")
                    .Fields("CItemCode").Value = GetFieldValue(rsSo3311E, "CitemCode")
                    .Fields("CItemName").Value = GetFieldValue(rsSo3311E, "CitemName")
                    .Fields("ShouldAmt").Value = NoZero(rsSo3311E("ShouldAmt").Value, True)
                    .Fields("PrtSNo").Value = NoZero(rsSo3311E("PrtSNo").Value)
                    .Fields("ShouldDate").Value = GetFieldValue(rsSo3311E, "ShouldDate")
                    .Fields("EntryEN").Value = garyGi(0)
                    .Fields("Note").Value = rsSo3311E("Note").Value & strNote
                    .Fields("ServiceType").Value = GetFieldValue(rsSo3311E, "ServiceType")
                    .Fields("CompCode").Value = GetFieldValue(rsSo3311E, "CompCode")
                    If strRealDate = "" Then
                        .Fields("RealDate") = RightDate
                    Else
                        .Fields("RealDate").Value = strRealDate
                    End If
                    .Fields("RealStartDate").Value = GetFieldValue(rsSo3311E, "RealStartDate")
                    .Fields("CancelFlag").Value = NoZero(rsSo3311E("CancelFlag").Value, True)
                    If strClctEn <> "" Then
                        .Fields("ClctEn").Value = strClctEn
                        .Fields("ClctName").Value = strClctName
                    Else
                        If Len(rsSo3311E("ClctEn").Value & "") = 0 Then
                            .Fields("ClctEn").Value = GetFieldValue(rsSo3311E, "OldClctEN")
                            .Fields("ClctName").Value = GetFieldValue(rsSo3311E, "OldClctName")
                        Else
                            .Fields("ClctEn").Value = GetFieldValue(rsSo3311E, "ClctEN")
                            .Fields("ClctName").Value = GetFieldValue(rsSo3311E, "ClctName")
                        End If
                    End If
                    If strCMCode <> "" Then
                        .Fields("CmCode").Value = strCMCode
                        .Fields("CMName").Value = strCMName
                    Else
                        .Fields("CmCode").Value = GetFieldValue(rsSo3311E, "CMCode")
                        .Fields("CMName").Value = GetFieldValue(rsSo3311E, "CMName")
                    End If
                    If strPTCode <> "" Then
                        .Fields("PTCode").Value = strPTCode
                        .Fields("PTName").Value = strPTName
                    Else
                        .Fields("PTCode").Value = GetFieldValue(rsSo3311E, "PTCode")
                        .Fields("PTName").Value = GetFieldValue(rsSo3311E, "PTName")
            
                    End If
                    .Fields("RcdRowID").Value = GetFieldValue(rsSo3311E, "RowID")
                   
                    If (Len(lngRealAmt) > 0 And lngRealAmt = "0") Then
                        .Fields("STCode").Value = strSTCode
                        .Fields("STName").Value = strSTName
                    End If
                    If Val(lngRealAmt) > 0 Then
                        .Fields("ShouldAmt").Value = lngRealAmt
                        .Fields("RealAmt").Value = lngRealAmt
                    Else
                         '#7314 fill out datafield of stcode that if the source of default realamt is zero by Kin 2016/10/04
                         If (Len(lngRealAmt) > 0 And lngRealAmt = "0") Then
                            .Fields("RealAmt").Value = 0
                             .Fields("STCode").Value = strSTCode
                            .Fields("STName").Value = strSTName
                         Else
                            .Fields("RealAmt").Value = NoZero(rsSo3311E("ShouldAmt").Value, True)
                         End If
                    End If
                    lblTotalAmt.Caption = lblTotalAmt.Caption + IIf(IsNull(.Fields("RealAmt").Value), 0, .Fields("RealAmt").Value)
                    If lngPeriod > 0 Then
                        .Fields("RealPeriod").Value = lngPeriod
                        Dim strMsg As String
                        Dim strRealStartDate As String
                        Dim strRealStopDate As String
                        Dim lngRealAmount As Long
                        Dim lngChk As Long
                        strRealStartDate = Format(.Fields("RealStartDate").Value, "yyyymmdd")
                        lngChk = SFGetAmount(False, 2, .Fields("CustId"), .Fields("CitemCode"), lngPeriod, strRealStartDate, .Fields("ServiceType"), .Fields("CompCode"), strRealStopDate, lngRealAmount, strMsg, lngRealPeriod)
                        If lngChk < 0 Then
                            If errMsg = Empty Then
                                errMsg = "(" & rsSo3311E("BillNo") & ") " & strMsg
                            Else
                                errMsg = errMsg & vbCrLf & "(" & rsSo3311E("BillNo") & ") " & strMsg
                            End If
                            rsSo3311.CancelUpdate
                            intFailCount = intFailCount + 1
                            GoTo 66
                        Else
                            If Len(strRealStopDate) = 0 Then lngChk = SFGetAmount(True, 2, .Fields("CustId"), .Fields("CitemCode"), lngPeriod, strRealStartDate, .Fields("ServiceType"), .Fields("CompCode"), strRealStopDate, lngRealAmount, strMsg, lngRealPeriod)
                            If lngChk < 0 Then
                                If errMsg = Empty Then
                                    errMsg = "(" & rsSo3311E("BillNo") & ") " & strMsg
                                 Else
                                    errMsg = errMsg & vbCrLf & "(" & rsSo3311E("BillNo") & ") " & strMsg
                                End If
                                rsSo3311.CancelUpdate
                                intFailCount = intFailCount + 1
                                GoTo 66
                            End If
                        End If
                    Else
                        .Fields("RealPeriod").Value = IIf(Val(rsSo3311E("RealPeriod") & "") > 0, NoZero(rsSo3311E("RealPeriod"), True), 0)
                    End If
                    If Len(strRealStopDate) <> 0 Then
                        .Fields("RealStopDate").Value = Format(strRealStopDate, "####/##/##")
                    Else
                        '#4346 ���n��OldStopDate Upd�^�h By Kin 2009/02/20
                        .Fields("RealStopDate").Value = IIf(Len(rsSo3311E("RealStopDate").Value & "") = 0, Null, NoZero(rsSo3311E("RealStopDate")))
                        '.Fields("RealStopDate").Value = IIf(Len(rsSo3311E("RealStopDate").Value & "") = 0, NoZero(rsSo3311E("OldStopDate")), NoZero(rsSo3311E("RealStopDate")))
                    End If
                    
                    'If strSTCode <> "" Then
                    '    .Fields("STCode").Value = strSTCode
                    '    .Fields("STName").Value = strSTName
                    'Else
                    '    .Fields("STCode").Value = GetFieldValue(rsSo3311E, "STCode")
                    '    .Fields("STName").Value = GetFieldValue(rsSo3311E, "STName")
                    'End If
                    .Fields("EntryEn").Value = garyGi(0)
                    '#6306 NextPeriod�BNextAmt�n�H��ڸ�Ƭ��D By Kin 2012/09/07
                    If Len(rsSo3311E("NextPeriod")) > 0 Then
                        .Fields("NextPeriod") = rsSo3311E("NextPeriod")
                    End If
                    If Len(rsSo3311E("NextAmt")) > 0 Then
                        .Fields("NextAmt") = rsSo3311E("NextAmt")
                    End If
                    .Update
                    intSucessCount = intSucessCount + 1
                End With
                rsSo3311E.MoveNext
            Loop
            gcnGi.CommitTrans
            
            Call SetConn
            lblBillCnt.Caption = rsSo3311.RecordCount
        End If
66:
       rsXls.MoveNext
    Loop
    If intSucessCount = 0 And Len(errMsg) = 0 Then
        MsgBox "�L�����ƶפJ�I", vbInformation, "�T��"
        'Screen.MousePointer = vbDefault
        Exit Sub
    End If
    If intSucessCount > 0 Then
        MsgBox "�פJ���\����:" & intSucessCount & vbCrLf & "�פJ���ѵ���:" & intFailCount
    End If
    If intSucessCount = 0 And Len(errMsg) > 0 Then
        MsgBox "�פJ���\����:" & intSucessCount & vbCrLf & "�פJ���ѵ���:" & intFailCount
    End If
    
    If Len(errMsg) > 0 Then
        Dim strErrPath As String
        strErrPath = ""
        strErrPath = strErrPath & "\" & "SO3311E.log"
        If Not OpenFile(txtErrFile, strErrPath, True) Then Exit Sub
        txtErrFile.WriteLine errMsg
        Shell "notepad " & strErrPath, vbNormalNoFocus
    End If
    
    
    On Error Resume Next
    If Not rsTmp Is Nothing Then
         If rsTmp.State = adStateOpen Then
        rsTmp.Close
        Set rsTmp = Nothing
        End If
    End If
    If Not txtErrFile Is Nothing Then
        txtErrFile.Close
        Set txtErrFile = Nothing
    End If
    Screen.MousePointer = vbDefault
  Exit Sub
lblRollback:
    gcnGi.RollbackTrans
    Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
    Screen.MousePointer = vbDefault
 Call ErrSub(Me.Name, "WriteFromExcel")
End Sub
Private Function GetXlsData(ByVal aFile As String) As Boolean
  On Error GoTo chkErr
    
    Dim strTmp As String
    
    Dim i As Long
    Set rsXls = GetXlsRs(aFile, "Sheet1", " * ")
    If rsXls Is Nothing Then Exit Function
    If rsXls.State = adStateClosed Then Exit Function
    'If rsXls.EOF Then Exit Function
    GetXlsData = True
    Exit Function
chkErr:
  ErrSub Me.Name, "GetXlsData"
End Function
Private Sub Form_Activate()
    txtBillNo.SetFocus
End Sub

Private Sub Form_Load()
    On Error Resume Next
        blnCho = False
        Call subGrd
        Call subGil
        Call SetConn
        Call DefaultValue
End Sub

Private Sub DefaultValue()
    On Error Resume Next
        gilCompCode.SetCodeNo strCompCode
        gilCompCode.Query_Description
        strSQL = "Select RowId,So033.* "
        intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, "", "SO043") & "")
        If intPara23 = 1 Then
            cboBillType.ListIndex = 1
            txtBillNo.MaxLength = 12
        ElseIf intPara23 = 2 Then
            cboBillType.ListIndex = 2
            txtBillNo.MaxLength = 11
        Else
            cboBillType.ListIndex = 0
            txtBillNo.MaxLength = 15
        End If
End Sub

Private Sub subGil()
    On Error Resume Next
        '���q�O
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode)

End Sub

Private Sub SetConn()
    On Error GoTo chkErr
    Dim strcnString As String
    Dim strSQL As String
        ggrData.Visible = False
        strSQL = "Select SO077.RowId,So077.*,So002.CustStatusName from " & GetOwner & "So077 SO077," & GetOwner & "So002 SO002 Where So077.CustId = So002.CustId And So077.ServiceType = So002.ServiceType And SO077.EntryEn='" & garyGi(0) & "' Order by EntryNo Desc"
        If Not GetRS(rsSo3311, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I": Unload Me: Exit Sub
        Set ggrData.Recordset = rsSo3311
        ggrData.Refresh
        ggrData.Visible = True
    Exit Sub
chkErr:
    ggrData.Visible = True
    ErrSub Me.Name, "SetConn"
End Sub

Private Sub subGrd()
'(��ڧ�,��ڽs��,�Ȥ�s��,�m�W,�ꦬ���B)
    On Error Resume Next
    Dim mFlds As New prjGiGridR.GiGridFlds
        With mFlds
            .Add "EntryNo", , , , False, " �Ǹ�", vbRightJustify
            .Add "BillNo", , , , False, "��ڽs��           ", vbLeftJustify
            .Add "CustID", , , , False, "�Ȥ�s��", vbRightJustify
            .Add "CustStatusName", , , , False, "�Ȥ᪬�A", vbLeftJustify
            .Add "CustName", , , , False, "�Ȥ�m�W       ", vbLeftJustify
            .Add "RealPeriod", , , , False, "����", vbRightJustify
            .Add "RealAmt", , , , False, "�ꦬ���B", vbRightJustify
            .Add "RealStartDate", giControlTypeDate, , , False, "   �_�l��  ", vbLeftJustify
            .Add "RealStopDate", giControlTypeDate, , , False, "   �I���  ", vbLeftJustify
            .Add "PrtSNo", , , , False, "�L��Ǹ�     ", vbLeftJustify
            .Add "MediaBillNo", , , , False, "�C��渹   ", vbLeftJustify
            .Add "Note", , , , False, "�Ƶ�              ", vbLeftJustify
        End With
        ggrData.AllFields = mFlds
        ggrData.SetHead
End Sub

Private Sub GetInitData()
On Error GoTo chkErr
Dim rsTmp As New ADODB.Recordset
    rsTmp.MaxRecords = 1
    If OpenRecordset(rsTmp, "Select Max(EntryNo) as BillCount From " & GetOwner & "So077 where EntryEn='" & garyGi(0) & "'", gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False, False) <> giOK Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
    If rsTmp.EOF Then
        intEntryNo = 0
    Else
        If IsNull(rsTmp("BillCount").Value) Then intEntryNo = 0
        intEntryNo = rsTmp("BillCount").Value
    End If
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "GetInitData")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        
        If rsSo3311.State = adStateOpen Then
            If rsSo3311.RecordCount > 0 Then
                frmSO3311G.Show vbModal
                If blnCho = False Then Cancel = True: Exit Sub
            End If
        End If
        intEntryNo = 0
        If rsSo3311.State = adStateOpen Then rsSo3311.Close
        If rsSo3311E.State = adStateOpen Then rsSo3311E.Close
        If rsXls.State = adStateOpen Then rsXls.Close
        Set rsSo3311 = Nothing
        Set rsSo3311E = Nothing
        Set rsXls = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub txtBillNo_Change()
    On Error GoTo chkErr
    Dim lngBillNo As Long
    Dim strQuerySql As String
    Dim intCustStatus As Integer, strCustStatusname As String
    Dim strCustName As String
    Dim lngAmt As Long
        If txtBillNo = "" Then Exit Sub
        If Len(txtBillNo) = 11 And chkUseOldBillNo.Value = 1 Then
            If InStr("BT", Mid(txtBillNo, 5, 1)) > 0 Then
                txtBillNo = GetADString(Left(txtBillNo, 4), False) & Mid(txtBillNo, 5, 1) & "C" & Format(Mid(txtBillNo, 6), "0000000")
            End If
        End If
        If Len(txtBillNo.Text) < 11 Then Exit Sub
        If Len(txtBillNo.Text) <> txtBillNo.MaxLength Then Exit Sub
     
        '���׬�15�X�B��7�X��'B', �h����ڽs��:  (�榡YYYYMMB99999999)
        '�_�h���~�I
        Select Case Len(txtBillNo.Text)
            Case 15
                If InStr(1, txtBillNo.Text, "B", vbTextCompare) = 7 Then
                    '�����s�����k��8�X�ର�Ʀr, ���Y��<�Ȥ�s��>, �ھڦ��Ƚs�ܫȤ�򥻸���ɨ�<�Ȥ�m�W>�P<�Ȥ᪬�A>
                    lngBillNo = Right(txtBillNo.Text, 7)
                    If Not GetRS(rsSo3311E, "Select CustName From " & GetOwner & "So001 Where Custid=" & lngBillNo & " And CompCode = " & strCompCode, gcnGi) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
                    ', �Y�d�L���Ƚs, �h���~: "�L����ڤ��Ȥ���", �õL�H�U�ʧ@. �Y�����Ƚs
                    If rsSo3311E.EOF Then
                        MsgBox "�L����ڤ��Ȥ��ơI", vbExclamation, "�T���I": txtBillNo_GotFocus: Exit Sub
                    Else
                        strCustName = rsSo3311E("CustName").Value & ""
                        'intCustStatus = rsSo3311E("CustStatusCode").Value
                        ', �hSQL="select <�K> from SO033 where BillNo='<��ڽs��>' and RealDate is null and UCCode is not null "
                        'strQuerySql = strSql & " From So033 Where Custid=" & lngBillNo & " And UCCOde is Not NUll"
                        '2001/11/12 Janis �S���n�諸....
                        strQuerySql = strSQL & " From " & GetOwner & "So033 Where BillNo='" & Trim(txtBillNo.Text) & "' And CompCode = " & strCompCode & " and UCCOde is Not NUll and CancelFlag=0"
                        If Not GetRS(rsSo3311E, strQuerySql, gcnGi) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
                    End If
                    If rsSo3311E.EOF Then
                        If cboBillType.ListIndex = 0 Then
                            If chkWarningDup = 1 Then MsgBox "�L����ڽs���Φ���ڤw�n���L�A�Юֹ�I", vbExclamation, "�T���I"
                            Call txtBillNo_GotFocus
                        End If
                        Exit Sub
                    End If
                Else
                    MsgBox "��ڽs�����~�I", vbExclamation, "�T���I":  Exit Sub
                End If
            Case 12
                '�E  �Y���׬�12�X, �h���L��Ǹ�: (�榡YYYYMM999999)
                'SQL = "select <�K> from SO033 where PrtSNo='<��ڽs��>' and RealDate is null and UCCode is not null "
                'lngBillNo = Right(txtBillNo.Text, 8)
                '2001/11/12 Janis ���n�諸
                strQuerySql = strSQL & " From " & GetOwner & "SO033 where PrtSNo='" & UCase(txtBillNo.Text) & "' And CompCode = " & strCompCode & " and UCCode is not null "
                If Not GetRS(rsSo3311E, strQuerySql, gcnGi) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
    '                '�Y��J�ȫDB��榡 (�ӬO�{�ɳ�ΦL��Ǹ��榡), �h�ھڨ��X�����O��ƲĤ@�����Ȥ�s��, �ܫȤ�򥻸���ɨ�
                If rsSo3311E.EOF Then
                    If cboBillType.ListIndex = 1 Then
                        If chkWarningDup = 1 Then MsgBox "�L����ڽs���Φ���ڤw�n���L�A�Юֹ�I", vbExclamation, "�T���I"
                        'Cancel = True
                        'Call ClearRcd
                        Call txtBillNo_GotFocus
                    End If
                    Exit Sub
                End If
            Case 11
                '�E  �Y���׬�11�X, �h���C��渹: (�榡YYMM9999999)
                'SQL = "select <�K> from SO033 where MediaBillNo='<��ڽs��>' and UCCode is not null "
                strQuerySql = strSQL & " From " & GetOwner & "SO033 where MediaBillNo='" & UCase(txtBillNo.Text) & "' And CompCode = " & strCompCode & " and UCCode is not null "
                If Not GetRS(rsSo3311E, strQuerySql, gcnGi) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
    '                '�Y��J�ȫDB��榡 (�ӬO�{�ɳ�ΦL��Ǹ��榡), �h�ھڨ��X�����O��ƲĤ@�����Ȥ�s��, �ܫȤ�򥻸���ɨ�
                If rsSo3311E.EOF Then
                    If cboBillType.ListIndex = 2 Then
                        If chkWarningDup = 1 Then MsgBox "�L����ڽs���Φ���ڤw�n���L�A�Юֹ�I", vbExclamation, "�T���I"
                        Call txtBillNo_GotFocus
                    End If
                    Exit Sub
                End If
            Case Else
                If Len(txtBillNo) = 15 Then
                    MsgBox "��ڽs�����~�I", vbExclamation, "�T���I"
                    txtBillNo.Text = ""
                End If
                Exit Sub
        End Select
        
        '�ˬd��ڬO�_�w�n���I
        If ChkDupData Then
            If chkWarningDup = 1 Then MsgBox "����ڤw�n���L�I�Э��s��J�I", vbExclamation, "�T���I"
            txtBillNo.Text = ""
            Exit Sub
        End If
        
        Call GetCustname(rsSo3311E("Custid").Value, strCustName, intCustStatus, strCustStatusname, rsSo3311E("ServiceType"))
        
        If intCustStatus <> 1 Then If MsgBox("���Ȥ᪬�A�w�D���`������, �O�_�~��?", vbYesNo + vbQuestion, "�T���I") = vbNo Then txtBillNo.Text = "": Exit Sub
        If Not IsDataOk Then objSelectAll txtBillNo: Exit Sub
         '#7314 to check wether the source of stcode is null when the source of default realamt is zero by Kin 2016/10/04
        If (Len(lngRealAmt) > 0) And (lngRealAmt = "0") Then
            If Len(strSTCode) = 0 Then
                MsgBox "�w�]�ꦬ���B��0�A�����]�w�u����]�I", vbExclamation, "�T���I"
                txtBillNo.Text = ""
                Exit Sub
            End If
        End If
        '�E  �Y�����, �h�s�W�ܦ��O��Ƶn���Ȧs��(SO077)
        lngAmt = 0
        
        If rsSo3311E.RecordCount > 0 Then rsSo3311E.MoveFirst
        If Not rsSo3311E.EOF Then
            rsSo3311E.MoveFirst
            On Error GoTo lblRollback
            gcnGi.BeginTrans
            intEntryNo = intEntryNo + 1
            Do While Not rsSo3311E.EOF
                With rsSo3311
                    .AddNew
                    .Fields("EntryNO").Value = intEntryNo
                    .Fields("ServiceType").Value = GetFieldValue(rsSo3311E, "ServiceType")
                    .Fields("BillNo").Value = GetFieldValue(rsSo3311E, "BillNO")
                    .Fields("MediaBillNo").Value = GetFieldValue(rsSo3311E, "MediaBillNO")
                    .Fields("Item").Value = GetFieldValue(rsSo3311E, "Item")
                    .Fields("CustName").Value = strCustName
                    .Fields("CustId").Value = GetFieldValue(rsSo3311E, "Custid")
                    .Fields("CItemCode").Value = GetFieldValue(rsSo3311E, "CitemCode")
                    .Fields("CItemName").Value = GetFieldValue(rsSo3311E, "CitemName")
                    .Fields("ShouldAmt").Value = NoZero(rsSo3311E("ShouldAmt").Value, True)
                    .Fields("PrtSNo").Value = NoZero(rsSo3311E("PrtSNo").Value)
                    .Fields("ShouldDate").Value = GetFieldValue(rsSo3311E, "ShouldDate")
                    .Fields("EntryEN").Value = garyGi(0)
                    .Fields("Note").Value = rsSo3311E("Note").Value & strNote
                    .Fields("ServiceType").Value = GetFieldValue(rsSo3311E, "ServiceType")
                    .Fields("CompCode").Value = GetFieldValue(rsSo3311E, "CompCode")
                    If strRealDate = "" Then
                        .Fields("RealDate") = RightDate
                    Else
                        .Fields("RealDate").Value = strRealDate
                    End If
                    '.Fields("RealDate").Value = rsSo3311E("RealDate").Value
                    .Fields("RealStartDate").Value = GetFieldValue(rsSo3311E, "RealStartDate")
                    .Fields("CancelFlag").Value = NoZero(rsSo3311E("CancelFlag").Value, True)
                    If strClctEn <> "" Then
                        .Fields("ClctEn").Value = strClctEn
                        .Fields("ClctName").Value = strClctName
                    Else
                        If Len(rsSo3311E("ClctEn").Value & "") = 0 Then
                            .Fields("ClctEn").Value = GetFieldValue(rsSo3311E, "OldClctEN")
                            .Fields("ClctName").Value = GetFieldValue(rsSo3311E, "OldClctName")
                        Else
                            .Fields("ClctEn").Value = GetFieldValue(rsSo3311E, "ClctEN")
                            .Fields("ClctName").Value = GetFieldValue(rsSo3311E, "ClctName")
                        End If
                    End If
                    If strCMCode <> "" Then
                        .Fields("CmCode").Value = strCMCode
                        .Fields("CMName").Value = strCMName
                    Else
                        .Fields("CmCode").Value = GetFieldValue(rsSo3311E, "CMCode")
                        .Fields("CMName").Value = GetFieldValue(rsSo3311E, "CMName")
                    End If
                    If strPTCode <> "" Then
                        .Fields("PTCode").Value = strPTCode
                        .Fields("PTName").Value = strPTName
                    Else
                        .Fields("PTCode").Value = GetFieldValue(rsSo3311E, "PTCode")
                        .Fields("PTName").Value = GetFieldValue(rsSo3311E, "PTName")
            
                    End If
                    .Fields("RcdRowID").Value = GetFieldValue(rsSo3311E, "RowID")
                    If Val(lngRealAmt) > 0 Then
                        .Fields("ShouldAmt").Value = lngRealAmt
                        .Fields("RealAmt").Value = lngRealAmt
                    Else
                        '#7314 fill out datafield of stcode that if the source of default realamt is zero by Kin 2016/10/04
                        If (Len(lngRealAmt) > 0 And lngRealAmt = "0") Then
                            .Fields("RealAmt").Value = 0
                             .Fields("STCode").Value = strSTCode
                            .Fields("STName").Value = strSTName
                        Else
                            .Fields("RealAmt").Value = NoZero(rsSo3311E("ShouldAmt").Value, True)
                        End If
                        
                    End If
                    lblTotalAmt.Caption = lblTotalAmt.Caption + IIf(IsNull(.Fields("RealAmt").Value), 0, .Fields("RealAmt").Value)
                    If lngPeriod > 0 Then
                        .Fields("RealPeriod").Value = lngPeriod
                        Dim strMsg As String
                        Dim strRealStartDate As String
                        Dim strRealStopDate As String
                        Dim lngRealAmount As Long
                        Dim lngChk As Long
                        strRealStartDate = Format(.Fields("RealStartDate").Value, "yyyymmdd")
                        lngChk = SFGetAmount(False, 2, .Fields("CustId"), .Fields("CitemCode"), lngPeriod, strRealStartDate, .Fields("ServiceType"), .Fields("CompCode"), strRealStopDate, lngRealAmount, strMsg, lngRealPeriod)
                        If lngChk < 0 Then
                            MsgBox strMsg, vbExclamation, "�T��"
                        Else
                            If Len(strRealStopDate) = 0 Then lngChk = SFGetAmount(True, 2, .Fields("CustId"), .Fields("CitemCode"), lngPeriod, strRealStartDate, .Fields("ServiceType"), .Fields("CompCode"), strRealStopDate, lngRealAmount, strMsg, lngRealPeriod)
                            If lngChk < 0 Then
                                MsgBox strMsg, vbExclamation, "�T��"
                                Exit Sub
                            End If
                        End If
                    Else
                        '#4346 �p�GRealPeriod=0 ���n��OldPeriod Upd�^�h By Kin 2009/02/20
                        .Fields("RealPeriod").Value = IIf(Val(rsSo3311E("RealPeriod") & "") > 0, NoZero(rsSo3311E("RealPeriod"), True), 0)
                        '.Fields("RealPeriod").Value = IIf(Val(rsSo3311E("RealPeriod") & "") > 0, NoZero(rsSo3311E("RealPeriod"), True), NoZero(rsSo3311E("OldPeriod"), True))
                    End If
                    If Len(strRealStopDate) <> 0 Then
                        .Fields("RealStopDate").Value = Format(strRealStopDate, "####/##/##")
                    Else
                        '#4346 ���n��OldStopDate Upd�^�h By Kin 2009/02/20
                        .Fields("RealStopDate").Value = IIf(Len(rsSo3311E("RealStopDate").Value & "") = 0, Null, NoZero(rsSo3311E("RealStopDate")))
                        '.Fields("RealStopDate").Value = IIf(Len(rsSo3311E("RealStopDate").Value & "") = 0, NoZero(rsSo3311E("OldStopDate")), NoZero(rsSo3311E("RealStopDate")))
                    End If
                    
                    'If strSTCode <> "" Then
                    '    .Fields("STCode").Value = strSTCode
                    '    .Fields("STName").Value = strSTName
                    'Else
                    '    .Fields("STCode").Value = GetFieldValue(rsSo3311E, "STCode")
                    '    .Fields("STName").Value = GetFieldValue(rsSo3311E, "STName")
                    'End If
                    .Fields("EntryEn").Value = garyGi(0)
                    '#6306 NextPeriod�BNextAmt�n�H��ڸ�Ƭ��D By Kin 2012/09/07
                    If Len(rsSo3311E("NextPeriod")) > 0 Then
                        .Fields("NextPeriod") = rsSo3311E("NextPeriod")
                    End If
                    If Len(rsSo3311E("NextAmt")) > 0 Then
                        .Fields("NextAmt") = rsSo3311E("NextAmt")
                    End If
                    .Update
                End With
                rsSo3311E.MoveNext
            Loop
            gcnGi.CommitTrans
            txtBillNo.Text = ""
            Call SetConn
            lblBillCnt.Caption = rsSo3311.RecordCount
        End If
    Exit Sub
lblRollback:
       gcnGi.RollbackTrans
chkErr:
    Call ErrSub(Me.Name, "txtBillNo_Change")
End Sub
Private Function IsExcelDataOK() As String
  On Error GoTo chkErr
    Dim rs As New ADODB.Recordset
        If rsSo3311E.RecordCount > 0 Then
            rsSo3311E.MoveFirst
            Do While Not rsSo3311E.EOF
                If Not GetRS(rs, "Select StartDate,StopDate,SeqNo,BudgetPeriod,FaciSeqNo From " & GetOwner & "So003 " & _
                                    " Where CustId = " & rsSo3311E("CustId") & " And CitemCode = " & rsSo3311E("CitemCode")) Then Exit Function
                If rs.RecordCount > 0 Then
                    If rs.RecordCount > 1 Then rs.Filter = "FaciSeqNo = '" & GetRsValue("Select FaciSeqNo From " & GetOwner & "So033 Where RowId = '" & rsSo3311E("RowId") & "'") & "'"
                    If rs.RecordCount > 0 Then
'                        If rsSo3311E.Fields("RealStartDate") < rs("StopDate") And rsSo3311E("RealStopDate") > rs("StartDate") Then
'                            If MsgBox("���O��ƻP�g���ʦ��O�����Ƴ���,�T�w�s��??", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then
'                                Call CloseRecordset(rs)
'                                Exit Function
'                            End If
'                        End If
                        If rsSo3311E("RealPeriod") > rs.Fields("BudgetPeriod") And rs.Fields("BudgetPeriod") > 0 Then
                             IsExcelDataOK = "(" & rsSo3311E("BillNo") & " )�J�b���Ƥj������Ѿl����,�L�k�J�b,�нT�{!"
                            Exit Function
                        End If
                    End If
                End If
                rsSo3311E.MoveNext
            Loop
        End If
        Call CloseRecordset(rs)
        IsExcelDataOK = Empty
  Exit Function
chkErr:
    ErrSub Me.Name, "IsExcelDataOK"
End Function
Private Function IsDataOk() As Boolean
    On Error GoTo chkErr
    Dim rs As New ADODB.Recordset
        If rsSo3311E.RecordCount > 0 Then
            rsSo3311E.MoveFirst
            Do While Not rsSo3311E.EOF
                If Not GetRS(rs, "Select StartDate,StopDate,SeqNo,BudgetPeriod,FaciSeqNo From " & GetOwner & "So003 Where CustId = " & rsSo3311E("CustId") & " And CitemCode = " & rsSo3311E("CitemCode")) Then Exit Function
                If rs.RecordCount > 0 Then
                    If rs.RecordCount > 1 Then rs.Filter = "FaciSeqNo = '" & GetRsValue("Select FaciSeqNo From " & GetOwner & "So033 Where RowId = '" & rsSo3311E("RowId") & "'") & "'"
                    If rs.RecordCount > 0 Then
                        If rsSo3311E.Fields("RealStartDate") < rs("StopDate") And rsSo3311E("RealStopDate") > rs("StartDate") Then
                            If MsgBox("���O��ƻP�g���ʦ��O�����Ƴ���,�T�w�s��??", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then
                                Call CloseRecordset(rs)
                                Exit Function
                            End If
                        End If
                        If rsSo3311E("RealPeriod") > rs.Fields("BudgetPeriod") And rs.Fields("BudgetPeriod") > 0 Then
                            MsgBox "�J�b���Ƥj������Ѿl����,�L�k�J�b,�нT�{!", vbCritical, giMsgWarning
                            Exit Function
                        End If
                    End If
                End If
                rsSo3311E.MoveNext
            Loop
        End If
        Call CloseRecordset(rs)
        IsDataOk = True
     Exit Function
chkErr:
    ErrSub Me.Name, "IsDataOk"
End Function

Private Sub txtBillNo_GotFocus()
    On Error Resume Next
        txtBillNo.SelStart = 0
        txtBillNo.SelLength = Len(txtBillNo.Text)
End Sub

Private Sub txtBillNo_KeyPress(keyAscii As Integer)
    On Error Resume Next
        keyAscii = Asc(UCase(Chr(keyAscii)))
'    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 98 Or _
'                KeyAscii = 66 Or KeyAscii = 105 Or KeyAscii = 73 Or KeyAscii = 116 Or _
'                KeyAscii = 84 Or KeyAscii = Asc("C") Then
'                KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    Else
'                KeyAscii = 9
'    End If
End Sub


'�E  �Y��J�ȫDB��榡(�ӬO�{�ɳ�ΦL��Ǹ��榡), �h�ھڨ��X�����O��ƲĤ@�����Ȥ�s��, �ܫȤ�򥻸���ɨ�<�Ȥ�m�W>�P<�Ȥ᪬�A�N�X/�W��>, �N<�Ȥ�m�W>�P<�Ȥ᪬�A�W��>��ܦb�e���W
Private Function GetCustname(ByVal intCustid As Long, ByRef RetCustName, ByRef RetCustStatusCode, _
        ByRef RetCustStatusName, ByVal strServiceType As String)
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        strSQL = "Select A.CustName,B.CustStatusCode,B.CustStatusName From " & GetOwner & "So001 A," & GetOwner & "So002 B Where A.CustId = B.CustId And A.Custid=" & intCustid & " And B.ServiceType='" & strServiceType & "'"
        If Not GetRS(rsTmp, strSQL, gcnGi) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I": Exit Function
        If Not rsTmp.EOF Then
            RetCustName = rsTmp("CustName").Value & ""
            RetCustStatusCode = rsTmp("CustStatusCode").Value & ""
            RetCustStatusName = rsTmp("CustStatusName").Value & ""
        End If
        rsTmp.Close
        Set rsTmp = Nothing
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "GetCustname")
End Function

Private Function ChkDupMdb() As Boolean
On Error GoTo chkErr
Dim rsTmp As New ADODB.Recordset
Dim strSQL As String
ChkDupMdb = True
    strSQL = "Select BillNo From So3311 Where Billno='" & txtBillNo.Text & "'"
    If OpenRecordset(rsTmp, strSQL, cnMDB, adOpenForwardOnly, adLockReadOnly, adUseClient, False, False) <> giOK Then
            MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
            GoTo CloseRS
    End If
    If rsTmp.EOF Then
        ChkDupMdb = False
    End If
CloseRS:
    rsTmp.Close
    Set rsTmp = Nothing
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "ChkDupData")
End Function

'�P�O����ڬO�_�w�L�b�I
Private Function ChkDupData() As Boolean
On Error GoTo chkErr
Dim rsTmp As New ADODB.Recordset
Dim strSQL As String
ChkDupData = True
    If Len(txtBillNo) = 12 Then
        strSQL = "Select BillNo From " & GetOwner & "So077 Where PrtSNo='" & txtBillNo.Text & "'"
    ElseIf Len(txtBillNo) = 11 Then
        strSQL = "Select BillNo From " & GetOwner & "So077 Where MediaBillno='" & txtBillNo.Text & "'"
    Else
        strSQL = "Select BillNo From " & GetOwner & "So077 Where Billno='" & txtBillNo.Text & "'"
    End If
    If OpenRecordset(rsTmp, strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False, False) <> giOK Then
            MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
            GoTo CloseRS
    End If
    If rsTmp.EOF Then
        ChkDupData = False
    End If
CloseRS:
    rsTmp.Close
    Set rsTmp = Nothing
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "ChkDupData")
End Function


'�w�]����(3311A)
Public Property Let uPeriod(ByVal vData As Long)
    lngPeriod = vData
End Property

'�w�]�Ƶ���(3311A)
Public Property Let uNote(ByVal vData As String)
    strNote = vData
End Property

'�����W�h�����O�H���m�W
Public Property Let uClctName(ByVal vData As String)
    strClctName = vData
End Property

'�����W�h�����O�H��
Public Property Let uClctEn(ByVal vData As String)
    strClctEn = vData
End Property

'�����W�h���w�]�J�b���
Public Property Let uRealDate(ByVal vDate As String)
    strRealDate = vDate
End Property

'�����W�h�ǨӤ�TranDate
Public Property Let uTranDate(ByVal vDate As String)
    strTranDate = vDate
End Property

'�����W�h�ǨӤ��w�]���O�覡
Public Property Let uCMCode(ByVal vData As String)
    strCMCode = vData
End Property

'�����W�ǨӤ��w�]���O�覡�W��
Public Property Let uCMName(ByVal vData As String)
    strCMName = vData
End Property

'�����W�h�ǨӤ��w�]�I�ں���
Public Property Let uPTCode(ByVal vData As String)
    strPTCode = vData
End Property

'�����W�h�ǨӤ��w�]�I�ں����W��
Public Property Let uPTName(ByVal vData As String)
    strPTName = vData
End Property

'�����W�h�ǨӤ��w�]�ꦬ���B
Public Property Let uRealAmt(ByVal vData As String)
    lngRealAmt = vData
End Property

'�����W�h�ǨӤ��w�]�u����]
Public Property Let uStCode(ByVal vData As String)
    strSTCode = vData
End Property

'�����W�h�ǨӤ��w�]�u����]�W��
Public Property Let uSTName(ByVal vData As String)
    strSTName = vData
End Property

'�����W�h�ǨӤ��]�w���O�~��I
Public Property Let uYM(ByVal vData As String)
    strYM = vData
End Property

'�����W�h����}�渹�H�K�Ǧ�So3311D
Public Property Let uManualNo(ByVal vData As String)
    strManualNo = vData
End Property

'�����W�h����}�渹
Public Property Get uManualNo() As String
    uManualNo = strManualNo
End Property

Public Property Let uCho(ByVal vData As Boolean)
    blnCho = vData
End Property

'���q�O
Public Property Let uCompCode(ByVal vData As String)
    strCompCode = vData
End Property

