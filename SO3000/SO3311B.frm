VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO3311B 
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   "���O��/������n�� [SO3311B]"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3311B.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame fraMain 
      Height          =   2025
      Left            =   180
      TabIndex        =   18
      Top             =   510
      Width           =   11715
      Begin VB.ComboBox cboBillType 
         Height          =   315
         ItemData        =   "SO3311B.frx":0442
         Left            =   150
         List            =   "SO3311B.frx":044F
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   38
         Top             =   270
         Width           =   1185
      End
      Begin VB.CheckBox chkUseOldBillNo 
         Caption         =   "�O�_�ϥ��ª��渹"
         Height          =   195
         Left            =   2850
         TabIndex        =   7
         Top             =   1560
         Width           =   1875
      End
      Begin VB.TextBox txtBillNo 
         Height          =   315
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   1
         Top             =   270
         Width           =   1545
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "F2. �s��"
         Height          =   375
         Left            =   10200
         TabIndex        =   8
         Top             =   1440
         Width           =   1185
      End
      Begin VB.TextBox txtManualNo 
         Height          =   315
         Left            =   975
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1470
         Width           =   1305
      End
      Begin prjGiList.GiList gilPTCode 
         Height          =   285
         Left            =   9390
         TabIndex        =   5
         Top             =   990
         Width           =   2175
         _ExtentX        =   3836
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
         FldWidth1       =   550
         FldWidth2       =   1300
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilCMCode 
         Height          =   285
         Left            =   6330
         TabIndex        =   4
         Top             =   990
         Width           =   2175
         _ExtentX        =   3836
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
         FldWidth1       =   550
         FldWidth2       =   1300
         F2Corresponding =   -1  'True
      End
      Begin Gi_Date.GiDate gdaRealDate 
         Height          =   375
         Left            =   4170
         TabIndex        =   3
         Top             =   930
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
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
      Begin prjGiList.GiList gilClctEn 
         Height          =   285
         Left            =   990
         TabIndex        =   2
         Top             =   990
         Width           =   2175
         _ExtentX        =   3836
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
         FldWidth1       =   750
         FldWidth2       =   1100
         F2Corresponding =   -1  'True
      End
      Begin VB.Label lblClassName 
         AutoSize        =   -1  'True
         Caption         =   "XXXXXXXXXXXX"
         ForeColor       =   &H00C000C0&
         Height          =   195
         Left            =   9570
         TabIndex        =   37
         Top             =   720
         Width           =   1620
      End
      Begin VB.Label lbl6 
         AutoSize        =   -1  'True
         Caption         =   "�Ȥ����O:"
         ForeColor       =   &H00C000C0&
         Height          =   195
         Left            =   8610
         TabIndex        =   36
         Top             =   720
         Width           =   825
      End
      Begin VB.Label lblCustId 
         BackColor       =   &H00C0C0C0&
         Caption         =   "99999999"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4170
         TabIndex        =   30
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblCMCode 
         AutoSize        =   -1  'True
         Caption         =   "���O�覡"
         Height          =   195
         Left            =   5490
         TabIndex        =   29
         Top             =   1050
         Width           =   780
      End
      Begin VB.Label lblClctEn 
         AutoSize        =   -1  'True
         Caption         =   "���O�H��"
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Top             =   1020
         Width           =   780
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "�m�W:"
         Height          =   195
         Left            =   5490
         TabIndex        =   27
         Top             =   360
         Width           =   435
      End
      Begin VB.Label lblBillNo 
         AutoSize        =   -1  'True
         Caption         =   "��ڽs��"
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   300
         Width           =   780
      End
      Begin VB.Label lblPTCode 
         AutoSize        =   -1  'True
         Caption         =   "�I�ں���"
         Height          =   195
         Left            =   8610
         TabIndex        =   25
         Top             =   1050
         Width           =   780
      End
      Begin VB.Label lbl4 
         AutoSize        =   -1  'True
         Caption         =   "�Ȥ�s��:"
         Height          =   195
         Left            =   3270
         TabIndex        =   24
         Top             =   360
         Width           =   825
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "XXXXXXXXXXXXXXXXXX"
         Height          =   195
         Left            =   6030
         TabIndex        =   23
         Top             =   360
         Width           =   2430
      End
      Begin VB.Label lblRealDate 
         AutoSize        =   -1  'True
         Caption         =   "�J�b���"
         Height          =   195
         Left            =   3270
         TabIndex        =   22
         Top             =   1050
         Width           =   780
      End
      Begin VB.Label lbl5 
         AutoSize        =   -1  'True
         Caption         =   "�Ȥ᪬�A:"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   8610
         TabIndex        =   21
         Top             =   360
         Width           =   825
      End
      Begin VB.Label lblCustStatus 
         AutoSize        =   -1  'True
         Caption         =   "XXXXXXXXXXXX"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   9570
         TabIndex        =   20
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label lblManualNo 
         AutoSize        =   -1  'True
         Caption         =   "��}�渹"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   1500
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   10590
      TabIndex        =   14
      Top             =   90
      Width           =   1215
   End
   Begin VB.Frame fraData 
      Height          =   4965
      Left            =   180
      TabIndex        =   15
      Top             =   2610
      Width           =   11715
      Begin VB.CommandButton cmdDetDelete 
         Caption         =   "�@�o(&3)"
         Height          =   375
         Left            =   3570
         TabIndex        =   11
         Top             =   180
         Width           =   930
      End
      Begin VB.CommandButton cmdDetEdit 
         Caption         =   "�ק�(&2)"
         Height          =   375
         Left            =   2430
         TabIndex        =   10
         Top             =   180
         Width           =   930
      End
      Begin VB.CommandButton cmdDetAdd 
         Caption         =   "�s�W(&1)"
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   180
         Width           =   930
      End
      Begin prjGiGridR.GiGridR ggrData 
         Height          =   4185
         Left            =   150
         TabIndex        =   12
         Top             =   630
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7382
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
      Begin VB.Label Label1 
         Caption         =   "(���O���ءB�X����B�B���ơB�ꦬ���B�B�����_�l�B�����I��B�@�o�B�Ƶ�)"
         Height          =   315
         Left            =   180
         TabIndex        =   17
         Top             =   630
         Visible         =   0   'False
         Width           =   8715
      End
      Begin VB.Label lblDetail 
         AutoSize        =   -1  'True
         Caption         =   "��ڤ��e"
         Height          =   195
         Left            =   300
         TabIndex        =   16
         Top             =   300
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "F3. �w�n�d��"
      Height          =   375
      Left            =   8670
      TabIndex        =   13
      Top             =   90
      Width           =   1575
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   285
      Left            =   1170
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
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "�n����ڼ�:"
      Height          =   195
      Left            =   3990
      TabIndex        =   35
      Top             =   180
      Width           =   1020
   End
   Begin VB.Label lblBillCnt 
      AutoSize        =   -1  'True
      Caption         =   "999999"
      Height          =   195
      Left            =   5160
      TabIndex        =   34
      Top             =   180
      Width           =   540
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      Caption         =   "�n���`���B:"
      Height          =   195
      Left            =   6060
      TabIndex        =   33
      Top             =   180
      Width           =   1020
   End
   Begin VB.Label lblTotalAmt 
      AutoSize        =   -1  'True
      Caption         =   "999,999,999"
      Height          =   195
      Left            =   7200
      TabIndex        =   32
      Top             =   180
      Width           =   900
   End
   Begin VB.Label lblComp 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��  �q  �O"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   300
      TabIndex        =   31
      Top             =   180
      Width           =   765
   End
End
Attribute VB_Name = "frmSO3311B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsSo3311B As New ADODB.Recordset
Private rsArray As New ADODB.Recordset
Private rsArrayNo As New ADODB.Recordset

Private intPara6 As Integer
Private intPara23 As Integer
Private intDayCut As Integer
Private lngSelMod As Long '�O����J��ڬO��ڽs���ΦL��Ǹ�
Private blnOption1 As Boolean '�����W�h���۰ʧ�s�g�����ت����B
Private blnOption2 As Boolean '�����W�h���۰ʧ�s�g�����ؤ�������
Private blnOption3 As Boolean '�����W�h���۰ʷs�W�g���ʸ�Ƹ��
Private strManualNo As String '�����W�h����}�渹
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
Private lngRealAmt As String   '�����W�h�Ǩӹw�]�ꦬ���B
Private strSTCode As String  '�����W�h�Ǩӹw�]�u����]�N�X
Private strSTName As String '�����W�h�Ǩӹw�]�u����]
Private strRealDate As String '�����W�h�Ǩӹw�]�J�b���
Private strCompCode As String
Private strSQL As String

Private intCustid As Long '�O���Ȥ�s��(��So3311F�ϥΡ^
Private strCustName As String '�O���Ȥ�m�W(��So3311F�ϥΡ^
Private blnBillValidate As Boolean

Private strConnString As String
Private strRowId As String '�ȦsRowId�H�Ǧ�So3311F �ϥ�
Private cnMDB As New ADODB.Connection
Private blnCanClose As Boolean '�O�_�i��������� True=Yes ,False =No
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
    On Error GoTo chkErr
        Unload Me
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdCancel_Click"
End Sub

Private Sub cmdDetAdd_Click()
    On Error GoTo chkErr
        With frmSO331AF
            Set .uParentForm = Me
            Set .Conn = cnMDB
            .uDBType = giAccessDb
            .EditMode = giEditModeInsert
            .uBillNo = rsArray("BillNO").Value & ""
            .uPrtSNo = rsArray("PrtSno").Value & ""
            .uClctEn = rsArray("ClctEn").Value & ""
            .uClctName = rsArray("ClctName").Value & ""
            .uPTCode = IIf(IsNull(rsArray("PTCode").Value), 0, rsArray("PTCode").Value)
            .uPTName = rsArray("PTName").Value & ""
            .uRealDate = rsArray("RealDate").Value & ""
            .uCustId = rsArray("CustID").Value
            .uCustName = rsArray("CustName").Value
            gServiceType = rsArray("ServiceType").Value
            .Show vbModal
        End With
'        With frmSO331AF
'            Set .Conn = cnMDB
'            .EditMode = giEditModeEdit
'            .uBillNo = rsArray("BillNO").Value
'            .uPrtSNo = rsArray("PrtSno").Value & ""
'            .uItem = rsArray("Item").Value
'            .uRowId = rsArray("RowId").Value
'            .uCompCode = rsArray("CompCode").Value & ""
'            gServiceType = rsArray("ServiceType").Value & ""
'            .Show vbModal
'        End With
        Call RefreshData
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdDetAdd_Click"
End Sub

Private Sub cmdDetDelete_Click()
        '�ˬd�O�_�i�H�@�o�A�Y�w�@�o�A�h���~�G���"�w�@�o��Ƥ��i�A�@�o "�A�L�H�U�ʧ@�I
    On Error Resume Next
        
        If rsArray("CancelFlag").Value = 1 Then MsgBox "�w�@�o��Ƥ��i�A�@�o", vbExclamation, "�T���I": Exit Sub
        If ChkHave2Discount(rsArray("CitemCode")) Then MsgBox "����2�h�u�f����@�o!!", vbCritical, gimsgPrompt:  Exit Sub
        If Chk2Discount(rsArray("CitemCode")) Then MsgBox "��2�h�u�f����@�o!!", vbCritical, gimsgPrompt:  Exit Sub

    On Error GoTo chkErr
        cnMDB.BeginTrans
        With frmSO1110E
            If Len(rsArray("RowId") & "") < 10 Then
                cnMDB.Execute "Delete From So3311 Where RowId='" & rsArray("RowId") & "'"
                MsgBox "�R�����\�I", vbInformation, "�T���I"
            Else
                If Not rsArray.EOF Then
                    With frmSO1110E
                        Set .uParentForm = Me
                        .uType = 1  '���O���
                        Set .uRS = rsArray
                        .uBillNo = rsArray("BillNO")
                        .uItem = rsArray("Item")
                        .Show vbModal
                    End With
                End If
            End If
        End With
        cnMDB.CommitTrans
    On Error Resume Next
        Call RefreshData
    Exit Sub
chkErr:
    cnMDB.RollbackTrans
    Call ErrSub(Me.Name, "cmdDetDelete")
End Sub

Private Sub cmdDetEdit_Click()
    On Error GoTo chkErr
        If rsArray("CancelFlag").Value = 1 Then MsgBox "�w�@�o��Ƥ���ק�I", vbExclamation, "�T���I": Exit Sub
        With frmSO331AF
            Set .uParentForm = Me
            Set .Conn = cnMDB
            .uDBType = giAccessDb
            .EditMode = giEditModeEdit
            .uBillNo = rsArray("BillNO").Value
            .uPrtSNo = rsArray("PrtSno").Value & ""
            .uItem = rsArray("Item").Value
            .uRowId = rsArray("RowId").Value
            .uClctEn = rsArray("ClctEn").Value & ""
            .uClctName = rsArray("ClctName").Value & ""
            .uPTCode = IIf(IsNull(rsArray("PTCode").Value), 0, rsArray("PTCode").Value)
            .uPTName = rsArray("PTName").Value & ""
            .uRealDate = rsArray("RealDate").Value & ""
            .uCustId = rsArray("CustID").Value
            .uCustName = rsArray("CustName").Value
            gServiceType = rsArray("ServiceType").Value
            .Show vbModal
        End With
'        With frmSO3311F
'            Set .Conn = cnMDB
'            .EditMode = giEditModeEdit
'            .uBillNo = rsArray("BillNO").Value
'            .uPrtSNo = rsArray("PrtSno").Value & ""
'            .uItem = rsArray("Item").Value
'            .uRowId = rsArray("RowId").Value
'            .uCompCode = rsArray("CompCode").Value & ""
'            gServiceType = rsArray("ServiceType").Value & ""
'            .Show vbModal
'        End With
        Call RefreshData
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdDetEdit_Click"
End Sub

Private Sub cmdQuery_Click()
    On Error GoTo chkErr
        With frmSO3311C
            Set .uParentForm = Me
            .Show vbModal
        End With
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdQuery_Click"
End Sub

Private Function IsDataOk() As Boolean
    On Error GoTo chkErr
    Dim rs As New ADODB.Recordset
        If Not ChkDTok Then Exit Function
        If txtBillNo.Text = "" Or rsArray.State = adStateClosed Then txtBillNo.SetFocus: GoTo Warning
        If gdaRealDate.GetValue = "" Then gdaRealDate.SetFocus: GoTo Warning
        If rsArray.RecordCount > 0 Then
            rsArray.MoveFirst
            Do While Not rsArray.EOF
                If rsArray("RealAmt") & "" <> rsArray("ShouldAmt") & "" And rsArray("CancelFlag") = 0 And Len(rsArray("STCode") & "") = 0 Then
                    MsgBox "�ꦬ���B�P�X����B���۵�,�����u����]", vbCritical, giMsgWarning
                    If cmdDetEdit.Enabled Then cmdDetEdit.Value = True
                    Exit Function
                End If
                If Not GetRS(rs, "Select StartDate,StopDate,SeqNo,BudgetPeriod,FaciSeqNo From " & GetOwner & _
                    "So003 Where CustId = " & rsArray("CustId") & " And CitemCode = " & rsArray("CitemCode")) Then Exit Function
                If rs.RecordCount > 0 Then
                    If rs.RecordCount > 1 Then rs.Filter = "FaciSeqNo = '" & GetRsValue("Select FaciSeqNo From " & GetOwner & "So033 Where RowId = '" & rsArray("RowId") & "'") & "'"
                    If rs.RecordCount > 0 Then
                        If rsArray.Fields("RealStartDate") < rs("StopDate") And rsArray("RealStopDate") > rs("StartDate") Then
                            If MsgBox("���O��ƻP�g���ʦ��O�����Ƴ���,�T�w�s��??", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then
                                Call CloseRecordset(rs)
                                Exit Function
                            End If
                        End If
                        If rsArray("RealPeriod") > rs.Fields("BudgetPeriod") And rs.Fields("BudgetPeriod") > 0 Then
                            MsgBox "�J�b���Ƥj������Ѿl����,�L�k�J�b,�нT�{!", vbCritical, giMsgWarning
                            Exit Function
                        End If
                    End If
                End If
                rsArray.MoveNext
            Loop
        End If
        Call CloseRecordset(rs)
        IsDataOk = True
     Exit Function
Warning:
     MsgBox "�Х���J��ڽs���d�ߡI", vbExclamation, "�T���I"
     Exit Function
chkErr:
    ErrSub Me.Name, "IsDataOk"
End Function

Private Sub ScrToRcd()
    On Error GoTo chkErr
    Dim lngBillCnt As Long
    Dim strSQL As String
    Dim rsSaveSo074 As New ADODB.Recordset
    
        strSQL = "Select * From " & GetOwner & "So074 Where 0=1"
        If Not GetRS(rsSaveSo074, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then MsgBox "�s�����ѡI�Э��s�ާ@�I", vbExclamation, "�T���I": Exit Sub
        lblBillCnt.Caption = lblBillCnt.Caption + 1
        lngBillCnt = lblBillCnt.Caption
        If rsArray.RecordCount = 0 Then Exit Sub
        rsArray.MoveFirst
        With rsSaveSo074
            While Not rsArray.EOF
                gcnGi.Execute "Delete  From " & GetOwner & "SO074 Where BillNo='" & rsArray("BillNo") & "' And Item=" & rsArray("Item")
                .AddNew
                .Fields("BillNo").Value = GetFieldValue(rsArray, "Billno")
                .Fields("Item").Value = GetFieldValue(rsArray, "Item")
                .Fields("Custid").Value = GetFieldValue(rsArray, "Custid")
                .Fields("CustName").Value = lblName.Caption
                .Fields("CitemCode").Value = GetFieldValue(rsArray, "CitemCode")
                .Fields("CitemName").Value = GetFieldValue(rsArray, "CitemName")
                .Fields("MediaBillNo").Value = GetFieldValue(rsArray, "MediaBillNo")
                .Fields("PrtSNo").Value = NoZero(rsArray("PrtSNo").Value)
                .Fields("ShouldAmt").Value = NoZero(rsArray("ShouldAmt").Value, True)
                .Fields("ShouldDate").Value = NoZero(rsArray("ShouldDate").Value, True)
                .Fields("RealAmt").Value = NoZero(rsArray("RealAmt").Value, True)
                .Fields("ManualNo").Value = NoZero(rsArray("ManualNo").Value)
                .Fields("RealDate").Value = gdaRealDate.GetValue(True)
                .Fields("RealPeriod").Value = NoZero(rsArray("RealPeriod").Value, True)
                .Fields("RealStartDate").Value = GetFieldValue(rsArray, "RealStartDate")
                .Fields("RealStopDate").Value = GetFieldValue(rsArray, "RealStopDate")
                .Fields("EntryEn").Value = GetFieldValue(rsArray, "EntryEn")
                .Fields("Note").Value = GetFieldValue(rsArray, "Note")
                .Fields("CMCode").Value = NoZero(rsArray("CMCode").Value)
                .Fields("CMName").Value = NoZero(rsArray("CMName").Value)
                .Fields("ClctEn").Value = NoZero(rsArray("ClctEn").Value)
                .Fields("ClctName").Value = NoZero(rsArray("ClctName").Value)
                .Fields("PTCode").Value = NoZero(rsArray("PTCode").Value)
                .Fields("PTName").Value = NoZero(rsArray("PTName").Value)
                .Fields("RcdRowId").Value = GetFieldValue(rsArray, "RowId")
                .Fields("EntryNO").Value = lngBillCnt
                .Fields("StCode").Value = GetFieldValue(rsArray, "STCode")
                .Fields("STName").Value = GetFieldValue(rsArray, "STName")
                .Fields("ServiceType").Value = GetFieldValue(rsArray, "ServiceType")
                .Fields("CompCode").Value = NoZero(rsArray("Compcode"))
                .Fields("ManualNo").Value = NoZero(rsArray("ManualNo").Value)
                .Fields("CancelFlag").Value = NoZero(rsArray("CancelFlag").Value, True)
                .Fields("CancelCode") = NoZero(rsArray("CancelCode"))
                .Fields("CancelName") = NoZero(rsArray("CancelName"))
                .Fields("BankCode") = NoZero(rsArray("BankCode"))
                .Fields("BankName") = NoZero(rsArray("BankName"))
                .Fields("AccountNo") = NoZero(rsArray("AccountNo"))
                .Fields("AuthorizeNo") = NoZero(rsArray("AuthorizeNo"))
                .Fields("AdjustFlag") = Val(rsArray("AdjustFlag") & "")
                .Fields("NextPeriod") = Val(rsArray("NextPeriod") & "")
                .Fields("NextAmt") = Val(rsArray("NextAmt") & "")
                '******************************************************************
                '#3465 �]�Ƭy�����g�J��MDB By Kin 2007/08/28
                .Fields("FaciSeqNo") = rsArray("FaciSeqNo") & ""
                '******************************************************************
                
                '******************************************************************
                '#3436 �x�s�o����Ƭ�����T By Kin 2007/12/17
                If Not IsNull(rsArray("InvSeqNo")) Then
                    .Fields("InvSeqNo") = rsArray("InvSeqNo") & ""
                End If
                '******************************************************************
                
                .Update
                rsArray.MoveNext
            Wend
        End With
    Exit Sub
chkErr:
    ErrSub Me.Name, "ScrToRcd"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo chkErr
        '���n����ˬd�A�Y���L���i�~��I
        If Not IsDataOk Then Exit Sub
        Call ScrToRcd
        Call InitData
        Call ClearRcd
        cmdSave.Tag = "Save"
        txtBillNo.SetFocus
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdSave")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo chkErr
        If Not ChkGiList(KeyCode) Then Exit Sub
        If cmdSave.Enabled And KeyCode = vbKeyF2 Then Call cmdSave_Click: Exit Sub
        If KeyCode = vbKeyF3 Then Call cmdQuery_Click
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
    On Error GoTo chkErr
        '��l�]�w
        Dim strPath As String
        Me.Enabled = False
        strPath = ReadGICMIS1("TmpMDBPath")
            
        'cnMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & IIf(Right(strPath, 1) = "\", strPath, strPath & "\") & "Tmp1111.MDB" & ";Persist Security Info=False"
        cnMDB.Open "Provider=" & GetOleDbProvider & ";Data Source=" & IIf(Right(strPath, 1) = "\", strPath, strPath & "\") & "Tmp1111.MDB" & ";Persist Security Info=False"
        If Not GetSo3311XTmpMDB(cnMDB) Then Exit Sub
        
        rsArray.CursorLocation = adUseClient
        rsArray.Open "Select * from So3311", cnMDB, adOpenKeyset, adLockOptimistic
        
        Call subGrd
        Call subGil
        Call InitData
        Call ClearRcd
        lngSelMod = 0
        Me.Enabled = True
    Exit Sub
chkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Sub InitData()
    On Error GoTo chkErr
    '�ܦ��O�n���Ȧs�ɨ��e���n���B�s�ɪ���T, �H�K���<�n����ڼ�>, <�n���`���B>�����e: (��: <�n���`���B>�榡��99, 999, 999)
    '�E  <�n����ڼ�> = ��Select max(EntryNo) from SO074 where EntryEn=��<�ާ@�H���N��>�� ��
    '�E  �Y<�n����ڼ�>�j��0, �h:
    '<�n���`���B> = ��select sum(RealAmt) from SO074 where EntryEn=��<�ާ@�H���N��>�� and CancelFalg=0��
    '�_�h: <�n���`���B> = 0
    Dim rsTmp As New ADODB.Recordset
        gilCompCode.SetCodeNo strCompCode
        gilCompCode.Query_Description
        
        strSQL = "Select RowId,So033.* "
        rsTmp.MaxRecords = 1
        If Not GetRS(rsTmp, "Select Max(EntryNo) as BillCount From " & GetOwner & "So074 where EntryEn='" & garyGi(0) & "'", gcnGi) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
        If Not rsTmp.EOF Then
            If rsTmp("BillCount").Value > 0 Then
                lblBillCnt.Caption = rsTmp("BillCount").Value
                If Not GetRS(rsTmp, "Select Sum(RealAmt) as AmtCount From " & GetOwner & "So074 where EntryEn='" & garyGi(0) & "'", gcnGi) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
                lblTotalAmt.Caption = rsTmp("AmtCount").Value & ""
                GoTo CloseRS
            End If
        End If
        intDayCut = Val(GetSystemParaItem("DayCut", gilCompCode.GetCodeNo, "", "SO041") & "")
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
        lblBillCnt.Caption = 0
        lblTotalAmt.Caption = 0
CloseRS:
        rsTmp.Close
        Set rsTmp = Nothing
    Exit Sub
chkErr:
     Call ErrSub(Me.Name, "InitData")
End Sub
Private Function chkREF3() As Boolean
  On Error GoTo chkErr
    Dim strQry As String
    Dim strBillNoSQL As String
    Select Case Len(txtBillNo)
        Case 11
            strBillNoSQL = "A.MediaBillno='" & txtBillNo.Text & "'"
        Case 12
            strBillNoSQL = "A.PrtSNo='" & txtBillNo.Text & "'"
        Case Else
            strBillNoSQL = "A.Billno='" & txtBillNo.Text & "'"
    End Select
    strQry = "Select Count(*) Cnt From " & GetOwner & "SO033 A," & GetOwner & "CD013 B" & _
            " Where " & strBillNoSQL & " And A.UCCode is Not Null " & _
            " And A.CompCode=" & gCompCode & _
            " And A.UCCode=B.CodeNo" & _
            " And B.RefNo=3"
    If gcnGi.Execute(strQry)(0) > 0 Then
        chkREF3 = True
    Else
        chkREF3 = False
    End If
    Exit Function
chkErr:
  ErrSub Me.Name, "ChkREF3"
End Function
Private Function ChkDupData() As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strBillNoSQL As String
        Select Case Len(txtBillNo)
            Case 11
                strBillNoSQL = "MediaBillno='" & txtBillNo.Text & "'"
            Case 12
                strBillNoSQL = "PrtSNo='" & txtBillNo.Text & "'"
            Case Else
                strBillNoSQL = "Billno='" & txtBillNo.Text & "'"
        End Select
        strSQL = "Select BillNo From " & GetOwner & "So074 Where " & strBillNoSQL
        If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I": Exit Function
        If rsTmp.EOF Then
            ChkDupData = False
        Else
            ChkDupData = True
        End If
        rsTmp.Close
        Set rsTmp = Nothing
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "ChkDupData")
End Function

Private Sub subGil()
    On Error GoTo chkErr
        '���O�H��
        SetgiList gilClctEn, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True
        '���O�覡
        SetgiList gilCMCode, "CodeNO", "Description", "CD031", , , , , 3, 12, , True
        '�I�ں���
        SetgiList gilPTCode, "CodeNO", "Description", "CD032", , , , , 3, 12, , True
        '���q�O
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode)
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub subGrd()
    On Error GoTo chkErr
    Dim mFlds As New prjGiGridR.GiGridFlds
        With mFlds
            .Add "CustId", , , , False, "�Ȥ�s��", vbLeftJustify
            .Add "Citemname", , , , False, "    ���O����    ", vbLeftJustify
            .Add "ShouldAmt", , , , False, "�X����B", vbRightJustify
            .Add "RealPeriod", , , , False, "����", vbLeftJustify
            .Add "RealAmt", , , , False, "�ꦬ���B", vbRightJustify
            .Add "RealStartDate", giControlTypeDate, , , False, "  �����_�l��  ", vbLeftJustify
            .Add "RealStopDate", giControlTypeDate, , , False, "  �����I���  ", vbLeftJustify
            .Add "CancelFlag", , , , False, "�@�o", vbLeftJustify
            .Add "Note", , , , False, Space(20) & "  �Ƶ�            " & Space(15), vbLeftJustify
        End With
            
        ggrData.AllFields = mFlds
        ggrData.SetHead
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGrd"
End Sub

Public Sub ClearRcd()
    '2.  ClearRcd(): �Ω󥻵e���Ұʮ�, �ΨC�@���n���s�ɫ�, �N�ù����U���󤺮e�]����l���A:
    '�E  <�m�W>lblName, <�Ȥ�s��>lblCustId, <�Ȥ᪬�A>lblCustStatus �Ҭ��ť�
    '�E  �Y<�w�]���O�H��>����, �h��J<���O�H���N�X>�P<���O�H���m�W>gilClctEn
    '�E  <�J�b���>gdaRealDate = <�w�]�J�b���>
    '�E  <���O�覡>gilCMCode = <�w�]���O�覡>  (�t�N�X/�W��)
    '�E  <�I�ں���>gilPTCode = <�w�]�I�ں���>  (�t�N�X/�W��)
    '�E  �M����ڤ��egrid ggrData
    '�E  �M��[�}�C]���e
    '�E  �Y<���O�~��>����, �h<��ڽs��> = <���O�~��>�����~��+' B ', �_�h���ť�
    '�E  disable�s�W/�ק�/�@�o�T����  (��: �ثe�����ѧ@�o�\��)
        On Error GoTo chkErr
        lblName.Caption = ""
        lblCustId.Caption = ""
        lblCustStatus.Caption = ""
        lblClassName.Caption = ""
        If strClctEn <> "" Then
            gilClctEn.SetCodeNo strClctEn
            gilClctEn.SetDescription strClctName
        End If
        If strCMCode <> "" Then
            gilCMCode.SetCodeNo strCMCode
            gilCMCode.SetDescription strCMName
        End If
        If strPTCode <> "" Then
            gilPTCode.SetCodeNo strPTCode
            gilPTCode.SetDescription strPTName
        End If
        gdaRealDate.SetValue strRealDate
        txtBillNo.Text = strYM
        If intPara23 = 1 Then
            lblBillNo = "�L��Ǹ�"
            chkUseOldBillNo.Visible = False
        ElseIf intPara23 = 2 Then
            lblBillNo = "�C��渹"
            chkUseOldBillNo.Visible = False
        Else
            If strYM <> "" Then txtBillNo.Text = strYM & "B"
            lblBillNo = "��ڽs��"
        End If
        
        
        txtManualNo.Text = ""
    '    cnMDB.Execute "Delete * from So3311"
    '    If rsArray.State = adStateOpen Then rsArray.Close
        If Not GetRS(rsArrayNo, "Select * from so3311 Where 0=1", cnMDB, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        Set ggrData.Recordset = rsArrayNo
        ggrData.Refresh
        cmdDetAdd.Enabled = False
        cmdDetEdit.Enabled = False
        cmdSave.Enabled = False

    Exit Sub
chkErr:
    ErrSub Me.Name, "ClearRcd"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        strSQL = "Select RowId From " & GetOwner & "So074 Where EntryEn='" & garyGi(0) & "' Order By EntryNO"
        rsTmp.MaxRecords = 1
        If GetRS(rsTmp, strSQL) Then
            If Not rsTmp.EOF Then
                frmSO3311D.Show vbModal
                If blnCanClose = False Then Cancel = True: Exit Sub
            End If
        End If
        Set rsTmp = Nothing
        If cnMDB.State = adStateOpen Then cnMDB.Close
        Set cnMDB = Nothing
    Exit Sub
chkErr:
    ErrSub Me.Name, "Form_QueryUnload"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub gdaRealDate_Validate(Cancel As Boolean)
    On Error Resume Next
        Cancel = Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "")
End Sub

Private Function GetCustname(ByVal intCustid As Long, ByRef RetCustName, ByRef RetCustStatusCode, _
        ByRef RetCustStatusName, ByVal strServiceType As String, ByRef strClassName As String)
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        strSQL = "Select CustName,B.CustStatusCode,B.CustStatusName,A.ClassName1 From " & GetOwner & "So001 A," & GetOwner & "So002 B Where A.CustId = B.CustId And A.Custid=" & intCustid & " And B.ServiceType='" & strServiceType & "'"
        If Not GetRS(rsTmp, strSQL, gcnGi) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I": Exit Function
        If Not rsTmp.EOF Then
            RetCustName = rsTmp("CustName").Value & ""
            RetCustStatusCode = rsTmp("CustStatusCode").Value & ""
            RetCustStatusName = rsTmp("CustStatusName").Value & ""
            strClassName = rsTmp("ClassName1").Value
        End If
        Call CloseRecordset(rsTmp)
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "GetCustname")
End Function

Private Sub ggrData_DblClick()
    On Error Resume Next
        If rsArray.State = adStateClosed Then Exit Sub
        If (rsArray.EOF Or rsArray.BOF) Then Exit Sub
        With frmSO331AF
            Set .uParentForm = Me
            Set .Conn = cnMDB
            .uDBType = giAccessDb
            .EditMode = giEditModeView
            .uBillNo = rsArray("BillNO").Value
            .uPrtSNo = rsArray("PrtSno").Value & ""
            .uItem = rsArray("Item").Value
            .uRowId = rsArray("RowId").Value
            .uClctEn = rsArray("ClctEn").Value & ""
            .uClctName = rsArray("ClctName").Value & ""
            .uPTCode = IIf(IsNull(rsArray("PTCode").Value), 0, rsArray("PTCode").Value)
            .uPTName = rsArray("PTName").Value & ""
            .uRealDate = rsArray("RealDate").Value & ""
            .uCustId = rsArray("CustID").Value
            .uCustName = rsArray("CustName").Value
            gServiceType = rsArray("ServiceType").Value
            .Show vbModal
        End With
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error GoTo chkErr
        Select Case UCase(Fld.Name)
            Case "REALAMT", "SHOULDAMT"
                Value = Format(Value, "###,###,###")
'            Case "REALSTARTDATE", "REALSTOPDATE"
'                Value = Format(Value, "ee/mm/dd")
            Case "CANCELFLAG"
                Value = IIf(Value = 0, "�_", "�O")
        End Select
    Exit Sub
chkErr:
    ErrSub Me.Name, "ggrData_ShowCellData"
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3260", "SO3261") Then Exit Sub
        Call subGil
        Call GiListFilter(gilClctEn, , gilCompCode.GetCodeNo)
    
End Sub

Private Sub txtBillNo_Change()
    On Error Resume Next
    Dim lngLoop As Long, blnChange As Boolean
        
        If txtBillNo = "" Then Exit Sub
        If Len(txtBillNo) = 11 And chkUseOldBillNo.Value = 1 Then
            If InStr("BT", Mid(txtBillNo, 5, 1)) > 0 Then
                txtBillNo = GetADString(Left(txtBillNo, 4), False) & Mid(txtBillNo, 5, 1) & "C" & Format(Mid(txtBillNo, 6), "0000000")
            End If
        End If
        If Len(txtBillNo) = 15 Or Len(txtBillNo) = 12 Or Len(txtBillNo) = 11 Then txtBillNo.Tag = "": txtBillNo_Validate (False)
End Sub

Private Sub txtBillNo_GotFocus()
    On Error GoTo chkErr
        txtBillNo.SelStart = 0
        txtBillNo.SelLength = Len(txtBillNo.Text)
    Exit Sub
chkErr:
    ErrSub Me.Name, "txtBillNo_GotFocus"
End Sub

Private Sub txtBillNo_KeyPress(keyAscii As Integer)
    On Error GoTo chkErr
        keyAscii = Asc(UCase(Chr(keyAscii)))
'        If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 98 Or _
'                    KeyAscii = 66 Or KeyAscii = 105 Or KeyAscii = 73 Or KeyAscii = 116 Or KeyAscii = 84 Then
'                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Else
'                    KeyAscii = 9
'        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "txtBillNo_KeyPress"
End Sub

Private Sub txtBillNo_Validate(Cancel As Boolean)
    On Error GoTo chkErr
    Dim lngChk As Long
    Dim lngBillNo As String
    Dim strQuerySql As String
    Dim intCustStatus As Integer
    Dim strCustStatusname As String
    Dim strClassName As String
    Dim strQryCD013 As String
        If Len(txtBillNo) = 11 And chkUseOldBillNo.Value = 1 Then
            If InStr("BT", Mid(txtBillNo, 5, 1)) > 0 Then
                txtBillNo = GetADString(Left(txtBillNo, 4), False) & Mid(txtBillNo, 5, 1) & "C" & Format(Mid(txtBillNo, 6), "0000000")
            End If
        End If
        If txtBillNo.Tag <> "" Then txtBillNo.Tag = "": Exit Sub
        txtBillNo.Tag = "Do"
        If Len(txtBillNo.Text) < 11 Then Exit Sub
        If Len(txtBillNo.Text) <> txtBillNo.MaxLength Then Exit Sub
        Select Case Len(txtBillNo.Text)
            '�Y���׬�15�X�B��7�X��'B', �h����ڽs��:  (�榡YYYYMMB99999999)
            Case 15
                If InStr(1, txtBillNo.Text, "B", vbTextCompare) = 7 Then
                    lngSelMod = 1
                    '�����s�����k��7�X�ର�Ʀr, ���Y��<�Ȥ�s��>, �ھڦ��Ƚs�ܫȤ�򥻸���ɨ�<�Ȥ�m�W>�P<�Ȥ᪬�A>
                    lngBillNo = Right(Trim(txtBillNo.Text), 7)
                    If Not GetRS(rsSo3311B, "Select CustName,Classname1 From " & GetOwner & "So001 Where Custid=" & lngBillNo & " And CompCode = " & strCompCode, gcnGi) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
                    ', �Y�d�L���Ƚs, �h���~: "�L����ڤ��Ȥ���", �õL�H�U�ʧ@. �Y�����Ƚs
                    If rsSo3311B.EOF Then
                        MsgBox "�L����ڤ��Ȥ��ơI", vbExclamation, "�T���I"
                        Cancel = True
                        Call ClearRcd
                        Exit Sub
                    Else
                        strCustName = rsSo3311B("CustName").Value & ""
                        strClassName = rsSo3311B("ClassName1") & ""
                        '2001/11/12 Janis ���n�諸
                        strQuerySql = strSQL & " From " & GetOwner & "So033 Where Billno='" & Trim(txtBillNo.Text) & "' And CompCode = " & strCompCode & " and UCCOde Is Not NUll "
                        If rsSo3311B.State = adStateOpen Then rsSo3311B.Close
                        If Not GetRS(rsSo3311B, strQuerySql, gcnGi) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
                        
                        If rsSo3311B.EOF Then MsgBox "�L����ڽs���Φ���ڤw�n���L�A�Юֹ�I", vbExclamation, "�T���I": Cancel = True: Call ClearRcd: Exit Sub
                        Call GetCustname(rsSo3311B("Custid").Value, strCustName, intCustStatus, strCustStatusname, rsSo3311B("ServiceType"), strClassName)
                    End If
                    
                    
                ElseIf InStr(1, txtBillNo.Text, "T", vbTextCompare) = 7 Then
                    lngSelMod = 1
                    '�E  �Y���׬�15�X, �B��7�X 'T', �h���{�ɳ�s��:   (�榡YYYYMMT99999999)
                    '2001/11/12 Janis ���n�諸
                    strQuerySql = strSQL & " From " & GetOwner & "SO033 where Billno='" & UCase(txtBillNo.Text) & "' And CompCode = " & strCompCode & " and UCCode is not null "
                    If Not GetRS(rsSo3311B, strQuerySql, gcnGi) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I":   Exit Sub
                    '�Y�L����ڽs�����X�Ƶ{��
                    If rsSo3311B.RecordCount = 0 Then MsgBox "�L����ڽs���Φ���ڤw�n���L�A�Юֹ�I", vbExclamation, "�T���I": Cancel = True: Call ClearRcd: Exit Sub
                    '�Y��J�ȫDB��榡 (�ӬO�{�ɳ�ΦL��Ǹ��榡), �h�ھڨ��X�����O��ƲĤ@�����Ȥ�s��, �ܫȤ�򥻸���ɨ�
                    '�Y��ڽs���s�b�h�̫Ȥ�s���ܡ]SO001)���Ȥ�m�W�ΫȤ᪬�A
                    Call GetCustname(rsSo3311B("Custid").Value, strCustName, intCustStatus, strCustStatusname, rsSo3311B("ServiceType"), strClassName)
                Else
                    MsgBox "��ڽs�����~�I", vbExclamation, "�T���I": Cancel = True: Call ClearRcd: Exit Sub
                End If
            Case 12
                '�E  �Y���׬�12�X, �h���L��Ǹ�: (�榡YYYYMM999999)
                'SQL = "select <�K> from SO033 where PrtSNo='<��ڽs��>' and UCCode is not null "
                'lngBillNo = Right(txtBillNo.Text, 8)
                '2001/11/12 Janis ���n�諸
                strQuerySql = strSQL & " From " & GetOwner & "SO033 where PrtSNo='" & UCase(txtBillNo.Text) & "' And CompCode = " & strCompCode & " and UCCode is not null "
                lngSelMod = 2
                If Not GetRS(rsSo3311B, strQuerySql, gcnGi) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
                If rsSo3311B.EOF Then
                    If cboBillType.ListIndex = 1 Then
                        MsgBox "�L����ڽs���Φ���ڤw�n���L�A�Юֹ�I", vbExclamation, "�T���I"
                        Cancel = True
                        Call ClearRcd
                    End If
                    Exit Sub
                Else
    '                '�Y��J�ȫDB��榡 (�ӬO�{�ɳ�ΦL��Ǹ��榡), �h�ھڨ��X�����O��ƲĤ@�����Ȥ�s��, �ܫȤ�򥻸���ɨ�
                    Call GetCustname(rsSo3311B("Custid").Value, strCustName, intCustStatus, strCustStatusname, rsSo3311B("ServiceType"), strClassName)
                End If
            Case 11
                '�E  �Y���׬�11�X, �h���C��渹: (�榡YYMM9999999)
                'SQL = "select <�K> from SO033 where MediaBillNo='<��ڽs��>' and UCCode is not null "
                strQuerySql = strSQL & " From " & GetOwner & "SO033 where MediaBillNo ='" & UCase(txtBillNo.Text) & "' And CompCode = " & strCompCode & " and UCCode is not null "
                lngSelMod = 3
                If Not GetRS(rsSo3311B, strQuerySql, gcnGi) Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
                If rsSo3311B.EOF Then
                    If cboBillType.ListIndex = 2 Then
                        MsgBox "�L����ڽs���Φ���ڤw�n���L�A�Юֹ�I", vbExclamation, "�T���I"
                        Cancel = True
                        Call ClearRcd
                    End If
                    Exit Sub
                Else
    '                '�Y��J�ȫDB��榡 (�ӬO�{�ɳ�ΦL��Ǹ��榡), �h�ھڨ��X�����O��ƲĤ@�����Ȥ�s��, �ܫȤ�򥻸���ɨ�
                    Call GetCustname(rsSo3311B("Custid").Value, strCustName, intCustStatus, strCustStatusname, rsSo3311B("ServiceType"), strClassName)
                End If
            Case Else
                    If txtBillNo.Text <> "" Then
                            MsgBox "��ڽs�����~�I", vbExclamation, "�T���I"
                            Cancel = True
                    End If
                    
                    Exit Sub
        End Select
            '�ˬd��ڬO�_�w�n���I
            If ChkDupData Then
                MsgBox "����ڤw�n���L�I�Э��s��J�I", vbExclamation, "�T���I"
                Cancel = True
                txtBillNo.Text = ""
                Exit Sub
            End If
            '********************************************************************
            '#4010 �ˬd���UCCode�O�_���ѦҬ�3 By Kin 2008/08/06
            If chkREF3 Then
                MsgBox "������d�O�w���I�Э��s��J�I", vbExclamation, "�T���I"
                Cancel = True
                txtBillNo.Text = ""
                Exit Sub
            End If
            '********************************************************************
        '�N�Ȥ�s���B�W�١B���A��ܦb�ù��W
            intCustid = rsSo3311B("CustID").Value
            lblCustId.Caption = intCustid
            lblName.Caption = strCustName
            lblCustStatus.Caption = strCustStatusname
            lblClassName.Caption = strClassName
        
        '�E  �Y<�Ȥ᪬�A�N�X>����1(���`), �hĵ�i: "���Ȥ᪬�A�w�D���`������, ���ˬd"
        If intCustStatus <> 1 Then MsgBox "���Ȥ᪬�A�w�D���`������, ���ˬd", vbExclamation, "�T���I"
        
        '�E  �Y�����, �h�s�W�ܦ��O��Ƶn���Ȧs��(SO074)
        '�E  �Y��7�X���O 'T'�B�����, �henable�s�W/�ק�/�@�o�T����
        If Mid(txtBillNo.Text, 7, 1) <> "T" Then
            cmdDetAdd.Enabled = True
            cmdDetEdit.Enabled = True
            cmdDetDelete.Enabled = True
        End If
            
        If Not rsSo3311B.EOF Then
        '�E  Loop���쪺�C�@�����, ���H�U�ʧ@, �÷s�W�@����Ʀ�[�}�C], ��줺�e�p�U:
        'O �H�U��줺�e�Y���ӵ����������:
        '��ڽs�� , ����, �Ȥ�s��, ���O���إN�X / �W��, �X����B, �������, ���İ_�l / �I����, �@�o
        'O   �ꦬ��� = <�J�b���>
        'O   �Y<�w�]�ꦬ���B>����, �h�ꦬ���B = <�w�]�ꦬ���B>, �_�h�ꦬ���B = �ӵ����X����B
        'O   �Y<�w�]�ꦬ����>����, �h: ���O���� = <�w�]�ꦬ����>, �_�h: ���O���� = �ӵ������O����
        'O   �Ȥ�m�W = <�e���W���Ȥ�m�W>
        'O �n���H���N�� = �ާ@�̥N��
        'O   �Ƶ� = �ӵ����Ƶ� + <�w�]�Ƶ���>
        'O   �Y<�w�]���O�H��>����, �h: ���O�H���N��/�W�� = <�w�]���O�H��>���N��/�W��, �_�h: ���O�H���N��/�W�� = �ӵ������O�H���N��/�W��
        'O �I�ں����N�X / �W�� <= �w�]�I�ں��� > ���N�X / �W��
        'O ���O�覡�N�X / �W�� <= �w�]���O�覡 > ���N�X / �W��
        'O   �Y<�w�]�u����]>����, �h: �u����]�N�X/�W�� = <�w�]�u����]>���N�X/�W��, �_�h: �u����]�N�X/�W�� = �ŭ�
        'O ��ƧǸ� = �ӵ���RowId
        'O   �s�W�� [�}�C]
            On Error GoTo lblRollback
            cnMDB.BeginTrans
            cnMDB.Execute "Delete * from So3311"
            If Not GetRS(rsArray, "Select * From SO3311", cnMDB, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
            rsSo3311B.MoveFirst
            While Not rsSo3311B.EOF
                With rsArray
                    .AddNew
                    .Fields("BillNo").Value = GetFieldValue(rsSo3311B, "BillNO")
                    .Fields("MediaBillNo").Value = GetFieldValue(rsSo3311B, "MediaBillNO")
                    .Fields("PrtsNo").Value = GetFieldValue(rsSo3311B, "PrtSno")
                    .Fields("Item").Value = GetFieldValue(rsSo3311B, "Item")
                    .Fields("CustName").Value = strCustName
                    .Fields("CustId").Value = GetFieldValue(rsSo3311B, "Custid")
                    .Fields("CItemCode").Value = GetFieldValue(rsSo3311B, "CitemCode")
                    .Fields("CItemName").Value = GetFieldValue(rsSo3311B, "CitemName")
                    .Fields("ShouldAmt").Value = Val(rsSo3311B("ShouldAmt").Value)
                    .Fields("ShouldDate").Value = GetFieldValue(rsSo3311B, "ShouldDate")
    '                If lngPeriod > 0 Then
    '                    .Fields("RealPeriod").Value = lngPeriod
    '                Else
    '                    .Fields("RealPeriod").Value = rsSo3311B("RealPeriod").Value
    '                End If
                    .Fields("EntryEN").Value = garyGi(0)
                    .Fields("Note").Value = rsSo3311B("Note").Value & strNote
                    .Fields("CancelFlag").Value = GetFieldValue(rsSo3311B, "CancelFlag")
                    .Fields("RealDate").Value = IIf(Len(gdaRealDate.GetValue) = 0, RightDate, gdaRealDate.GetValue(True))
                    If gilClctEn.GetCodeNo <> "" Then
                        .Fields("ClctEn").Value = gilClctEn.GetCodeNo2
                        .Fields("ClctName").Value = gilClctEn.GetDescription2
                    Else
                        If Len(rsSo3311B("ClctEn").Value & "") = 0 Then
                            .Fields("ClctEn").Value = GetFieldValue(rsSo3311B, "OldClctEN")
                            .Fields("ClctName").Value = GetFieldValue(rsSo3311B, "OldClctName")
                        Else
                            .Fields("ClctEn").Value = GetFieldValue(rsSo3311B, "ClctEN")
                            .Fields("ClctName").Value = GetFieldValue(rsSo3311B, "ClctName")
                        End If
                    End If
                    If gilCMCode.GetCodeNo <> "" Then
                        .Fields("CmCode").Value = gilCMCode.GetCodeNo2
                        .Fields("CMName").Value = gilCMCode.GetDescription2
                    Else
                        .Fields("CmCode").Value = NoZero(rsSo3311B("CMCode"))
                        .Fields("CMName").Value = NoZero(rsSo3311B("CMName"))
                    End If
                    If gilPTCode.GetCodeNo <> "" Then
                        .Fields("PTCode").Value = gilPTCode.GetCodeNo2
                        .Fields("PTName").Value = gilPTCode.GetDescription2
                    Else
                        .Fields("PTCode").Value = NoZero(rsSo3311B("PTCode"))
                        .Fields("PTName").Value = NoZero(rsSo3311B("PTName"))
        
                    End If
                    .Fields("RowID").Value = GetFieldValue(rsSo3311B, "RowID")
                    If lngRealAmt <> "" Then
                        .Fields("RealAmt").Value = lngRealAmt
                        '.Fields("ShouldAmt").Value = lngRealAmt
                    Else
                        .Fields("RealAmt").Value = NoZero(rsSo3311B("ShouldAmt"), True)
                    End If
                    
                    If .Fields("RealAmt") <> .Fields("ShouldAmt") Then
                        If strSTName <> "" Then
                            .Fields("STCode").Value = strSTCode
                            .Fields("STName").Value = strSTName
                        Else
                            .Fields("STCode").Value = NoZero(rsSo3311B("STCode"))
                            .Fields("STName").Value = NoZero(rsSo3311B("STName"))
                        End If
                    End If
                    .Fields("CancelCode") = Null
                    .Fields("CancelName") = Null
'                    If txtManualNo <> "" Then
'                        .Fields("ManualNo").Value = txtManualNo.Text
'                    Else
'                        .Fields("ManualNo").Value = GetFieldValue(rsSo3311B, "ManualNo")
'                    End If

                    '***********************************************************************************************
                    '#3163 �W�[��}�渹���ˮ� By Kin 2007/12/20
                    If txtManualNo <> "" Then
                        If Not ChkManualNoOk(txtManualNo, rsSo3311B("BillNo") & "", rsSo3311B("ServiceType")) Then
                            txtBillNo.Tag = ""
                            cnMDB.RollbackTrans
                            ClearRcd
                            txtManualNo.SetFocus
                            Exit Sub
                        Else
                            '#4021 �p�G�ˮ֦��q�L���ܭn�^��^�h By Kin 2008/07/21
                            .Fields("ManualNo") = txtManualNo.Text
                        End If
                    Else
                        .Fields("ManualNo").Value = GetFieldValue(rsSo3311B, "ManualNo")
                    End If
                    '***********************************************************************************************
                    
                    .Fields("CompCode") = GetFieldValue(rsSo3311B, "CompCode")
                    .Fields("ServiceType") = GetFieldValue(rsSo3311B, "ServiceType")
                    .Fields("EntryEn").Value = garyGi(0)
                    .Fields("RealStartDate").Value = GetFieldValue(rsSo3311B, "RealStartDate")
                    
                    strRowId = GetFieldValue(rsSo3311B, "RowId")
                    If lngPeriod > 0 Then
                        .Fields("RealPeriod").Value = lngPeriod
                        Dim strMsg As String
                        Dim strRealStartDate As String
                        Dim strRealStopDate As String
                        Dim lngRealAmount As Long
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
                        '#4346 �p�G�w�]���Ƭ�0,���n��OldPeriod Upd�^�h By Kin 2009/02/20
                        .Fields("RealPeriod").Value = NoZero(rsSo3311B("RealPeriod") & "", True)
                        '.Fields("RealPeriod").Value = IIf(Val(rsSo3311B("RealPeriod") & "") > 0, NoZero(rsSo3311B("RealPeriod"), True), NoZero(rsSo3311B("OldPeriod"), True))
                    End If
                    If Len(strRealStopDate) <> 0 Then
                        .Fields("RealStopDate").Value = Format(strRealStopDate, "####/##/##")
                    Else
                        '#4346 ���n��OldStopDate Update�^�h By Kin 2009/02/20
                        .Fields("RealStopDate").Value = IIf(Len(rsSo3311B("RealStopDate").Value & "") = 0, Null, NoZero(rsSo3311B("RealStopDate")))
                        '.Fields("RealStopDate").Value = IIf(Len(rsSo3311B("RealStopDate").Value & "") = 0, NoZero(rsSo3311B("OldStopDate")), NoZero(rsSo3311B("RealStopDate")))
                    End If
                    '#3026 �����ˬd,�J�b��,�n�P�_���O�覡�N�X���@�ˮ�, �n�M���Ȧ�O�αb��
'                    If .Fields("CMCode").Value & "" <> rsSo3311B("CMCode") & "" Or GetRsValue("Select Count(*) From " & GetOwner & "SO002A Where CustId = " & _
'                        .Fields("CustId") & " And AccountNo = '" & rsSo3311B("AccountNo") & "' And CompCode = " & .Fields("CompCode")) = 0 And rsSo3311B("AccountNo") & "" <> "" Then
'                        .Fields("BankCode") = Null
'                        .Fields("BankName") = Null
'                        .Fields("AccountNo") = Null
'                    Else
                        .Fields("BankCode") = rsSo3311B("BankCode")
                        .Fields("BankName") = rsSo3311B("BankName")
                        .Fields("AccountNo") = rsSo3311B("AccountNo")
'                    End If
'                    .Fields("AuthorizeNo") = rsSo3311B("AuthorizeNo")
'                    .Fields("AdjustFlag") = rsSo3311B("AdjustFlag")
                     .Fields("FaciSeqNo") = rsSo3311B("FaciSeqNo")
                     
                     '******************************************************
                     '#3436 �x�s�ҨϥΪ��o����T By Kin 2007/12/17
                     .Fields("InvSeqNo") = rsSo3311B("InvSeqNo")
                     '******************************************************
                     '#6306 NextPeriod�BNextAmt�n�H��ڸ�Ƭ��D By Kin 2012/09/07
                     If Len(rsSo3311B("NextPeriod")) > 0 Then
                        .Fields("NextPeriod") = rsSo3311B("NextPeriod")
                     End If
                     If Len(rsSo3311B("NextAmt")) > 0 Then
                        .Fields("NextAmt") = rsSo3311B("NextAmt")
                     End If
                    .Update
                End With
                rsSo3311B.MoveNext
           Wend
           cnMDB.CommitTrans
           On Error GoTo chkErr
           Call RefreshData
           txtBillNo_GotFocus
           
           cmdSave.Enabled = True
        End If
    
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "txtBillNo_Validate")
    Exit Sub
lblRollback:
    cnMDB.RollbackTrans
    ErrSub Me.Name, "txtBillNo_Validate, ���ʥ��ѡI�Э��s�ާ@�I"
End Sub

Private Sub RefreshData()
    On Error GoTo chkErr
        If rsArray.State = adStateOpen Then rsArray.Close
        rsArray.CursorLocation = adUseClient
        rsArray.Open "Select * from So3311", cnMDB, adOpenForwardOnly, adLockOptimistic
        Set ggrData.Recordset = rsArray
        ggrData.Refresh
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "RefReshData")
End Sub

'�w�]����(3311A)
Public Property Let uPeriod(ByVal vData As Long)
    On Error GoTo chkErr
        lngPeriod = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uPeriod"
End Property

'�w�]�Ƶ���(3311A)
Public Property Get uNote() As String
    On Error GoTo chkErr
        uNote = strNote
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uNote"
End Property

'�w�]�Ƶ���(3311A)
Public Property Let uNote(ByVal vData As String)
    On Error GoTo chkErr
        strNote = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uNote"
End Property

'�w�]���O�H���m�W
Public Property Get uClctName() As String
    On Error GoTo chkErr
        uClctName = strClctName
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uClctName"
End Property

'�����W�h�����O�H���m�W
Public Property Let uClctName(ByVal vData As String)
    On Error GoTo chkErr
        strClctName = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uClctName"
End Property

'�w�]���O�H
Public Property Get uClctEn() As String
    On Error GoTo chkErr
        uClctEn = strClctEn
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uClctEn"
End Property

'�����W�h�����O�H��
Public Property Let uClctEn(ByVal vData As String)
    On Error GoTo chkErr
        strClctEn = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uClctEn"
End Property

'�]�w���J�b���
Public Property Get uRealDate() As String
    On Error GoTo chkErr
        uRealDate = strRealDate
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uRealDate"
End Property

'�����W�h���w�]�J�b���
Public Property Let uRealDate(ByVal vDate As String)
    On Error GoTo chkErr
    strRealDate = vDate
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uRealDate"
End Property

'�����W�h�ǨӤ�TranDate
Public Property Let uTranDate(ByVal vDate As String)
    On Error GoTo chkErr
    strTranDate = vDate
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uTranDate"
End Property

'�����W�h�ǨӤ�<para6>���Ѽ�
Public Property Let uPara6(ByVal vDate As Integer)
    On Error GoTo chkErr
    intPara6 = vDate
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uPara6"
End Property

'�w�]���O�覡
Public Property Get uCMCode() As String
    On Error GoTo chkErr
    uCMCode = strCMCode
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uCMCode"
End Property

'�����W�h�ǨӤ��w�]���O�覡
Public Property Let uCMCode(ByVal vData As String)
    On Error GoTo chkErr
    strCMCode = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uCMCode"
End Property

'�w�]���O�覡�W��
Public Property Get uCMName() As String
    On Error GoTo chkErr
    uCMName = strCMName
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uCMName"
End Property

'�����W�ǨӤ��w�]���O�覡�W��
Public Property Let uCMName(ByVal vData As String)
    On Error GoTo chkErr
    strCMName = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uCMName"
End Property

'�w�]�I�ں���
Public Property Get uPTCode() As String
    On Error GoTo chkErr
    uPTCode = strPTCode
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uPTCode"
End Property

'�����W�h�ǨӤ��w�]�I�ں���
Public Property Let uPTCode(ByVal vData As String)
    On Error GoTo chkErr
    strPTCode = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uPTCode"
End Property

'�w�]�I�ں����W��
Public Property Get uPTName() As String
    On Error GoTo chkErr
    uPTName = strPTName
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uPTName"
End Property

'�����W�h�ǨӤ��w�]�I�ں����W��
Public Property Let uPTName(ByVal vData As String)
    On Error GoTo chkErr
    strPTName = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uPTName"
End Property

'�����W�h�ǨӤ��w�]�ꦬ���B
Public Property Let uRealAmt(ByVal vData As String)
    On Error GoTo chkErr
    lngRealAmt = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uRealAmt"
End Property

'�����W�h�ǨӤ��w�]�u����]
Public Property Let uStCode(ByVal vData As String)
    On Error GoTo chkErr
        strSTCode = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uStCode"
End Property

'�����W�h�ǨӤ��w�]�u����]�W��
Public Property Let uSTName(ByVal vData As String)
    On Error GoTo chkErr
        strSTName = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uSTName"
End Property

'�����W�h�ǨӤ��]�w���O�~��I
Public Property Let uYM(ByVal vData As String)
    On Error GoTo chkErr
        strYM = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uYM"
End Property

'�����W�h����}�渹�H�K�Ǧ�So3311D
Public Property Let uManualNo(ByVal vData As String)
    On Error GoTo chkErr
        strManualNo = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uManualNo"
End Property

'�����W�h����}�渹
Public Property Get uManualNo() As String
    On Error GoTo chkErr
        uManualNo = strManualNo
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uManualNo"
End Property

'�����W�h���۰ʧ�s�g�����ت����B
Public Property Get uOption1() As Boolean
    On Error GoTo chkErr
        uOption1 = blnOption1
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uOption1"
End Property

'�����W�h���۰ʧ�s�g�����ت����B
Public Property Let uOption1(ByVal vData As Boolean)
    On Error GoTo chkErr
        blnOption1 = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uOption1"
End Property

'�����W�h���۰ʧ�s�g�����ؤ�������
Public Property Get uOption2() As Boolean
    On Error GoTo chkErr
        uOption2 = blnOption2
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uOption2"
End Property

'�����W�h���۰ʧ�s�g�����ؤ�������
Public Property Let uOption2(ByVal vData As Boolean)
    On Error GoTo chkErr
        blnOption2 = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uOption2"
End Property

'�����W�h���۰ʷs�W�g���ʸ�Ƹ��
Public Property Get uOption3() As Boolean
    On Error GoTo chkErr
        uOption3 = blnOption3
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uOption3"
End Property

'�����W�h���۰ʷs�W�g���ʸ�Ƹ��
Public Property Let uOption3(ByVal vData As Boolean)
    On Error GoTo chkErr
        blnOption3 = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uOption3"
End Property

'�]�w�O�_���L��Ǹ��γ�ڽs��
Public Property Get uSelBill() As Integer
    On Error GoTo chkErr
       uSelBill = lngSelMod
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uSelBill"
End Property

'�Ȥ�s��
Public Property Get uCustId() As Integer
    On Error GoTo chkErr
        uCustId = intCustid
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uCustID"
End Property

'�Ȥ�m�W
Public Property Get uCustName() As String
    On Error GoTo chkErr
        uCustName = strCustName
    Exit Property
chkErr:
    ErrSub Me.Name, "Get uCustName"
End Property

'�i���_
Public Property Let uCanClose(ByVal vData As Boolean)
    On Error GoTo chkErr
        blnCanClose = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uCanClose"
End Property

'���q�O
Public Property Let uCompCode(ByVal vData As String)
    On Error GoTo chkErr
        strCompCode = vData
    Exit Property
chkErr:
    ErrSub Me.Name, "Let uCompCode"
End Property

Private Sub txtManualNo_Change()
    On Error Resume Next
    strManualNo = txtManualNo.Text
End Sub

Private Sub txtManualNo_GotFocus()
    On Error Resume Next
        txtManualNo.SelStart = 0
        txtManualNo.SelLength = Len(txtManualNo.Text)
End Sub
