VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.17#0"; "csMulti.ocx"
Begin VB.Form frmSO3254A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "���O�覡�վ� [SO3254A]"
   ClientHeight    =   6120
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
   Icon            =   "SO3254A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2. �T�w"
      Height          =   375
      Left            =   7920
      TabIndex        =   14
      Top             =   5460
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   10230
      TabIndex        =   15
      Top             =   5460
      Width           =   1365
   End
   Begin VB.Frame fraData 
      Height          =   5985
      Left            =   120
      TabIndex        =   16
      Top             =   30
      Width           =   11655
      Begin Gi_Multi.GiMulti gmsClctArea 
         Height          =   375
         Left            =   90
         TabIndex        =   6
         Top             =   2010
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   661
         ButtonCaption   =   "���O��"
      End
      Begin Gi_Multi.GiMulti gmsServArea 
         Height          =   375
         Left            =   90
         TabIndex        =   5
         Top             =   1620
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   661
         ButtonCaption   =   "�A�Ȱ�"
      End
      Begin CS_Multi.CSmulti gmsStrtCode 
         Height          =   405
         Left            =   90
         TabIndex        =   9
         Top             =   3180
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   714
         ButtonCaption   =   "��D�d��"
      End
      Begin CS_Multi.CSmulti gmsMduId 
         Height          =   405
         Left            =   90
         TabIndex        =   8
         Top             =   2790
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   714
         ButtonCaption   =   "�j�ӦW��"
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "�s��������"
         Height          =   375
         Left            =   5940
         TabIndex        =   13
         Top             =   5430
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskPrtSNo2 
         Height          =   345
         Left            =   3000
         TabIndex        =   2
         ToolTipText     =   "YYYYMMxxxxxx"
         Top             =   630
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         Mask            =   "999999999999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPrtSNo1 
         Height          =   345
         Left            =   1290
         TabIndex        =   1
         ToolTipText     =   "YYYYMMxxxxxx"
         Top             =   630
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   12
         Mask            =   "999999999999"
         PromptChar      =   "_"
      End
      Begin prjGiList.GiList gilCitemCode 
         Height          =   315
         Left            =   1410
         TabIndex        =   10
         Top             =   3720
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F2Corresponding =   -1  'True
      End
      Begin Gi_YM.GiYM gdaClctYM2 
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
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
      Begin Gi_YM.GiYM gdaClctYM1 
         Height          =   375
         Left            =   1290
         TabIndex        =   3
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
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
      Begin prjGiList.GiList gilOldCMCode 
         Height          =   315
         Left            =   1410
         TabIndex        =   11
         Top             =   4230
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilNewCMCode 
         Height          =   315
         Left            =   1410
         TabIndex        =   12
         Top             =   4680
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilCompCode 
         Height          =   315
         Left            =   1290
         TabIndex        =   0
         Top             =   210
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth1       =   600
         FldWidth2       =   2000
         F5Corresponding =   -1  'True
      End
      Begin CS_Multi.CSmulti gmsServEmp 
         Height          =   345
         Left            =   90
         TabIndex        =   7
         Top             =   2400
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   609
         ButtonCaption   =   "���O�H��"
      End
      Begin prjGiList.GiList gilServiceType 
         Height          =   315
         Left            =   5400
         TabIndex        =   25
         Top             =   210
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F3Corresponding =   -1  'True
      End
      Begin VB.Label lblServiceType 
         AutoSize        =   -1  'True
         Caption         =   "�A�����O"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4500
         TabIndex        =   26
         Top             =   270
         Width           =   780
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "��  �q  �O"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   330
         TabIndex        =   24
         Top             =   300
         Width           =   765
      End
      Begin VB.Label lblNewCMCode 
         AutoSize        =   -1  'True
         Caption         =   "�s���O�覡"
         Height          =   195
         Left            =   330
         TabIndex        =   23
         Top             =   4770
         Width           =   975
      End
      Begin VB.Label lblNote2A 
         AutoSize        =   -1  'True
         Caption         =   "�����~��"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   330
         TabIndex        =   22
         Top             =   1170
         Width           =   780
      End
      Begin VB.Label lblNote2B 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   2370
         TabIndex        =   21
         Top             =   1200
         Width           =   195
      End
      Begin VB.Label lblOldCMCode 
         AutoSize        =   -1  'True
         Caption         =   "�즬�O�覡"
         Height          =   195
         Left            =   330
         TabIndex        =   20
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label lblNote1B 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   2730
         TabIndex        =   19
         Top             =   750
         Width           =   195
      End
      Begin VB.Label lblNote1A 
         AutoSize        =   -1  'True
         Caption         =   "�L��Ǹ�"
         Height          =   195
         Left            =   330
         TabIndex        =   18
         Top             =   750
         Width           =   780
      End
      Begin VB.Label lblCitemCode 
         AutoSize        =   -1  'True
         Caption         =   "���O����"
         Height          =   195
         Left            =   330
         TabIndex        =   17
         Top             =   3795
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmSO3254A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���\��i�N������������ɤ�(SO033), �ŦX�����������, �N��X����B���妸�վ�
Option Explicit

Private Sub cmdCancel_Click()
    On Error GoTo ChkErr
        Unload Me
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdCancel_Click"
End Sub

'�E  �̾کҤU���Ѽ�, �ꦨSQL���O�ܫ�ݰ��妸�վ�:
'�E  �Y��/�s�O�Τ��۵��B��/�s���Ƥ��۵�, SQL:
'    "update SO033 set ShouldAmt = <�s�O��>, OldPeriod = <�s����> where <�U�ѼƱ���> and RealDate is null "
'�E  �Y�u����/�s�O�Τ��۵�, SQL:
'    "update SO033 set ShouldAmt = <�s�O��> where <�U�ѼƱ���> and RealDate is null "
'�E  �Y�u����/�s���Ƥ��۵�, SQL:
'    "update SO033 set OldPeriod = <�s����> where <�U�ѼƱ���> and RealDate is null"
'�E  �Y���\, �h���"�վ㧹��, ����=<�Q���ʪ���Ƶ���>"

Private Sub cmdOk_Click()
    On Error GoTo ChkErr
    Dim lngAffectCount As Long
    Dim strSQL As String
    Dim strUpdTime As String
        If Not IsDataOk Then Exit Sub
        If MsgBox("�T�w�վ�?", vbQuestion + vbYesNo, "����") = vbNo Then Exit Sub
        strSQL = GetParaStr
        strUpdTime = GetDTString(RightNow)
        Screen.MousePointer = vbHourglass
        gcnGi.BeginTrans
        If ExecuteSQL("update " & GetOwner & "SO033 set CMCode = " & gilNewCMCode.GetCodeNo & ",CMName = '" & gilNewCMCode.GetDescription & _
                "'  where " & strSQL & "  and RealDate is null ", gcnGi, lngAffectCount) <> giOK Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I": GoTo ChkErr
        
        
'        If gnbOldAmt.Value <> gnbNewAmt.Value And gnbOldPeriod.Value <> gnbNewPeriod.Value Then
'            If ExecuteSQL("update SO033 set UpdTime='" & strUpdTime & "',Upden='" & garyGi(1) & "', ShouldAmt =" & NoZero(gnbNewAmt.Value, True) & "  , RealPeriod =" & NoZero(gnbNewPeriod.Value, True) & "  where " & strSQL & "  and RealDate is null ", gcnGi, lngAffectCount) <> giOK Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
'        ElseIf gnbOldAmt.Value <> gnbNewAmt.Value And gnbOldPeriod.Value = gnbNewPeriod.Value Then
'                    If ExecuteSQL("update SO033 set UpdTime='" & strUpdTime & "',Upden='" & garyGi(1) & "', ShouldAmt =" & NoZero(gnbNewAmt.Value, True) & " where " & strSQL & " and RealDate is null ", gcnGi, lngAffectCount) <> giOK Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
'            ElseIf gnbOldAmt.Value = gnbNewAmt.Value And gnbOldPeriod.Value <> gnbNewPeriod.Value Then
'                             If ExecuteSQL("update SO033 set UpdTime='" & strUpdTime & "',Upden='" & garyGi(1) & "', RealPeriod =" & NoZero(gnbNewPeriod.Value, True) & " where " & strSQL & " and RealDate is null", gcnGi, lngAffectCount) <> giOK Then MsgBox "�s�u���ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
'        End If
        Screen.MousePointer = vbDefault
        gcnGi.CommitTrans
        MsgBox "�վ㧹���A����=" & IIf(lngAffectCount <= 0, 0, lngAffectCount), vbInformation, " �T���I"
    Exit Sub
ChkErr:
    Screen.MousePointer = vbDefault
    gcnGi.RollbackTrans
    MsgBox "���ʥ��ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
    Unload Me
End Sub

Private Sub cmdView_Click()
'�O�d
    On Error GoTo ChkErr
    Dim rsView As New ADODB.Recordset
    Dim strSQL As String
        If Not IsDataOk Then Exit Sub
        strSQL = GetParaStr
'        rsView.CursorLocation = adUseClient
'        rsView.Open "Select CustId,CustName,Tel1,ChargeAddress,CustStatusName From " & GetOwner & "SO001 Where " & strSQL & " and RealDate is null", gcnGi, adOpenForwardOnly, adLockReadOnly
        If Not GetRS(rsView, "Select CustId,CustName,Tel1,ChargeAddress,CustStatusName From " & GetOwner & "SO001 Where " & strSQL & " and RealDate is null", gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then Exit Sub
        With ViewForm
                .uRecordset = rsView
                .Show vbModal
        End With
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdView_Click"
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ChkErr
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF2 Then Call cmdOk_Click: Exit Sub
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
    On Error Resume Next
        If Alfa2 Then
            Call GetGlobal
        End If
        Call subGim
        Call subGil
End Sub

Private Sub subGim()
    On Error Resume Next
'    '���O�覡
'        SetgiMulti gimCMCode, "CodeNo", "Description", "CD031", "�N�X", "�W��"
    '�A�Ȱ�
        SetgiMulti gmsServArea, "CodeNo", "Description", "CD002", "�N�X", "�W��", , True
    '���O��
        SetgiMulti gmsClctArea, "CodeNo", "Description", "CD040", "�N�X", "�W��", , True
    '���O�H��
        SetgiMulti gmsServEmp, "EmpNo", "EmpName", "CM003", "�N�X", "�W��", , True
    '�j��
        SetgiMulti gmsMduId, "Mduid", "Name", "SO017", "�j�ӽs��", "�j�ӦW��"
    '��D
        SetgiMulti gmsStrtCode, "CodeNo", "Description", "CD017", "�N�X", "�W��", , True
        
End Sub

Private Sub subGil()
    On Error Resume Next
        SetgiList gilCitemCode, "CodeNo", "Description", "CD019", , , , , 3, 12, " Where PeriodFlag=1", True
         gilCitemCode.Clear
        SetgiList gilOldCMCode, "CodeNo", "Description", "CD031", , , , , 3, 12, , True
        gilOldCMCode.Clear
        SetgiList gilNewCMCode, "CodeNo", "Description", "CD031", , , , , 3, 12, , True
        gilNewCMCode.Clear
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
    
End Sub

Private Function IsDataOk() As Boolean
'2.  ���n���: �����~��d��, �վ㶵��, ��/�s�O��, ��/�s����
'(��:��O��=�s�O��, �B�����=�s����, �h�����L)
    On Error GoTo ChkErr
    IsDataOk = False
    Dim objText As Object
        If Not ChkDTok Then Exit Function
        If Not MustExist(gilCompCode, 2, "���q�O") Then Exit Function
        If Not MustExist(gdaClctYM1, 1, "�����~��_�l��") Then Exit Function
        If Not MustExist(gdaClctYM2, 1, "�����~��I���") Then Exit Function
        
        If gdaClctYM1.Text > gdaClctYM2.Text Then Call MsgDateRangeX("�����~��"): gdaClctYM1.SetFocus: Exit Function
        'If Not MustExist(gilCitemCode, 2, "���O����") Then Exit Function
        If Not MustExist(gilOldCMCode, 2, "�즬�O�覡") Then Exit Function
        
        IsDataOk = True
    Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub gdaClctYM1_Validate(Cancel As Boolean)
    If gdaClctYM1.Text <> "" Then
        Call gdaClctYM2.SetValue(gdaClctYM1.GetValue)
    End If
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        
        If Not ChgComp(gilCompCode, "SO3250", "SO3254") Then Exit Sub
        subGil
        subGim
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.SetCodeNo ""
        gilServiceType.Query_Description
        gilServiceType.ListIndex = 1
        Call GiMultiFilter(gmsServArea, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmsClctArea, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmsServEmp, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmsMduId, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gmsStrtCode, , gilCompCode.GetCodeNo)
       
        

End Sub

Private Sub gilServiceType_Change()
    On Error Resume Next
        Call GiListFilter(gilCitemCode, gilServiceType.GetCodeNo)
        gilCitemCode.Filter = gilCitemCode.Filter & IIf(Trim(gilCitemCode.Filter) = "", " Where ", " And ") & " PeriodFlag = 1"
        Call GiListFilter(gilOldCMCode, gilServiceType.GetCodeNo)
        Call GiListFilter(gilNewCMCode, gilServiceType.GetCodeNo)
        
End Sub

Private Sub mskprtSNo1_GotFocus()
    mskPrtSNo1.SelStart = 0
    mskPrtSNo1.SelLength = Len(mskPrtSNo1.Text)
End Sub

Private Sub mskprtSNo2_GotFocus()
    mskPrtSNo2 = mskPrtSNo1
    mskPrtSNo2.SelStart = 0
    mskPrtSNo2.SelLength = Len(mskPrtSNo2.Text)
End Sub

Private Function GetParaStr() As String
    On Error GoTo ChkErr
    Dim strSQL As String
        If mskPrtSNo1.Text <> "" And mskPrtSNo2.Text <> "" Then
            strSQL = "PrtSNo >='" & mskPrtSNo1.Text & "'  and  prtsno  <='" & mskPrtSNo2.Text & "'"
        ElseIf mskPrtSNo1.Text <> "" And mskPrtSNo2.Text = "" Then
                strSQL = "PrtSNo ='" & mskPrtSNo1.Text & "'"
            ElseIf mskPrtSNo1.Text = "" And mskPrtSNo2.Text <> "" Then
                strSQL = "PrtSNo='" & mskPrtSNo2.Text & "' "
        End If
            
        strSQL = strSQL & IIf(strSQL <> "", " And", "") & " ( ClctYM >=" & gdaClctYM1.Text & " And ClctYM <=" & gdaClctYM2.Text & " )"
        
        If gilCompCode.GetCodeNo <> "" Then
            strSQL = strSQL & " And  CompCode = " & gilCompCode.GetCodeNo
        End If
        
        If gmsServArea.GetQryStr <> "" Then
            strSQL = strSQL & " And  ServCode " & gmsServArea.GetQryStr
        End If
        
        If gmsClctArea.GetQryStr <> "" Then
            strSQL = strSQL & " And ClctAreaCode " & gmsClctArea.GetQryStr
        End If
    '
    '    If gimCMCode.GetQryStr <> "" Then
    '        strSQL = strSQL & " And CMCode " & gimCMCode.GetQryStr
    '    End If
        
        If gmsServEmp.GetQryStr <> "" Then
            strSQL = strSQL & " And ClctEn " & gmsServEmp.GetQryStr
        End If
        
        If gmsMduId.GetQryStr <> "" Then
            strSQL = strSQL & "  And MduId " & gmsMduId.GetQryStr
        End If
        
        If gmsStrtCode.GetQryStr <> "" Then
            strSQL = strSQL & " And StrtCode " & gmsStrtCode.GetQryStr
        End If
        
        If gilCitemCode.GetCodeNo <> "" Then
            strSQL = strSQL & " And CitemCode=" & gilCitemCode.GetCodeNo
        End If
        
        If gilOldCMCode.GetCodeNo <> "" Then
            strSQL = strSQL & " And CMCode=" & gilOldCMCode.GetCodeNo
        End If
    
    GetParaStr = strSQL
Exit Function
ChkErr:
    Call ErrSub(Me.Name, "GetParaStr")
End Function
