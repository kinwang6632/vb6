VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO5530A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�Ȥ�򥻸�Ƥ@���� [SO5530A]"
   ClientHeight    =   8010
   ClientLeft      =   1890
   ClientTop       =   435
   ClientWidth     =   8265
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO5530A.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '���ݵ�������
   Begin VB.ComboBox CmbNoe 
      Height          =   315
      ItemData        =   "SO5530A.frx":0442
      Left            =   1170
      List            =   "SO5530A.frx":044F
      TabIndex        =   28
      Top             =   6405
      Width           =   1335
   End
   Begin VB.CheckBox ChkDoor 
      Caption         =   "�O�_����"
      Height          =   300
      Left            =   2610
      TabIndex        =   29
      Top             =   6435
      Width           =   1110
   End
   Begin VB.PictureBox pic2 
      Height          =   4980
      Left            =   15
      ScaleHeight     =   4920
      ScaleWidth      =   8115
      TabIndex        =   51
      Top             =   1275
      Width           =   8175
      Begin VB.VScrollBar vsl1 
         Height          =   4935
         LargeChange     =   100
         Left            =   7830
         Max             =   100
         SmallChange     =   100
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   0
         Width           =   285
      End
      Begin VB.Frame fraMulti 
         BorderStyle     =   0  '�S���ؽu
         Caption         =   "Frame1"
         Height          =   7185
         Left            =   15
         TabIndex        =   52
         Top             =   0
         Width           =   8220
         Begin Gi_Multi.GiMulti gimWipCode 
            Height          =   345
            Left            =   -45
            TabIndex        =   27
            Top             =   6765
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   609
            ButtonCaption   =   "��  �u  ��  �A"
         End
         Begin CS_Multi.CSmulti gimStrtCode 
            Height          =   405
            Left            =   -45
            TabIndex        =   23
            Top             =   5265
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   714
            ButtonCaption   =   "��  �D  �d  ��"
         End
         Begin Gi_Multi.GiMulti gimBuyCode 
            Height          =   375
            Left            =   -45
            TabIndex        =   19
            Top             =   3900
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   661
            ButtonCaption   =   "�R  ��  ��  ��"
         End
         Begin Gi_Multi.GiMulti gimFaciCode 
            Height          =   375
            Left            =   -45
            TabIndex        =   18
            Top             =   3540
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   661
            ButtonCaption   =   "�]  ��  ��   ��"
         End
         Begin Gi_Multi.GiMulti gimBankCode 
            Height          =   375
            Left            =   -45
            TabIndex        =   17
            Top             =   3180
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   661
            ButtonCaption   =   "��  �b  ��  ��"
         End
         Begin Gi_Multi.GiMulti gimStatusCode 
            Height          =   375
            Left            =   -45
            TabIndex        =   11
            Top             =   690
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   661
            ButtonCaption   =   "��  ��  ��  �A"
         End
         Begin Gi_Multi.GiMulti gimPWCode 
            Height          =   375
            Left            =   -45
            TabIndex        =   15
            Top             =   2460
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   661
            ButtonCaption   =   "�I  �O  �N  �@"
         End
         Begin Gi_Multi.GiMulti gimCMCode 
            Height          =   375
            Left            =   -45
            TabIndex        =   16
            Top             =   2820
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   661
            ButtonCaption   =   "��  �O  ��  ��"
         End
         Begin Gi_Multi.GiMulti gimAreaCode 
            Height          =   375
            Left            =   -45
            TabIndex        =   22
            Top             =   4920
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   661
            ButtonCaption   =   "��     �F     ��"
         End
         Begin Gi_Multi.GiMulti gimServCode 
            Height          =   375
            Left            =   -45
            TabIndex        =   21
            Top             =   4590
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   661
            ButtonCaption   =   "�A     ��     ��"
         End
         Begin Gi_Multi.GiMulti gimOrder 
            Height          =   405
            Left            =   -45
            TabIndex        =   25
            Top             =   6015
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   714
            ButtonCaption   =   "��  ��  ��  ��"
            DataType        =   2
            ColumnOrder     =   -1  'True
         End
         Begin Gi_Multi.GiMulti gimSales 
            Height          =   375
            Left            =   -45
            TabIndex        =   26
            Top             =   6390
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   661
            ButtonCaption   =   "�~     ��     ��"
         End
         Begin CS_Multi.CSmulti gimMduId 
            Height          =   375
            Left            =   -30
            TabIndex        =   24
            Top             =   5655
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   661
            ButtonCaption   =   "�j  ��  �s  ��"
            Enabled         =   0   'False
         End
         Begin Gi_Multi.GiMulti gimClctAreaCode 
            Height          =   375
            Left            =   -45
            TabIndex        =   20
            Top             =   4245
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   661
            ButtonCaption   =   "��     �O     ��"
         End
         Begin CS_Multi.CSmulti gimClassCode 
            Height          =   375
            Left            =   -45
            TabIndex        =   13
            Top             =   1395
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   661
            ButtonCaption   =   "�Ȥ����O(�@)"
         End
         Begin CS_Multi.CSmulti gimClassCode2 
            Height          =   375
            Left            =   -45
            TabIndex        =   14
            Top             =   1740
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   661
            ButtonCaption   =   "�Ȥ����O(�G)"
         End
         Begin Gi_Multi.GiMulti gimNodeNo 
            Height          =   375
            Left            =   -45
            TabIndex        =   10
            Top             =   340
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   661
            ButtonCaption   =   "NodeNo"
            OneColumn       =   -1  'True
         End
         Begin CS_Multi.CSmulti gimCitemCode 
            Height          =   315
            Left            =   -45
            TabIndex        =   9
            Top             =   0
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   556
            ButtonCaption   =   "��  �O  ��  ��"
         End
         Begin CS_Multi.CSmulti gimClassCode3 
            Height          =   375
            Left            =   -45
            TabIndex        =   54
            Top             =   2085
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   661
            ButtonCaption   =   "�Ȥ����O(�T)"
         End
         Begin Gi_Multi.GiMulti gimWipCode3 
            Height          =   375
            Left            =   -45
            TabIndex        =   12
            Top             =   1040
            Width           =   7785
            _ExtentX        =   13732
            _ExtentY        =   661
            ButtonCaption   =   "�D�������"
         End
      End
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "�צ�Excel"
      Height          =   525
      Left            =   3540
      TabIndex        =   41
      Top             =   7320
      Width           =   1245
   End
   Begin VB.CheckBox chkEMail 
      Caption         =   "�Ȥ� E-Mail ²��C�L"
      Height          =   435
      Left            =   2970
      TabIndex        =   5
      Top             =   690
      Width           =   1245
   End
   Begin MSMask.MaskEdBox mskCircuitNo 
      Height          =   315
      Left            =   5175
      TabIndex        =   8
      Top             =   915
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskPeriod 
      Height          =   315
      Left            =   2460
      TabIndex        =   4
      Top             =   750
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   2
      Mask            =   "99"
      PromptChar      =   "_"
   End
   Begin prjNumber.GiNumber ginAmount 
      Height          =   315
      Left            =   870
      TabIndex        =   3
      Top             =   750
      Width           =   1095
      _ExtentX        =   1931
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
      ForeColor       =   0
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "�W���έp���G"
      Height          =   525
      Left            =   1740
      TabIndex        =   40
      Top             =   7320
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "����(&X)"
      Height          =   525
      Left            =   6930
      TabIndex        =   42
      Top             =   7320
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:�C�L"
      Default         =   -1  'True
      Height          =   525
      Left            =   90
      TabIndex        =   39
      Top             =   7320
      Width           =   1245
   End
   Begin VB.Frame fraPage 
      Caption         =   "�����覡"
      ForeColor       =   &H00FF0000&
      Height          =   870
      Left            =   3765
      TabIndex        =   49
      Top             =   6375
      Width           =   4320
      Begin VB.OptionButton OptStrt 
         BackColor       =   &H80000004&
         Caption         =   "��D"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   37
         Top             =   570
         Width           =   885
      End
      Begin VB.OptionButton optClctAreaCode 
         Caption         =   "���O��"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2160
         TabIndex        =   35
         Top             =   270
         Width           =   915
      End
      Begin VB.OptionButton optMduId 
         BackColor       =   &H80000004&
         Caption         =   "�j�ӽs��"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3120
         TabIndex        =   36
         Top             =   270
         Width           =   1155
      End
      Begin VB.OptionButton optNothing 
         Caption         =   "�L"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1170
         TabIndex        =   38
         Top             =   570
         Value           =   -1  'True
         Width           =   585
      End
      Begin VB.OptionButton optAreaCode 
         Caption         =   "��F��"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1170
         TabIndex        =   34
         Top             =   270
         Width           =   915
      End
      Begin VB.OptionButton optServCode 
         Caption         =   "�A�Ȱ�"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.Frame fraCustStatus 
      Caption         =   "�έp��H"
      ForeColor       =   &H00C00000&
      Height          =   510
      Left            =   90
      TabIndex        =   44
      Top             =   6705
      Width           =   3555
      Begin VB.OptionButton optNormal 
         Caption         =   "�D�j�Ӥ�"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   2340
         TabIndex        =   32
         Top             =   240
         Width           =   1125
      End
      Begin VB.OptionButton optMdu 
         Caption         =   "�j�Ӥ�"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1215
         TabIndex        =   31
         Top             =   240
         Width           =   885
      End
      Begin VB.OptionButton optAll 
         Caption         =   "����"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   165
         TabIndex        =   30
         Top             =   240
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin Gi_Date.GiDate gdaDate2 
      Height          =   345
      Left            =   2940
      TabIndex        =   2
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   12583104
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
   Begin Gi_Date.GiDate gdaDate1 
      Height          =   345
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   12583104
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
   Begin VB.ComboBox cboDateType 
      ForeColor       =   &H00C000C0&
      Height          =   315
      ItemData        =   "SO5530A.frx":046B
      Left            =   60
      List            =   "SO5530A.frx":047E
      TabIndex        =   0
      Text            =   "�U�����O���"
      Top             =   150
      Width           =   1545
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   5190
      TabIndex        =   6
      Top             =   120
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
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   5190
      TabIndex        =   7
      Top             =   525
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
      F5Corresponding =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
      Caption         =   "�A�����O"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4320
      TabIndex        =   50
      Top             =   630
      Width           =   780
   End
   Begin VB.Label lblCircuitNo 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�����s��"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4335
      TabIndex        =   48
      Top             =   945
      Width           =   780
   End
   Begin VB.Label lblPeriod 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "����"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2040
      TabIndex        =   47
      Top             =   795
      Width           =   390
   End
   Begin VB.Label lblAmount 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "���O���B"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   30
      TabIndex        =   46
      Top             =   795
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�� �q �O"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4320
      TabIndex        =   45
      Top             =   180
      Width           =   675
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   2820
      TabIndex        =   43
      Top             =   210
      Width           =   105
   End
End
Attribute VB_Name = "frmSO5530A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table:SO001,SO002,SO003,SO014,(SO004)
Option Explicit
Dim blnExcel As Boolean
Dim strIndex As String
Dim strExlSort As String

Private Sub chkEMail_Click()
 On Error Resume Next
   cmdExcel.Enabled = IIf(chkEMail.Value = 1, False, True)
End Sub

Private Sub cmdExcel_Click()
  On Error GoTo ChkErr
    blnExcel = True
    cmdPrint_Click
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExcel_Click")
End Sub

Private Sub cmdExit_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
    Call PreviousRpt(GetPrinterName(5), RptName("SO5530"), "�Ȥ�򥻸�Ƥ@���� [SO5530A]")
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo ChkErr
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    Screen.MousePointer = vbHourglass
      cmdExit.SetFocus
      Me.Enabled = False
        ReadyGoPrint
        Call subChoose
        Call subPrint
      Me.Enabled = True
    Screen.MousePointer = vbDefault
  
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Sub subPrint()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
  Dim strSubQry(0) As String
    rsTmp.CursorLocation = adUseClient
    Set rsTmp = gcnGi.Execute("SELECT " & strIndex & " Count(*) as intCount FROM " & strChoose & " And RowNum=1")
      
      If rsTmp("intCount") = 0 Then
         MsgNoRcd
      Else
        If chkEMail.Value = 1 Then
          strSQL = "SELECT SO001.CUSTID,SO001.CUSTNAME,SO002.Email From " & strChoose & " Order by SO001.Custid"
          Call PrintRpt2(GetPrinterName(5), RptName("SO5530", "1"), , "�Ȥ�򥻸�Ƥ@���� [SO5530A]", strSQL, strChooseString, , True, , , , GiPaperLandscape)
        Else
          strGroupName = strGroupName & ";blnExcel=" & blnExcel
          strSubQry(0) = "Select SO014.NodeNo,Count(Distinct SO001.Custid) CountCust From " & strChoose & " Group By SO014.NodeNo"
          If CreateSubView2(strSubQry) = False Then Exit Sub
          '���D��2814 ����q�X�ɡA�Ƨǳ��O�̫Ȥ�s���A�S���̷өҿ諸����
          strSQL = "SELECT " & strIndex & " SO001.CUSTID,SO001.CUSTNAME,SO001.TEL1,SO001.TEL2,SO001.Tel2Ext,SO001.TEL3,SO001.INSTADDRESS,SO002.CUSTSTATUSNAME,SO001.CLASSNAME1," & _
                   "SO001.CLASSNAME2,SO002.INSTTIME,SO003.CMNAME,SO003.CITEMNAME,SO003.PERIOD,SO003.AMOUNT,SO003.STARTDATE,SO003.STOPDATE,SO003.CLCTDATE,SO014.AreaCode," & _
                   "So014.AreaName,SO001.ServCode,So001.ServArea,SO001.ViewLevel,SO002.InstCount,SO014.AddrSort,SO001.CustNote,SO001.MduId,SO014.MDUName,SO002.PRTime," & _
                   "SO001.ClctAreaCode,SO001.ClctAreaName,SO003.BankName ,Decode(SO001.CLASSCODE2,NULL,Null,SO002.CustId) AS CLASSCODE_CustCount " & _
                   "From " & strChoose & IIf(Trim(strExlSort) = "", " Order By SO001.CustId ", " Order By ") & strExlSort
          strSubQry(0) = "SELECT * From " & strSubViewName(0) & " V"
          
          If blnExcel Then
            Call toExcel(strSQL, strSubQry(0))
          Else
            Call PrintRpt2(GetPrinterName(5), RptName("SO5530"), , "�Ȥ�򥻸�Ƥ@���� [SO5530A]", strSQL, strChooseString, , True, , , strGroupName, GiPaperLandscape, , strSubQry(0))
          End If
        End If
      End If
    CloseRecordset rsTmp
    blnExcel = False
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subPrint")
End Sub

Private Sub toExcel(ByVal strSQL As String, ByVal strSubQry1 As String)
  On Error GoTo ChkErr
    Dim rsExcel As New ADODB.Recordset
    Dim rsSubExcel1 As New ADODB.Recordset
    '950510 #2401 ��excel�|�֤@��,�Ƶ���ƤӪ�
'    RptToTxt RptName("SO5530", "E"), , strsql, strChooseString, , , strGroupName, , , Environ("Temp") & "\SO5530", , , strSubQry1
'    If Not Get_RS_From_Txt(Environ("Temp"), "SO5530.txt", rsExcel) Then blnExcel = False: Exit Sub
    Call GetRS(rsSubExcel1, strSubQry1, gcnGi, adUseServer)
    '#5718 �W�[�Q�P���X By Kin 2010/10/11
    GetRS rsExcel, "Select SO001.CUSTID �Ȥ�s��,SO001.CUSTNAME �Ȥ�m�W,SO001.TEL1 �q��1,SO001.TEL2 �q��2,SO001.Tel2Ext ����,SO001.TEL3 �q��3," & _
                    "To_char(SO002.INSTTIME,'YYYY/MM/DD hh:mm:ss')  �˾��ɶ�,To_char(SO002.PRTime,'YYYY/MM/DD hh:mm:ss') ����ɶ�, " & _
                    "SO002.CUSTSTATUSNAME �Ȥ᪬�A,SO001.ViewLevel ��������,SO002.InstCount �x��,SO001.CLASSNAME1 �Ȥ����O1,SO001.CLASSNAME2 �Ȥ����O2," & _
                    "SO003.AMOUNT ���B,SO003.PERIOD ����,To_char(SO003.CLCTDATE,'YYYY/MM/DD') �����O���,To_char(SO003.STARTDATE,'YYYY/MM/DD') ���Ĵ����}�l��," & _
                    "To_char(SO003.STOPDATE,'YYYY/MM/DD') ���Ĵ���������, SO001.INSTADDRESS �˾��a�},SO003.CMNAME ���O�覡,SO003.CITEMNAME ���O����," & _
                    "SO003.BankName �Ȧ�O,SO014.MDUNAME �j�ӦW��,SO001.CustNote �Ƶ�, SO002.PinCode �Q�P���X " & _
                    " From " & strChoose & IIf(strExlSort = "", " ORDER BY SO001.CUSTID", " Order By ") & strExlSort, gcnGi, adUseServer
     '#7074 To Add a choice-condition parameter into excel print function by Kin 2015/09/09
    Call UseProperty(rsExcel, "�Ȥ�򥻸�Ƥ@����", "�Ĥ@��", , , rsSubExcel1, "NodeNo,�Ȥ��", _
                    , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , strChooseString)
    
    blnExcel = False
    CloseRecordset rsExcel
    CloseRecordset rsSubExcel1
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "toExcel")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: strErrFile = "���q�O": GoTo 66
    If gilServiceType.GetCodeNo = "" Then gilServiceType.SetFocus: strErrFile = "�A�����O": GoTo 66
    IsDataOk = True
  Exit Function
66:
  MsgMustBe (strErrFile)
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub subChoose()
  On Error GoTo ChkErr
    Dim str1 As String
    Dim strFrom As String
    Dim strCustStatus As String
    Dim strPage As String
    Dim strSO3 As String
    Dim strSO4 As String
    Dim strGetGroupName As String
    Dim arrOrder() As String
    Dim varOrder As Variant
    Dim intSort As Integer
    Dim strDateName As String
    
    strChoose = ""
    strChooseString = ""
    strFrom = " SO003,SO014,SO001 "
    strGroupName = ""
    strSO4 = ""
    strIndex = ""
    
  '���
    If gdaDate1.GetValue <> "" Then Call subAnd(subDateType(">= To_Date('" & gdaDate1.GetValue & "','YYYYMMDD')"))
    If gdaDate2.GetValue <> "" Then Call subAnd(subDateType("< To_Date('" & gdaDate2.GetValue & "','YYYYMMDD')+1"))
    If gdaDate1.GetValue <> "" Or gdaDate2.GetValue <> "" Then
        strDateName = subDateType("")
        strIndex = GetUseIndexStr(Left(strDateName, InStr(strDateName, ".") - 1), Mid(strDateName, InStr(strDateName, ".") + 1))
    End If
    
  'GiList
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("Instr(Nvl(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "'),'" & gilServiceType.GetCodeNo & "') > 0")
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO002.ServiceType='" & gilServiceType.GetCodeNo & "'")
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("instr(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "')>0")
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO001.CompCode=" & gilCompCode.GetCodeNo)
  
  'GiMulti
    If gimCitemCode.GetQryStr <> "" Then Call subAnd("SO003.CitemCode " & gimCitemCode.GetQryStr)
    If gimNodeNo.GetQryStr <> "" Then Call subAnd("SO014.NodeNo " & gimNodeNo.GetQryStr)
    If gimClassCode.GetQryStr <> "" Then Call subAnd("SO001.ClassCode1 " & gimClassCode.GetQryStr)
    If gimClassCode2.GetQryStr <> "" Then Call subAnd("SO001.ClassCode2 " & gimClassCode2.GetQryStr)
    If gimClassCode3.GetQryStr <> "" Then Call subAnd("SO001.ClassCode3 " & gimClassCode3.GetQryStr)
    If gimCMCode.GetQryStr <> "" Then Call subAnd("SO003.CMCode " & gimCMCode.GetQryStr)
    'If gimMduId.GetQryStr <> "" Then Call subAnd("SO001.MduId " & gimMduId.GetQryStr)
    If gimPWCode.GetQryStr <> "" Then Call subAnd("SO001.PWCode " & gimPWCode.GetQryStr)
    If gimClctAreaCode.GetQryStr <> "" Then Call subAnd("SO001.ClctAreaCode " & gimClctAreaCode.GetQryStr)
    If gimServCode.GetQryStr <> "" Then Call subAnd("SO001.ServCode " & gimServCode.GetQryStr)
    If gimStatusCode.GetQryStr <> "" Then Call subAnd("SO002.CustStatusCode " & gimStatusCode.GetQryStr)
    If gimWipCode3.GetQryStr <> "" Then Call subAnd("NOT SO002.WipCode3 " & gimWipCode3.GetQryStr)   '���D��2322 KC�� ���D2341 �אּ�D�������,�簣���������A (�[ NOT)
    If gimAreaCode.GetQryStr <> "" Then Call subAnd("SO014.AreaCode " & gimAreaCode.GetQryStr)
    If gimStrtCode.GetQueryCode <> "" Then
       If gimStrtCode.GetQryStr <> "" Then Call subAnd("SO014.StrtCode " & gimStrtCode.GetQryStr)
       If Left(CmbNoe.Text, 1) = 1 Or Left(CmbNoe.Text, 1) = 2 Then Call subAnd("(SO014.Noe1=" & Left(CmbNoe.Text, 1) & " or SO014.Noe2=" & Left(CmbNoe.Text, 1) & " or SO014.Noe3=" & Left(CmbNoe.Text, 1) & " or SO014.Noe4=" & Left(CmbNoe.Text, 1) & ")")
       If ChkDoor.Value = 1 Then Call subAnd("SO014.Door1=" & ChkDoor.Value)
        
    End If
    If gimBuyCode.GetQryStr <> "" Then strSO4 = "SO004.BuyCode " & gimBuyCode.GetQryStr
    If gimFaciCode.GetQryStr <> "" Then strSO4 = IIf(strSO4 = "", "", strSO4 & " And ") & "SO004.FaciCode " & gimFaciCode.GetQryStr
    If gimSales.GetQryStr <> "" Then Call subAnd("SO014.SalesCode " & gimSales.GetQryStr)
    If gimWipCode.GetQueryCode <> "" Then Call subAnd("(" & SplitWipCode(gimWipCode.GetQueryCode, "SO002") & ")")
    
    If gimStrtCode.GetQueryCode = "" Then
        CmbNoe.Text = ""
        ChkDoor.Value = 0
        CmbNoe.Enabled = False
        ChkDoor.Enabled = False
    Else
        CmbNoe.Enabled = True
        ChkDoor.Enabled = True
    End If
  '�Ȧ����O
    If gimBankCode.GetQryStr <> "" Then
        Call subAnd("SO003.BankCode " & gimBankCode.GetQryStr)
    Else
       If gimBankCode.GetDispStr <> "" Then Call subAnd("SO003.BankCode is not null")
    End If
    
    '���O����
'    If gilCitemCode.GetCodeNo <> "" Then Call subAnd("SO003.CitemCode=" & gilCitemCode.GetCodeNo)
    
  '���O���B
    If ginAmount.Text <> "" Then Call subAnd("SO003.Amount=" & ginAmount.Value)
  '����
    If mskPeriod.Text <> "" Then Call subAnd("SO003.Period=" & mskPeriod.Text)
  '�����s��
    If mskCircuitNo.Text <> "" Then Call subAnd("SO014.CircuitNo ='" & mskCircuitNo.Text & "'")
    
  '�έp��H
    Select Case True
        Case optMdu.Value
             If gimMduId.GetQryStr <> "" Then
                subAnd "SO001.MduId " & gimMduId.GetQryStr
             Else
                subAnd "SO001.MduId Is Not Null"
             End If
             strCustStatus = "�j�Ӥ�"
        Case optNormal.Value
             subAnd "SO001.MduId Is Null"
             strCustStatus = "�D�j�Ӥ�"
        Case Else
             If gimMduId.GetQryStr <> "" Then
                subAnd "(SO001.MduId Is Null or SO001.Mduid " & gimMduId.GetQryStr & ")"
             End If
             strCustStatus = "����"
    End Select
    If chkEMail.Value = 1 Then strChoose = strChoose & " And SO002.EMail is not Null "
    str1 = " So001.InstAddrNo= So014.AddrNo And SO001.CompCode = SO014.CompCode "
    
  '�P�_So002�O���nJoin
    'If InStr(1, UCase(strChoose), "SO002.") > 0 Or chkEMail.Value = 1 Then
        strFrom = strFrom & ",SO002"
        str1 = str1 & " And  SO001.CustId = SO002.CustId And SO001.CompCode = SO002.CompCode " '& IIf(gilServiceType.GetCodeNo <> "", " And So002.ServiceType = '" & gilServiceType.GetCodeNo & "' ", "")
    'End If
    
  '�P�_So002A��Join �覡
'    If InStr(1, UCase(strChoose), "SO002A.") > 0 Then
'        strFrom = strFrom & ",SO002A"
'        str1 = str1 & " And SO001.CustId = SO002A.CustId "
'        '�P�_So003��Join �覡
'        If InStr(1, UCase(strChoose), "SO003.") > 0 Then
'            str1 = str1 & " And SO002A.CustId = SO003.CustId And SO002A.AccountNo = SO003.AccountNo"
'        Else
'            str1 = str1 & " And SO002A.CustId = SO003.CustId(+) And SO002A.AccountNo = SO003.AccountNo(+) "
'        End If
'    Else
        '�P�_So003��Join �覡
'        If InStr(1, UCase(strChoose), "SO003.") > 0 Then
'            str1 = str1 & " And SO002.CustId = SO003.CustId(+) "
''            str1 = str1 & " And SO002.CustId = SO003.CustId(+) And SO002.ServiceType=SO003.ServiceType(+) "
'        Else
            '���D��1066 ���P����U93/08/01~93/08/03�A����]�X��11���A�������ܸ�Ʈw������ƫo��14���A�g�d�ҬO��Ȥ�S�g���ʸ�ƮɡA�|�Q�ư��b�~(SQL�����g��SO002.servicetype=so003.servicetype�o�y�|�N�LA������Ʊư�)
            'for Debby ---- Crystal Edit
'            str1 = str1 & " And SO001.CustId = SO003.CustId(+)  And SO002.ServiceType=SO003.ServiceType(+) "
            str1 = str1 & " And  SO003.CustId(+) = SO002.CustId AND SO003.ServiceType(+)=SO002.ServiceType AND SO003.CompCode(+)=SO002.CompCode "
'        End If
'    End If
    
  '�P�_�O�_�n�J�nTable:SO004
  '#4323 �|�s���O�ӪA�Ȫ���Ƴ��D��X��,�ҥH�n�W�[�e���W���A�ȧO By Kin 2009/01/12
    If strSO4 <> "" Then
       strFrom = strFrom & " ,(SELECT distinct CustId, CompCode From SO004 Where " & strSO4 & " And SO004.ServiceType='" & gilServiceType.GetCodeNo & "' ) SO004 "
       str1 = str1 & " And SO001.CustId=SO004.CustId And SO001.CompCode = SO004.CompCode "
    End If
    
    If strChoose = "" Then
       strChoose = strFrom & " Where " & str1
    Else
       strChoose = strFrom & " Where " & str1 & " And " & strChoose
    End If
      
  '�����覡
    Select Case True
           Case optAreaCode.Value
                strGroupName = "ReportType=True;GroupName={SO014.AreaCode};GroupName2={SO014.AreaName}"
                strPage = "��F��"
           Case optServCode.Value
                strGroupName = "ReportType=True;GroupName={SO001.ServCode};GroupName2={SO001.ServArea}"
                strPage = "�A�Ȱ�"
           Case optClctAreaCode.Value
                strGroupName = "ReportType=True;GroupName={SO001.ClctAreaCode};GroupName2={SO001.ClctAreaName}"
                strPage = "���O��"
           Case optMduId.Value
                strGroupName = "ReportType=True;GroupName={SO001.MduId};GroupName2={SO014.MDUName}"
                strPage = "�j�ӽs��"
           Case OptStrt.Value
                strGroupName = "ReportType=True;GroupName={SO014.StrtCode};GroupName2={SO014.StrtName}"
                strPage = "��D"
           Case optNothing.Value
            '���D��2814 �N�ǤJ���ƧǮ���,�אּSQL�Ƨ�
                Select Case Left(gimOrder.GetColumnOrderCode, 1)
                       Case "A"
                            strGroupName = "ReportType=False;GroupName={SO001.Custid};GroupName2={SO001.Custid}" ';Sort" & intSort & "={SO001.Custid}"
                       Case "B"
                            strGroupName = "ReportType=False;GroupName={SO014.AddrSort};GroupName2={SO014.AddrSort}" ';Sort" & intSort & "={SO014.AddrSort}"
                       Case "C"
                            strGroupName = "ReportType=False;GroupName={SO003.CitemCode};GroupName2={SO003.CitemCode}" ';Sort" & intSort & "={SO003.CitemCode}"
                       Case "D"
                            strGroupName = "ReportType=False;GroupName={SO001.ServCode};GroupName2={SO001.ServCode}" ';Sort" & intSort & "={SO001.ServCode}"
                       Case "E"
                            strGroupName = "ReportType=False;GroupName={SO001.ClassCode1};GroupName2={SO001.ClassCode1}" ';Sort" & intSort & "={SO001.ClassCode1}"
                       Case Else
                            strGroupName = "ReportType=False;GroupName={SO001.Custid};GroupName2={SO001.Custid}" ';Sort" & intSort & "={SO001.Custid}"
                End Select
                strPage = "�L"
    End Select
    
  '�ƧǤ覡
    strExlSort = ""
    If gimOrder.GetColumnOrderCode <> "" Then
      arrOrder = Split(gimOrder.GetColumnOrderCode, ",")
      intSort = 0
      
      For Each varOrder In arrOrder
      '���D��2814 �N�ǤJ���ƧǮ���,�אּSQL�Ƨ�
        Select Case arrOrder(intSort)
               Case "A"
                    'strGroupName = strGroupName & ";Sort" & intSort & "={SO001.Custid}"
                    strExlSort = strExlSort & IIf(strExlSort = "", "", ",") & "SO001.Custid"
               Case "B"
                    'strGroupName = strGroupName & ";Sort" & intSort & "={SO014.AddrSort}"
                    strExlSort = strExlSort & IIf(strExlSort = "", "", ",") & "SO014.AddrSort"
               Case "C"
                    'strGroupName = strGroupName & ";Sort" & intSort & "={SO003.CitemCode}"
                    strExlSort = strExlSort & IIf(strExlSort = "", "", ",") & "SO003.CitemCode"
               Case "D"
                    'strGroupName = strGroupName & ";Sort" & intSort & "={SO001.ServCode}"
                    strExlSort = strExlSort & IIf(strExlSort = "", "", ",") & "SO001.ServCode"
               Case "E"
                    'strGroupName = strGroupName & ";Sort" & intSort & "={SO001.ClassCode1}"
                    strExlSort = strExlSort & IIf(strExlSort = "", "", ",") & "SO001.ClassCode1"
        End Select
        intSort = intSort + 1
      Next
    End If
       
    strChooseString = cboDateType.Text & ":" & subSpace(gdaDate1.GetValue(True)) & "~" & subSpace(gdaDate2.GetValue(True)) & ";" & _
                      "���q�O�@:" & subSpace(gilCompCode.GetDescription) & ";" & _
                      "�A�����O:" & subSpace(gilServiceType.GetDescription) & ";" & _
                      "���O����:" & subSpace(gimCitemCode.GetDispStr) & ";" & _
                      "���@�@�B:" & subSpace(ginAmount.Text) & ";" & "���@�@��:" & subSpace(mskPeriod.Text) & ";" & _
                      "�����s��:" & subSpace(mskCircuitNo.Text) & ";" & _
                      "NodeNo  :" & subSpace(gimNodeNo.GetDispStr) & ";" & _
                      "�Ȥ᪬�A:" & subSpace(gimStatusCode.GetDispStr) & ";����������A:" & subSpace(gimWipCode3.GetDispStr) & ";" & _
                      "�Ȥ����O(�@):" & subSpace(gimClassCode.GetDispStr) & ";" & _
                      "�Ȥ����O(�G):" & subSpace(gimClassCode2.GetDispStr) & ";" & _
                      "�Ȥ����O(�T):" & subSpace(gimClassCode3.GetDispStr) & ";" & _
                      "�I�O�N�@:" & subSpace(gimPWCode.GetDispStr) & ";" & _
                      "���O�覡:" & subSpace(gimCMCode.GetDispStr) & ";" & _
                      "��b�Ȧ�:" & subSpace(gimBankCode.GetDispStr) & ";" & _
                      "�]�ƶ���:" & subSpace(gimFaciCode.GetDispStr) & ";" & _
                      "�R��覡:" & subSpace(gimBuyCode.GetDispStr) & ";" & _
                      "���O��  :" & subSpace(gimClctAreaCode.GetDispStr) & ";" & _
                      "�A�Ȱϡ@:" & subSpace(gimServCode.GetDispStr) & ";" & _
                      "��F�ϡ@:" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                      "��D�d��:" & subSpace(gimStrtCode.GetDispStr) & ";" & "������:" & subSpace(CmbNoe.Text) & ";" & "�O�_����:" & subSpace(IIf(ChkDoor.Value = 1, "�O", "�_")) & ";" & _
                      "�j�ӦW��:" & subSpace(gimMduId.GetDispStr) & ";" & _
                      "�~�Ȱ�  :" & subSpace(gimSales.GetDispStr) & ";" & _
                      "���u���A:" & subSpace(gimWipCode.GetDispStr) & ";" & _
                      "�έp��H:" & subSpace(strCustStatus) & ";" & _
                      "�ƧǤ覡:" & subSpace(gimOrder.GetColumnOrderDspStr) & ";"
    strChooseString = strChooseString & "�����覡:" & subSpace(strPage)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Private Function subAndStr(strBody As String, strAnd As String) As String
  On Error GoTo ChkErr
    If strBody = "" Then
       subAndStr = strAnd
    Else
       subAndStr = strBody & " And " & strAnd
    End If
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Function

Private Function subDateType(strGetValue As String) As String
  On Error GoTo ChkErr
    Select Case cboDateType.ListIndex
           Case 1
                subDateType = "SO002.InstTime " & strGetValue
           Case 2
                subDateType = "SO002.PRTime " & strGetValue
           Case 3
                subDateType = "SO002.StopTime " & strGetValue
           Case 4
                subDateType = "SO002.DelDate " & strGetValue
           Case Else
                subDateType = "SO003.ClctDate " & strGetValue
    End Select
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subDateType")
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5530A)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGim
    Call subGil
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description

  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "���O���إN�X", "���O���ئW��")
    Call SetgiMulti(gimNodeNo, "CompCode", "CodeNo", "CD047", "CompCode", "NodeNo")
    Call SetgiMulti(gimStatusCode, "CodeNo", "Description", "CD035", "�Ȥ᪬�A�N�X", "�Ȥ᪬�A�W��")
    Call SetgiMulti(gimWipCode3, "CodeNo", "Description", "CD036", "���u���A�N�X", "���u���A�W��")      '���D��2322
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "�Ȥ����O�N�X�@", "�Ȥ����O�W�٤@")
    Call SetgiMulti(gimClassCode2, "CodeNo", "Description", "CD004", "�Ȥ����O�N�X�G", "�Ȥ����O�W�٤G")
    Call SetgiMulti(gimClassCode3, "CodeNo", "Description", "CD004", "�Ȥ����O�N�X�T", "�Ȥ����O�W�٤T")
    Call SetgiMulti(gimPWCode, "CodeNo", "Description", "CD020", "�I�O�N�@�N�X", "�I�O�N�@�W��")
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "���O�覡�N�X", "���O�覡�W��")
    Call SetgiMulti(gimBankCode, "CodeNo", "Description", "CD018", "�Ȧ�N�X", "�Ȧ�W��")
    Call SetgiMulti(gimFaciCode, "CodeNo", "Description", "CD022", "�]�ƶ��إN�X", "�]�ƶ��ئW��")
    Call SetgiMulti(gimBuyCode, "CodeNo", "Description", "CD034", "�R��覡�N�X", "�R��覡�W��")
    Call SetgiMulti(gimClctAreaCode, "CodeNo", "Description", "CD040", "���O�ϥN�X", "���O�ϦW��")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "�A�ȰϥN�X", "�A�ȰϦW��")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "��F�ϥN�X", "��F�ϦW��")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "��D�N�X", "��D�W��")
    Call SetgiMulti(gimMduId, "MduId", "Name", "SO017", "�j�ӥN��", "�j�ӦW��")
    Call SetgiMultiAddItem(gimOrder, "A,B,C,D,E", "�Ȥ�s��,�˾��a�},���O����,�A�Ȱ�,�Ȥ����O")
    Call SetgiMulti(gimSales, "CodeNo", "Description", "CD050", "�~�ȰϥN�X", "�~�ȰϦW��")
    Call SetgiMulti(gimWipCode, "CodeNo", "Description", "CD036", "���u���A�N�X", "���u���A�W��", "Where CodeNo<>0")
    gimStatusCode.SetQueryCode "1,2"
    gimWipCode3.SetQueryCode "13"
    If gimStrtCode.GetQryStr = "" Then
        CmbNoe.Enabled = False
        ChkDoor.Enabled = False
    Else
        CmbNoe.Enabled = True
        ChkDoor.Enabled = True
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    'Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
'    Call SetgiList(gilCitemCode, "CodeNo", "Description", "CD019")
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5530A
End Sub

Private Sub gdaDate1_GotFocus()
  On Error Resume Next
  If gdaDate1.GetValue = "" Then gdaDate1.SetValue (RightDate)
End Sub

Private Sub gdaDate2_GotFocus()
  On Error Resume Next
  If gdaDate1.GetValue = "" Or gdaDate2.GetValue = "" Then gdaDate2.SetValue (gdaDate1.GetValue)
End Sub

Private Sub gdaDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaDate1, gdaDate2)
End Sub

Private Sub gilCompCode_Change()
On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    GiMultiFilter gimNodeNo, gilServiceType.GetCodeNo, gilCompCode.GetCodeNo
    GiMultiFilter gimClctAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimBankCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimMduId, , gilCompCode.GetCodeNo
    GiMultiFilter gimSales, , gilCompCode.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Change")
End Sub

Private Sub gilCompCode_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then
        MsgMustBe ("���q�O")
        Cancel = True
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Sub gilServiceType_Change()
  On Error GoTo ChkErr
    GiMultiFilter gimCitemCode, gilServiceType.GetCodeNo
    GiMultiFilter gimNodeNo, gilServiceType.GetCodeNo, gilCompCode.GetCodeNo
    GiMultiFilter gimBuyCode, gilServiceType.GetCodeNo
    GiMultiFilter gimClassCode, gilServiceType.GetCodeNo
    GiMultiFilter gimClassCode2, gilServiceType.GetCodeNo
    GiMultiFilter gimClassCode3, gilServiceType.GetCodeNo
    GiMultiFilter gimCMCode, gilServiceType.GetCodeNo
    GiMultiFilter gimFaciCode, gilServiceType.GetCodeNo
    GiMultiFilter gimPWCode, gilServiceType.GetCodeNo

'    Call GiListFilter(gilCitemCode, gilServiceType.GetCodeNo)
'    gimCitemCode.Filter = gimCitemCode.Filter & IIf(gimCitemCode.Filter = "", " WHERE ", " AND ") & " PeriodFlag = 1"
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilServiceType_Change")
End Sub

Private Sub gimNodeNo_GotFocus()
    On Error Resume Next
        vsl1.Value = 0
End Sub

Private Sub gimStrtCode_Change()
    CmbNoe.Text = "": ChkDoor.Value = 0
End Sub

Private Sub gimStrtCode_GotFocus()
        CmbNoe.Enabled = True
        ChkDoor.Enabled = True
End Sub

Private Sub gimWipCode_GotFocus()
  On Error Resume Next
    vsl1.Value = 100
End Sub

Private Sub optAll_Click()
  On Error Resume Next
    gimMduId.Clear
    gimMduId.Enabled = False
End Sub

Private Sub optMdu_Click()
  On Error Resume Next
    gimMduId.Enabled = True
End Sub

Private Sub optNormal_Click()
  On Error Resume Next
    gimMduId.Clear
    gimMduId.Enabled = False
End Sub

Private Sub vsl1_Change()
  On Error Resume Next
    If vsl1.Value = 0 Then
        fraMulti.Top = 20
    ElseIf vsl1.Value = 100 Then
        fraMulti.Top = -(pic2.Height) + 480
    End If
End Sub
