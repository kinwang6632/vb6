VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.17#0"; "csMulti.ocx"
Begin VB.Form frmSO5410A 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�Ȥ�]�ƥ[�ˤ@���� [SO5410A]"
   ClientHeight    =   8685
   ClientLeft      =   1230
   ClientTop       =   525
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO5410A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form22"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '���ݵ�������
   Begin VB.TextBox txtCustId 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2880
      TabIndex        =   55
      Top             =   6360
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   3705
      TabIndex        =   53
      Top             =   6795
      Width           =   1800
      Begin VB.OptionButton OptI 
         Caption         =   "��]�ƥ[��"
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   165
         TabIndex        =   54
         Top             =   165
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.OptionButton optM 
         Caption         =   "���׬���"
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   165
         TabIndex        =   34
         Top             =   450
         Width           =   1095
      End
      Begin VB.OptionButton optP 
         Caption         =   "�������"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   165
         TabIndex        =   35
         Top             =   705
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "�צ�Excel"
      Height          =   525
      Left            =   3840
      TabIndex        =   40
      Top             =   8055
      Width           =   1245
   End
   Begin VB.CheckBox chkFac 
      Caption         =   "��ܳ]�ƸԲӸ��"
      Height          =   345
      Left            =   5640
      TabIndex        =   36
      Top             =   6915
      Width           =   1905
   End
   Begin VB.CheckBox chkContractCust 
      Caption         =   "�j����"
      Height          =   345
      Left            =   5655
      TabIndex        =   37
      Top             =   7260
      Width           =   1065
   End
   Begin VB.Frame fraSort 
      Caption         =   "�����覡"
      ForeColor       =   &H00FF0000&
      Height          =   1050
      Left            =   75
      TabIndex        =   44
      Top             =   6795
      Width           =   3570
      Begin VB.OptionButton OptIntroId 
         Caption         =   "���ФH"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2520
         TabIndex        =   33
         Top             =   600
         Width           =   945
      End
      Begin VB.OptionButton optClctAreaCode 
         Caption         =   "���O��"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   645
         Width           =   915
      End
      Begin VB.OptionButton OptServCode 
         Caption         =   "�A�Ȱ�"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   135
         TabIndex        =   28
         Top             =   270
         Width           =   885
      End
      Begin VB.OptionButton OptAreaCode 
         Caption         =   "��F��"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   1170
         TabIndex        =   29
         Top             =   255
         Width           =   945
      End
      Begin VB.OptionButton OptInstEn1 
         Caption         =   "�w�ˤH���@"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1185
         TabIndex        =   32
         Top             =   630
         Width           =   1305
      End
      Begin VB.OptionButton optNothing 
         Caption         =   "�L"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2520
         TabIndex        =   30
         Top             =   240
         Value           =   -1  'True
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdPrevRpt 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�W���έp���G"
      Height          =   525
      Left            =   2115
      TabIndex        =   39
      Top             =   8055
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   525
      Left            =   6240
      TabIndex        =   41
      Top             =   8010
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F5:�C�L"
      Height          =   525
      Left            =   180
      TabIndex        =   38
      Top             =   8040
      Width           =   1245
   End
   Begin VB.Frame fraType 
      Caption         =   "����榡"
      ForeColor       =   &H00004080&
      Height          =   705
      Left            =   4275
      TabIndex        =   43
      Top             =   1170
      Width           =   3225
      Begin VB.OptionButton optSummaries 
         Caption         =   "�J�`��"
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   885
      End
      Begin VB.OptionButton optDail 
         Caption         =   "���Ӫ�"
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   420
         TabIndex        =   12
         Top             =   300
         Value           =   -1  'True
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�]�ƨϥΪ��p"
      ForeColor       =   &H00004080&
      Height          =   975
      Left            =   4290
      TabIndex        =   42
      Top             =   105
      Width           =   3225
      Begin VB.OptionButton optPRing 
         Caption         =   "���"
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   1035
      End
      Begin VB.OptionButton optUse 
         Caption         =   "�ϥΤ�"
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   420
         TabIndex        =   8
         Top             =   240
         Width           =   1035
      End
      Begin VB.OptionButton optPR 
         Caption         =   "�v�"
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   420
         TabIndex        =   10
         Top             =   600
         Width           =   1035
      End
      Begin VB.OptionButton optAll 
         Caption         =   "����"
         ForeColor       =   &H00004080&
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Top             =   600
         Value           =   -1  'True
         Width           =   1005
      End
   End
   Begin CS_Multi.CSmulti gimStrtRange 
      Height          =   345
      Left            =   0
      TabIndex        =   21
      Top             =   4320
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   609
      ButtonCaption   =   "��  �D  �d  ��"
   End
   Begin prjNumber.GiNumber ginAmount1 
      Height          =   315
      Left            =   1230
      TabIndex        =   4
      Top             =   840
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
      MaxLength       =   10
   End
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   345
      Left            =   0
      TabIndex        =   20
      Top             =   3990
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   609
      ButtonCaption   =   "��     �F     ��"
   End
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   345
      Left            =   0
      TabIndex        =   19
      Top             =   3645
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   609
      ButtonCaption   =   "�A     ��     ��"
   End
   Begin Gi_Multi.GiMulti gimCustStatus 
      Height          =   345
      Left            =   0
      TabIndex        =   18
      Top             =   3300
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   609
      ButtonCaption   =   "��  ��  ��  �A"
   End
   Begin Gi_Multi.GiMulti gimBuy 
      Height          =   345
      Left            =   0
      TabIndex        =   16
      Top             =   2640
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   609
      ButtonCaption   =   "�R  ��  ��  ��"
   End
   Begin Gi_Multi.GiMulti gimFaciCode 
      Height          =   345
      Left            =   0
      TabIndex        =   15
      Top             =   2310
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   609
      ButtonCaption   =   "�]  ��  ��  ��"
   End
   Begin Gi_Date.GiDate gdaInstDate2 
      Height          =   375
      Left            =   2580
      TabIndex        =   1
      Top             =   60
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ForeColor       =   12582912
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
   Begin Gi_Date.GiDate gdaInstDate1 
      Height          =   375
      Left            =   1230
      TabIndex        =   0
      Top             =   60
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ForeColor       =   12582912
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
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1230
      TabIndex        =   6
      Top             =   1200
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
      Left            =   1230
      TabIndex        =   7
      Top             =   1590
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
   Begin Gi_Multi.GiMulti gimOrder 
      Height          =   345
      Left            =   0
      TabIndex        =   23
      Top             =   4980
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   609
      ButtonCaption   =   "��  ��  ��  ��"
      DataType        =   2
      ColumnOrder     =   -1  'True
   End
   Begin prjNumber.GiNumber ginAmount2 
      Height          =   315
      Left            =   2580
      TabIndex        =   5
      Top             =   840
      Width           =   1125
      _ExtentX        =   1984
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
      MaxLength       =   10
   End
   Begin CS_Multi.CSmulti gimInstEn 
      Height          =   345
      Left            =   0
      TabIndex        =   22
      Top             =   4650
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   609
      ButtonCaption   =   "�w  ��  �H  ��"
   End
   Begin CS_Multi.CSmulti gimClassCode 
      Height          =   345
      Left            =   0
      TabIndex        =   17
      Top             =   2970
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   609
      ButtonCaption   =   "��  ��  ��  �O"
   End
   Begin Gi_Multi.GiMulti gimBillType 
      Height          =   315
      Left            =   0
      TabIndex        =   14
      Top             =   1980
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   556
      ButtonCaption   =   "��  ��  ��  �O"
      DataType        =   2
      DIY             =   -1  'True
      Exception       =   -1  'True
   End
   Begin Gi_Multi.GiMulti gimNodeNo 
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   5310
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   661
      ButtonCaption   =   "NodeNo"
      OneColumn       =   -1  'True
   End
   Begin CS_Multi.CSmulti gimMduId 
      Height          =   375
      Left            =   15
      TabIndex        =   25
      Top             =   5655
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   661
      ButtonCaption   =   "�j  ��  �s  ��"
   End
   Begin Gi_Date.GiDate gdaPRDate2 
      Height          =   345
      Left            =   2580
      TabIndex        =   3
      Top             =   465
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   12582912
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
   Begin Gi_Date.GiDate gdaPRDate1 
      Height          =   345
      Left            =   1230
      TabIndex        =   2
      Top             =   465
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   12582912
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
   Begin Gi_Multi.GiMulti gimMediaCode 
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   6000
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   661
      ButtonCaption   =   "��  ��  �C  ��"
   End
   Begin CS_Multi.CSmulti gimIntroId 
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   6345
      Visible         =   0   'False
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   661
      ButtonCaption   =   "��      ��     �H"
   End
   Begin VB.Label lblCustId 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�Ȥ�s��(�H,���j �H-���϶�)"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   56
      Top             =   6460
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "---"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2370
      TabIndex        =   52
      Top             =   525
      Width           =   180
   End
   Begin VB.Label lblPRDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�]�Ʃ���"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   15
      TabIndex        =   51
      Top             =   525
      Width           =   1170
   End
   Begin VB.Label lblInstDay 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�]�Ʀw�ˤ��"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   30
      TabIndex        =   50
      Top             =   150
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "---"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2370
      TabIndex        =   49
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
      Caption         =   "�A�����O"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   330
      TabIndex        =   48
      Top             =   1620
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
      Caption         =   "��  �q  �O"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   360
      TabIndex        =   47
      Top             =   1260
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�`���B"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   555
      TabIndex        =   46
      Top             =   930
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "---"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2370
      TabIndex        =   45
      Top             =   900
      Width           =   180
   End
End
Attribute VB_Name = "frmSO5410A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table:SO001,SO002,SO004,SO014
Option Explicit
Dim strGroup As String
Dim strWhere As String
Dim strFilter As String
Dim strViewName1 As String
Dim blnSO002 As Boolean
Dim strField As String
Dim blnExcel As Boolean

'���D2653 �H���Ӫ��P�_�з�2006/07/24 Crystal Edit
'Private Sub chkFac_Click()
'    If chkFac.Value = 1 Then
'        cmdExcel.Enabled = True
'    Else
'        cmdExcel.Enabled = False
'    End If
'End Sub

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
    ErrSub Me.Name, "cmdExit_Click"
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    PreviousRpt GetPrinterName(5), RptName("SO5410"), "�Ȥ�]�ƥ[�ˤ@���� [SO5410A]"
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo ChkErr
    Dim rs As New ADODB.Recordset
    Dim strQry As String
    Dim strSubRpt(3) As String
    Dim intFor As Integer
'    Dim strSubRpt1 As String
'    Dim strSubRpt2 As String
    'Dim strCode As String
      If Not ChkDTok Then Exit Sub
      If Not IsDataOk Then Exit Sub
        Screen.MousePointer = vbHourglass
          Me.Enabled = False
            ReadyGoPrint
            If Not subChoose Then Exit Sub
            strSubRpt(0) = "Select SO004.FaciName,SO004.BuyName,Count(*) intCount " & strChoose & " Group By SO004.FaciName,SO004.BuyName"
            strSubRpt(1) = "Select SO014.AreaCode,SO014.AreaName,SO004.CMBaudRate,Count(Distinct SO004.Custid) CountCust,sum(SO004.Quantity) SumQuantity " & strChoose & " And SO004.FaciCode in (Select CodeNo From CD022 Where RefNo=2)" & " Group By SO014.AreaCode,SO014.AreaName,SO004.CMBaudRate"
            strSubRpt(3) = "Select SO004.IntroId,SO004.IntroName,Count(*) intCount " & strChoose & " Group By SO004.IntroId,SO004.IntroName"
            If Not CreateSubView2(strSubRpt()) Then Exit Sub
            If Not CreateSubView3(strChoose) Then Exit Sub
            If Not GetRS(rs, "Select Count(*) CountValue " & strChoose & " And RowNum=1", gcnGi, adUseClient) Then Exit Sub
                If rs!CountValue = 0 Then
                   MsgNoRcd
                   SendSQL , , True
                Else
                   For intFor = 0 To UBound(strSubRpt)
                      strSubRpt(intFor) = "Select * From " & strSubViewName(intFor) & " V"
                   Next
                   strSubRpt(2) = "Select * From " & strViewName & " V"
                   strQry = strField & strChoose
'                   If optSummaries.Value Then strQry = strQry & " AND 1=0"
                    If blnExcel Then
                        Call toExcel(strQry)
                    Else
                        PrintRpt GetPrinterName(5), RptName("SO5410"), , "�Ȥ�]�ƥ[�ˤ@���� [SO5410A]", strQry, strChooseString, , True, , , strGroup, GiPaperLandscape, GiPaper_A4, strSubRpt(0), strSubRpt(1), strSubRpt(2), strSubRpt(3)
                    End If
                End If
            CloseRecordset rs
          Me.Enabled = True
        Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Sub toExcel(ByVal strsql As String)
  On Error GoTo ChkErr
    Dim rsExcel As New ADODB.Recordset
'    PrintRpt GetPrinterName(5), RptName("SO5410"), , "�Ȥ�]�ƥ[�ˤ@���� [SO5410A]", strQry, strChooseString, , True, , , strGroup, GiPaperLandscape, GiPaper_A4, strSubRpt(0), strSubRpt(1), strSubRpt(2)
    RptToTxt RptName("SO5410", "E"), , strsql, strChooseString, , , strGroupName, , , Environ("Temp") & "\SO5410"
    If Not Get_RS_From_Txt(Environ("Temp"), "SO5410.txt", rsExcel) Then blnExcel = False: Exit Sub
    Call UseProperty(rsExcel, "�Ȥ�]�ƥ[�ˤ@����", "�Ĥ@��")
    blnExcel = False
    CloseRecordset rsExcel
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

Private Function GetSort() As String
  On Error GoTo ChkErr
    Select Case True
           Case optServCode.Value
                GetSort = "�A�Ȱ�"
           Case optAreaCode.Value
                GetSort = "��F��"
           Case OptInstEn1.Value
                GetSort = "�w�ˤH���@"
           Case optIntroID.Value
                GetSort = "���ФH"
    End Select
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "Function GetSort")
End Function

Private Sub Form_Load()
  On Error GoTo ChkErr
    If Alfa2 Then GetGlobal
    InitializeCOM
    DefaultValue
    Call subGil
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    'Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub DefaultValue()
  On Error GoTo ChkErr
    gimCustStatus.SetQueryCode "1,2"
    optNothing.Value = True
  Exit Sub
ChkErr:
    ErrSub Me.Name, "DefaultValue"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5410A)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub InitializeCOM()
  On Error GoTo ChkErr
    SetMulti gimFaciCode, "CodeNo", "Description", "CD022", "�]�ƶ��إN�X", "�]�ƶ��ئW��"
    SetMulti gimBuy, "CodeNo", "Description", "CD034", "�R��覡�N�X", "�R��覡�W��"
    SetMulti gimClassCode, "CodeNo", "Description", "CD004", "�Ȥ����O�N�X", "�Ȥ����O�W��"
    SetMulti gimCustStatus, "CodeNo", "Description", "CD035", "�Ȥ᪬�A�N�X", "�Ȥ᪬�A�W��"
    SetMulti gimServCode, "CodeNo", "Description", "CD002", "�A�ȰϥN�X", "�A�ȰϦW��"
    SetMulti gimAreaCode, "CodeNo", "Description", "CD001", "��F�ϥN�X", "��F�ϦW��"
    SetMulti gimStrtRange, "CodeNo", "Description", "CD017", "��D�N�X", "��D�W��"
    SetMulti gimInstEn, "EmpNo", "EmpName", "CM003", "���u�N��", "���u�m�W"
    SetMulti gimNodeNo, "CompCode", "CodeNo", "CD047", "CompCode", "NodeNo"
    Call SetgiMulti(gimMediaCode, "CodeNo", "Description", "CD009", "���дC���N�X", "���дC���W��")
    Call SetgiMulti(gimIntroId, "IntroId", "NameP", "So013", "���ФH�N�X", "���ФH�m�W")
    Call SetgiMultiAddItem(gimBillType, "B,I,P,M,T", "���O��,�˾���,�������,���׳�,�{�ɦ��O��")
    Call SetgiMultiAddItem(gimOrder, "A,B,C,D,E,F,G,H,I", "�a�},�]�Ʀw�ˤ��,�Ȥ�s��,�]�ƥN�X,���~�Ǹ�,�ѽX���s��,�A�Ȱ�,�w�ˤH���@,�w�ˤH���G")
    Call SetgiMulti(gimMduId, "MduId", "Name", "SO017", "�j�ӽs��", "�j�ӦW��")
  Exit Sub
ChkErr:
    ErrSub Me.Name, "InitializeCOM"
End Sub

Private Sub SetMulti(objMulti As Object, Code As String, Desc As String, TableName As String, FldCaption1 As String, FldCaption2 As String)
  On Error GoTo ChkErr
    With objMulti
        .SendConn gcnGi
        .FldName1 = Code
        .FldName2 = Desc
        .TableName = TableName
        .FldCaption1 = FldCaption1
        .FldCaption2 = FldCaption2
    End With
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "SetMulti")
End Sub

Private Function subChoose() As Boolean
  On Error GoTo ChkErr
  Dim arrOrder() As String
  Dim varOrder As Variant
  Dim intSort As Integer
  Dim strUse As String
    strGroup = ""
    strChoose = "SO001.CustId=SO004.CustId And SO001.InstAddrNo=SO014.AddrNo "
    If blnExcel Then
        strField = "Select SO004.CustId,SO001.CustName," & _
                    "SO001.Tel1,SO004.SNo,SO004.FaciName," & _
                    "SO004.InstDate,SO004.BuyName," & _
                    "SO004.Amount,SO004.DueDate,SO004.CvtId," & _
                    "SO004.FaciSNo,SO004.IntroName,SO004.Quantity," & _
                    "SO004.InstName1,SO004.InstName2,SO001.ServArea," & _
                    "SO001.InstAddress,SO014.AreaCode,SO004.PRDate," & _
                    "SO004.IPAddress,SO002.CustStatusName,SO004.ContName," & _
                    "SO004.ContTel,SO004.DomicileAddress,SO001.InstAddress," & _
                    "SO004.PTName,SO004.CheckNo,SO001.CMBaudRate,SO004.Email," & _
                    "SO004.DialAccount,SO004.DialPassword,SO014.MDUName," & _
                    "SO014.NodeNo,SO014.CircuitNo "
    
    Else
        strField = "Select SO004.CustId,SO001.CustName," & _
                    "SO001.Tel1,SO004.SNo,SO004.FaciName," & _
                    "SO004.InstDate,SO004.BuyName," & _
                    "SO004.Amount,SO004.DueDate,SO004.CvtId," & _
                    "SO004.FaciSNo,SO004.IntroName,SO004.Quantity," & _
                    "SO004.InstName1,SO004.InstName2,SO001.ServArea," & _
                    "SO001.InstAddress,SO014.AreaCode,SO004.PRDate,SO004.IPAddress "
    End If
  '���
    If gdaInstDate1.GetValue <> "" Then subAnd "SO004.InstDate >= To_Date('" & gdaInstDate1.GetValue & "','YYYYMMDD')"
    If gdaInstDate2.GetValue <> "" Then subAnd "SO004.InstDate < To_Date('" & gdaInstDate2.GetValue & "','YYYYMMDD')+1"
    If gdaPRDate1.GetValue <> "" Then subAnd "SO004.PRDate >= To_Date('" & gdaPRDate1.GetValue & "','YYYYMMDD')"
    If gdaPRDate2.GetValue <> "" Then subAnd "SO004.PRDate < To_Date('" & gdaPRDate2.GetValue & "','YYYYMMDD')+1"
'    If gdaPRDate1.GetValue <> "" Or gdaPRDate2.GetValue <> "" Then strIndex = GetUseIndexStr("SO004", "PRDate")
    
  '�`���B
    If Trim(ginAmount1.Text) <> "" Then subAnd "SO004.Amount>=" & ginAmount1.Value
    If Trim(ginAmount2.Text) <> "" Then subAnd "SO004.Amount<=" & ginAmount2.Value
  
  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("SO001.CompCode=" & gilCompCode.GetCodeNo)
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("(SO001.ServiceType Like '%" & gilServiceType.GetCodeNo & "%' OR SO001.ServiceType is null)")
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("instr(SO001.ServiceType,'" & gilServiceType.GetCodeNo & "')>0")
    'If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO002.ServiceType='" & gilServiceType.GetCodeNo & "'")
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("SO004.ServiceType='" & gilServiceType.GetCodeNo & "'")
  'GiMulti
    If gimBillType.GetQryStr <> "" Then subAnd "SubStr(SO004.SNo,7,1) " & gimBillType.GetQryStr
    If gimFaciCode.GetQryStr <> "" Then subAnd "SO004.FaciCode " & gimFaciCode.GetQryStr
    If gimBuy.GetQryStr <> "" Then subAnd "SO004.BuyCode " & gimBuy.GetQryStr
    If gimClassCode.GetQryStr <> "" Then subAnd "SO001.ClassCode1 " & gimClassCode.GetQryStr
    If gimCustStatus.GetQryStr <> "" Then subAnd "SO002.CustStatuscode " & gimCustStatus.GetQryStr
    If gimServCode.GetQryStr <> "" Then subAnd "SO001.ServCode " & gimServCode.GetQryStr
    If gimAreaCode.GetQryStr <> "" Then subAnd "SO014.AreaCode " & gimAreaCode.GetQryStr
    If gimStrtRange.GetQryStr <> "" Then subAnd "SO001.InstAddrNo=SO014.AddrNo And SO014.StrtCode " & gimStrtRange.GetQryStr
    If gimInstEn.GetQryStr <> "" Then subAnd "SO004.InstEn1 " & gimInstEn.GetQryStr
    If gimNodeNo.GetQryStr <> "" Then subAnd "SO014.NodeNo " & gimNodeNo.GetQryStr
    
    '�[�J���дC��, ���ФH
    If gimMediaCode.GetQryStr <> "" Then Call subAnd("SO004.MediaCode " & gimMediaCode.GetQryStr)
    If gimIntroId.Visible = True And gimIntroId.GetQryStr <> "" Then Call subAnd("SO004.IntroId " & gimIntroId.GetQryStr)
    
    '�j��
    If gimMduId.GetQryStr <> "" Then
        Call subAnd("SO014.MduId " & gimMduId.GetQryStr)
    Else
        If gimMduId.GetDispStr <> "" Then subAnd ("SO014.MduId is Not Null")
    End If

  '�]�ƨϥΪ��p
    Select Case True
           Case optUse.Value
                subAnd "SO004.instdate > nvl(prdate,instdate-1)"
                '���D��1770,��ܨϥΤ��A�|�s����@�_�έp
                subAnd "SO004.InstDate is Not Null and SO004.PRSNO IS NULL"
                strUse = "�ϥΤ�"
           Case optPR.Value
                subAnd "SO004.prdate >= nvl(instdate,prdate-1) "
                strUse = "�v�"
           Case optPRing.Value
                subAnd "SO004.PRSNO IS Not Null AND PRDate IS Null "
                strUse = "���"
           Case optAll.Value
                strUse = "����"
    End Select
    
  '����A
    Select Case True
           Case optM.Value
                subAnd "SubStr(SO004.PRSNO,7,1)='M'"
           Case optP.Value
                subAnd "SubStr(SO004.PRSNO,7,1)='P'"
    End Select
    
  '�j����
    If chkContractCust.Value = 1 Then subAnd "SO004.ContractCust=1"
    
    strChoose = "From SO001,SO002,SO004,SO014 Where SO004.Custid=SO002.Custid And SO004.ServiceType=SO002.ServiceType And " & strChoose
    
  '�����覡
    Select Case True
           Case optServCode.Value '�A�Ȱ�
                strGroup = "ReportType=True;GroupName={SO001.ServArea}"
           Case optAreaCode.Value '��F��
                strGroup = "ReportType=True;GroupName={SO014.AreaName}"
           Case optClctAreaCode.Value '���O��
                strGroup = "ReportType=True;GroupName={SO001.ClctAreaCode}"
           Case OptInstEn1.Value '�w�ˤH���@
                strGroup = "ReportType=True;GroupName={SO004.InstName1}"
           Case optIntroID.Value '���ФH
                strGroup = "ReportType=True;GroupName={SO004.IntroId}"
           Case optNothing.Value '�L
                Select Case Left(gimOrder.GetColumnOrderCode, 1)
                '�ק���D��2883�ɡ@�o�{�p�G�����覡��ܵL�A�a�}�ƧǷ|�����T�A�쥻��SO014.AddrNo 2006/12/6
                  Case "A"
                    strGroup = "ReportType=False;GroupName={SO014.AddrSort}"
                  Case "B"
                    strGroup = "ReportType=False;GroupName=Date({SO004.InstDate})"
                  Case "C"
                    strGroup = "ReportType=False;GroupName={SO001.CustID}"
                  Case "D"
                    strGroup = "ReportType=False;GroupName={SO004.FaciCode}"
                  Case "E"
                    strGroup = "ReportType=False;GroupName={SO004.FaciSNo}"
                  Case "F"
                    strGroup = "ReportType=False;GroupName={SO004.CvtId}"
                  Case "G"
                    strGroup = "ReportType=False;GroupName={SO014.ServCode}"
                  Case "H"
                    strGroup = "ReportType=False;GroupName={SO004.InstEn1}"
                  Case "I"
                    strGroup = "ReportType=False;GroupName={SO004.InstEn2}"
                  Case Else
                    strGroup = "ReportType=False;GroupName={SO001.CustID}"
                End Select
    End Select
       
  '�ƧǤ覡
    If gimOrder.GetColumnOrderCode <> "" Then
      arrOrder = Split(gimOrder.GetColumnOrderCode, ",")
      intSort = 0
      For Each varOrder In arrOrder
        Select Case arrOrder(intSort)
          Case "A"
            strGroup = strGroup & ";Sort" & intSort & "={SO014.AddrSort}"
          Case "B"
            strGroup = strGroup & ";Sort" & intSort & "=Date({SO004.InstDate})"
          Case "C"
            strGroup = strGroup & ";Sort" & intSort & "={SO001.CustID}"
          Case "D"
            strGroup = strGroup & ";Sort" & intSort & "={SO004.FaciCode}"
          Case "E"
            strGroup = strGroup & ";Sort" & intSort & "={SO004.FaciSNo}"
          Case "F"
            strGroup = strGroup & ";Sort" & intSort & "={SO004.CvtId}"
          Case "G"
            strGroup = strGroup & ";Sort" & intSort & "={SO014.ServCode}"
          Case "H"
            strGroup = strGroup & ";Sort" & intSort & "={SO004.InstEn1}"
          Case "I"
            strGroup = strGroup & ";Sort" & intSort & "={SO004.InstEn2}"

        End Select
        intSort = intSort + 1
      Next
    End If
    
  '����榡
    If optDail Then
        strGroup = strGroup & ";SuppressDail=False"
    Else
        strGroup = strGroup & ";SuppressDail=True"
    End If
    
    If gilServiceType.GetCodeNo & "" <> "I" Then
        strGroup = strGroup & ";blnServiceType=False"
    Else
        strGroup = strGroup & ";blnServiceType=True"
    End If
    subChoose = True
    
    strChooseString = "�]�Ʀw�ˤ��:" & subSpace(gdaInstDate1.GetOriginalValue) & "~" & subSpace(gdaInstDate2.GetOriginalValue) & ";" & _
                      "�]�Ʃ���:" & subSpace(gdaPRDate1.GetOriginalValue) & "~" & subSpace(gdaPRDate2.GetOriginalValue) & ";" & _
                      "�`���B  :" & subSpace(ginAmount1.Text) & "~" & subSpace(ginAmount2.Text) & ";" & _
                      "���q�O  :" & subSpace(gilCompCode.GetDescription) & ";" & _
                      "�A�����O:" & subSpace(gilServiceType.GetDescription) & ";" & _
                      "������O: " & subSpace(gimBillType.GetDispStr) & ";" & _
                      "�]�ƶ���:" & subSpace(gimFaciCode.GetDispStr) & ";" & _
                      "�R��覡:" & subSpace(gimBuy.GetDispStr) & ";" & _
                      "�Ȥ����O:" & subSpace(gimClassCode.GetDispStr) & ";" & _
                      "�Ȥ᪬�A:" & subSpace(gimCustStatus.GetDispStr) & ";" & _
                      "�A�Ȱ�:" & subSpace(gimServCode.GetDispStr) & ";" & _
                      "��F��:" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                      "��D�W��:" & subSpace(gimStrtRange.GetDispStr) & ";" & _
                      "�w�ˤH��:" & subSpace(gimInstEn.GetDispStr) & ";" & _
                      "�ƧǤ覡:" & subSpace(gimOrder.GetColumnOrderDspStr) & ";" & _
                      "�j�ӽs��:" & subSpace(gimMduId.GetDispStr) & ";" & _
                      "NodeNo:" & subSpace(gimNodeNo.GetDispStr) & ";" & _
                      "���дC��:" & subSpace(gimMediaCode.GetDispStr) & ";" & _
                      "���ФH�@:" & subSpace(gimIntroId.GetDispStr) & ";" & _
                      "�]�ƨϥΪ��p:" & strUse & ";" & _
                      "�j����:" & IIf(chkContractCust.Value = 1, "�O", "���z��") & ";" & _
                      "����A:" & IIf(optM.Value = True, "���׬���", "") & IIf(optP.Value = True, "�������", "") & ";" & _
                      "�����覡:" & GetSort & ";" & _
                      "����榡:" & IIf(optDail.Value, "���Ӫ�", "�J�`��")
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  Call DropTMPVIEW(strViewName, strSubViewName())
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5410A
End Sub
Private Sub gdaPRDate1_GotFocus()
  On Error Resume Next
    If gdaPRDate1.GetValue = "" Then gdaPRDate1.SetValue (RightDate)
End Sub

Private Sub gdaPRDate2_GotFocus()
  On Error Resume Next
    If gdaPRDate1.GetValue = "" Or gdaPRDate2.GetValue = "" Then gdaPRDate2.SetValue gdaPRDate1.GetValue
End Sub

Private Sub gdaPRDate2_Validate(Cancel As Boolean)
  On Error Resume Next
    Cancel = ChkDate2(gdaPRDate1, gdaPRDate2)
End Sub
Private Sub gdaInstDate1_GotFocus()
  On Error Resume Next
    If gdaInstDate1.GetValue = "" Then gdaInstDate1.SetValue (RightDate)
End Sub

Private Sub gdaInstDate2_GotFocus()
  On Error Resume Next
    If gdaInstDate1.GetValue = "" Or gdaInstDate2.GetValue = "" Then gdaInstDate2.SetValue gdaInstDate1.GetValue
End Sub

Private Sub gdaInstDate2_Validate(Cancel As Boolean)
  On Error Resume Next
    Cancel = ChkDate2(gdaInstDate1, gdaInstDate2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtRange, , gilCompCode.GetCodeNo
    GiMultiFilter gimInstEn, , gilCompCode.GetCodeNo
    GiMultiFilter gimNodeNo, gilServiceType.GetCodeNo, gilCompCode.GetCodeNo
    GiMultiFilter gimMduId, , gilCompCode.GetCodeNo
    GiMultiFilter gimIntroId, , gilCompCode.GetCodeNo
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
    GiMultiFilter gimFaciCode, gilServiceType.GetCodeNo
    GiMultiFilter gimClassCode, gilServiceType.GetCodeNo
    GiMultiFilter gimBuy, gilServiceType.GetCodeNo
    GiMultiFilter gimNodeNo, gilServiceType.GetCodeNo, gilCompCode.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Function CreateSubView3(ByVal fStrWhere As String) As Boolean
  Dim strView As String
    On Error Resume Next
      gcnGi.Execute "Drop View " & strViewName
    On Error GoTo ChkErr
        strViewName = GetTmpViewName
        
        Select Case True
               Case optUse.Value '�ϥΤ�
                    strView = "Select SO004.FaciName,SO004.FaciCode,so014.NODENO,'�ϥΤ�' TYPE,COUNT(*) intCounts " & fStrWhere & " GROUP BY SO004.FaciName,SO004.FaciCode,SO014.NODENO "
               Case optPR.Value  '�v�
                    strView = "Select SO004.FaciName,SO004.FaciCode,so014.NODENO,'�v�' TYPE,COUNT(*) intCounts " & fStrWhere & " GROUP BY SO004.FaciName,SO004.FaciCode,SO014.NODENO "
               Case optPRing.Value '���
                    strView = "Select SO004.FaciName,SO004.FaciCode,so014.NODENO,'���' TYPE,COUNT(*) intCounts " & fStrWhere & " GROUP BY SO004.FaciName,SO004.FaciCode,SO014.NODENO "
               Case optAll.Value '����
                    strView = "Select SO004.FaciName,SO004.FaciCode,so014.NODENO,'�ϥΤ�' TYPE,COUNT(*) intCounts " & fStrWhere & " AND instdate > nvl(prdate,instdate-1) GROUP BY SO004.FaciName,SO004.FaciCode,SO014.NODENO " & _
                              "Union ALL " & _
                              "Select SO004.FaciName,SO004.FaciCode,so014.NODENO,'�v�' TYPE,COUNT(*) intCounts " & fStrWhere & " AND prdate >= nvl(instdate,prdate-1) GROUP BY SO004.FaciName,SO004.FaciCode,SO014.NODENO " & _
                              "Union All " & _
                              "Select SO004.FaciName,SO004.FaciCode,so014.NODENO,'�˾���' TYPE,COUNT(*) intCounts " & fStrWhere & " AND InstDate is Null And PRDate is Null GROUP BY SO004.FaciName,SO004.FaciCode,SO014.NODENO "
        End Select

        strView = "Create View " & strViewName & " as (" & strView & ")"
        gcnGi.Execute strView
        SendSQL strView, True
        CreateSubView3 = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "CreateSubView3"
End Function

Private Sub gimMediaCode_Change()
  On Error GoTo ChkErr
    If gimMediaCode.GetQueryCode = "" Then gimIntroId.Visible = False: Exit Sub
    Select Case InStr(gimMediaCode.GetQueryCode, ",")
           Case Is > 1
                gimIntroId.Visible = False
                'lblCustId.Visible = False
                'txtCustId.Visible = False
           Case 0
                'Call ChkRefNo
                ChangIntroVisible Me, gimMediaCode.GetQryStr, gilCompCode.GetCodeNo
    End Select
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "gimMediaCode_Change")
End Sub

Private Sub optDail_Click()
    If optDail.Value = True Then
        cmdExcel.Enabled = True
    Else
        cmdExcel.Enabled = False
    End If
End Sub

Private Sub optSummaries_Click()
    If optSummaries.Value = True Then
        cmdExcel.Enabled = False
    Else
        cmdExcel.Enabled = True
    End If
End Sub
