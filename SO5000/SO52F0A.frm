VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO52F0A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "���O�覡����� [SO52F0A]"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   Icon            =   "SO52F0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '���ݵ�������
   Visible         =   0   'False
   Begin VB.CommandButton cmdExcel 
      Caption         =   "�צ�Excel"
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
      Height          =   525
      Left            =   3480
      TabIndex        =   29
      Top             =   6555
      Width           =   1245
   End
   Begin VB.Frame fraGroup 
      Caption         =   "�έp�̾�"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   705
      Left            =   30
      TabIndex        =   42
      Top             =   5655
      Width           =   4815
      Begin VB.OptionButton optOldClctEn 
         Caption         =   "�즬�O�H��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   3270
         TabIndex        =   26
         Top             =   315
         Width           =   1320
      End
      Begin VB.OptionButton optAcceptTime 
         Caption         =   "�ꦬ���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   210
         TabIndex        =   24
         Top             =   315
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optNode 
         Caption         =   "���O�H��+��D"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   1500
         TabIndex        =   25
         Top             =   315
         Width           =   1650
      End
   End
   Begin VB.TextBox txtPrtSNo1 
      Enabled         =   0   'False
      Height          =   330
      Left            =   4785
      MaxLength       =   12
      TabIndex        =   5
      Top             =   90
      Width           =   1380
   End
   Begin VB.TextBox txtPrtSNo2 
      Enabled         =   0   'False
      Height          =   330
      Left            =   6360
      MaxLength       =   12
      TabIndex        =   6
      Top             =   90
      Width           =   1380
   End
   Begin VB.Frame FrmAdd 
      Caption         =   "�a�}�̾�"
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
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   4860
      TabIndex        =   38
      Top             =   4920
      Width           =   3675
      Begin VB.OptionButton optInstAddress 
         Caption         =   "�˾��a�}"
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   420
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optChargeAddress 
         Caption         =   "���O�a�}"
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1830
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.Frame frmNetNo 
      Caption         =   "�����s��"
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
      Height          =   735
      Left            =   30
      TabIndex        =   35
      Top             =   4920
      Width           =   4800
      Begin VB.OptionButton optCircuitNo 
         Caption         =   "�u�L�ťպ����s��"
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1680
         TabIndex        =   20
         Top             =   330
         Width           =   1890
      End
      Begin VB.OptionButton optAll 
         Caption         =   "����"
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3795
         TabIndex        =   21
         Top             =   330
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.TextBox txtCircuitNo 
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   105
         TabIndex        =   19
         Top             =   240
         Width           =   1545
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:�C�L"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   75
      TabIndex        =   27
      Top             =   6555
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
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
      Height          =   525
      Left            =   7305
      TabIndex        =   30
      Top             =   6555
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "�W���έp���G"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1365
      TabIndex        =   28
      Top             =   6555
      Width           =   1395
   End
   Begin Gi_Date.GiDate gdaRealDate2 
      Height          =   345
      Left            =   2370
      TabIndex        =   2
      Top             =   510
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   16711680
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
   Begin Gi_Date.GiDate gdaRealDate1 
      Height          =   345
      Left            =   930
      TabIndex        =   1
      Top             =   510
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   16711680
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
   Begin prjGiList.GiList gilServiceType 
      Height          =   345
      Left            =   4785
      TabIndex        =   8
      Top             =   945
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   609
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
   Begin prjGiList.GiList gilCompCode 
      Height          =   345
      Left            =   4785
      TabIndex        =   7
      Top             =   510
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   609
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
   Begin Gi_Multi.GiMulti gimCMCode 
      Height          =   345
      Left            =   15
      TabIndex        =   10
      Top             =   1710
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   609
      ButtonCaption   =   "��  �O  ��  ��"
   End
   Begin Gi_Multi.GiMulti gimBillType 
      Height          =   345
      Left            =   15
      TabIndex        =   13
      Top             =   2805
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   609
      ButtonCaption   =   "��  ��  ��  �O"
      DataType        =   2
      DIY             =   -1  'True
      Exception       =   -1  'True
   End
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   345
      Left            =   15
      TabIndex        =   12
      Top             =   2445
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   609
      ButtonCaption   =   "��     �F     ��"
   End
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   345
      Left            =   15
      TabIndex        =   11
      Top             =   2085
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   609
      ButtonCaption   =   "�A     ��     ��"
   End
   Begin CS_Multi.CSmulti gimClctEn 
      Height          =   345
      Left            =   15
      TabIndex        =   14
      Top             =   3150
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   609
      ButtonCaption   =   "��  �O  �H   ��"
   End
   Begin CS_Multi.CSmulti gimClassCode 
      Height          =   345
      Left            =   15
      TabIndex        =   16
      Top             =   3855
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   609
      ButtonCaption   =   "��  ��  ��  �O"
   End
   Begin CS_Multi.CSmulti gimStrtCode 
      Height          =   345
      Left            =   15
      TabIndex        =   18
      Top             =   4590
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   609
      ButtonCaption   =   "��  �D  �d  ��"
   End
   Begin Gi_Date.GiDate gdaShouldDate2 
      Height          =   345
      Left            =   2370
      TabIndex        =   4
      Top             =   945
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
   Begin Gi_Date.GiDate gdaShouldDate1 
      Height          =   345
      Left            =   930
      TabIndex        =   3
      Top             =   945
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
   Begin Gi_Multi.GiMulti gimCustStatus 
      Height          =   345
      Left            =   15
      TabIndex        =   17
      Top             =   4215
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   609
      ButtonCaption   =   "��  ��  ��  �A"
   End
   Begin Gi_Multi.GiMulti gimPWCode 
      Height          =   345
      Left            =   15
      TabIndex        =   15
      Top             =   3510
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   609
      ButtonCaption   =   "�I  �O  �N  �@"
   End
   Begin Gi_YM.GiYM gymPrtSNo 
      Height          =   345
      Left            =   930
      TabIndex        =   0
      Top             =   90
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   345
      Left            =   15
      TabIndex        =   9
      Top             =   1335
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   609
      ButtonCaption   =   "��  �O  ��  ��"
   End
   Begin VB.Label lblClctYM 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�L��~��"
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   41
      Top             =   180
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   6210
      TabIndex        =   40
      Top             =   165
      Width           =   90
   End
   Begin VB.Label lblPrtSNo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
      Caption         =   "�L��Ǹ�"
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3885
      TabIndex        =   39
      Top             =   180
      Width           =   780
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2160
      TabIndex        =   37
      Top             =   1035
      Width           =   105
   End
   Begin VB.Label lblShouldDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   1035
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�ꦬ���"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   607
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2160
      TabIndex        =   33
      Top             =   600
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
      Caption         =   "�A�����O"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3885
      TabIndex        =   32
      Top             =   1035
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
      Caption         =   "��  �q  �O"
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
      Height          =   195
      Left            =   3885
      TabIndex        =   31
      Top             =   607
      Width           =   765
   End
End
Attribute VB_Name = "frmSO52F0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table:SO033 or SO034 A
Option Explicit
Dim strAddressField As String
Dim strFrom As String
Dim strWhere As String
Dim strAddress As String
Dim blnExcel As Boolean

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
      If optNode.Value Then
          Call PreviousRpt(GetPrinterName(5), RptName("SO52F0", 1), Me.Caption)
      Else
          Call PreviousRpt(GetPrinterName(5), RptName("SO52F0"), Me.Caption)
      End If
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
        Call subCreateView
        Call subPrint
      Me.Enabled = True
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Sub subPrint()
  On Error GoTo ChkErr
    Dim rsCmcode As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim intCountR As Long
    Dim interror As Integer
    Dim intFC As Integer
    Dim strCmcode As String, strCm As String
    Dim strField As String
    interror = 1
        Call CreateTable(intCountR)
        If intCountR = 0 Then
            MsgNoRcd
            SendSQL , , True
        Else
            If blnExcel Then    '���D��1180 �W�[�צ�excel �p�����X 20041006 Edit by Crystal
                intFC = 3
                strSQL = "SELECT DISTINCT CMCODE, CMNAME FROM SO52F0A WHERE CMCODE<>-2 AND CMCODE<>-1"
                If Not GetRS(rsCmcode, strSQL, cnn) Then Exit Sub
                While Not rsCmcode.EOF
                    strCmcode = strCmcode & "," & rsCmcode(0)
                    strCm = strCm & "," & rsCmcode(0) & rsCmcode(1)
                    strField = strField & " ,sum(IIf([CMCode] = " & rsCmcode(0) & ", [REALCOUNT], 0)) As CMCode" & Format(intFC, "0#")
                    intFC = intFC + 1
                    rsCmcode.MoveNext
                    DoEvents
                Wend
                strSQL = "Select GroupName,StrtName, sum(IIf([CMCode] = -2,[REALCOUNT],0)) as CMCode01,sum(IIf([CMCode] =-1,[REALCOUNT],0)) as CMCode02" & _
                         strField & ",sum(IIf([CMCode] is null,[REALCOUNT],0)) as CMCode00 from SO52F0A GROUP BY GroupName,StrtName union all " & _
                         "Select GroupName, chr(255) & '�p�p' as [STRTNAME],sum(IIf([CMCode] = -2,[REALCOUNT],0)) as CMCode01,sum(IIf([CMCode] =-1,[REALCOUNT],0)) as CMCode02" & _
                         strField & ",sum(IIf([CMCode] is null,[REALCOUNT],0)) as CMCode00 from SO52F0A GROUP BY GroupName " & _
                         "union all select  chr(255) &'�X�p' as [GROUPNAME], '' as [STRTNAME],sum(IIf([CMCode] = -2,[REALCOUNT],0)) as CMCode01,sum(IIf([CMCode] =-1,[REALCOUNT],0)) as CMCode02" & _
                         strField & ",sum(IIf([CMCode] is null,[REALCOUNT],0)) as CMCode00 from SO52F0A Order by GroupName,StrtName"
                Call toExcel(strSQL, strCm)
            Else
            
                strSQL = "SELECT * From SO52F0A V"
                If optNode.Value Then
                    Call PrintRpt(GetPrinterName(5), RptName("SO52F0", 1), , Me.Caption, strSQL, strChooseString, , True, "Tmp0000.MDB", , , GiPaperLandscape)
                Else
                    Call PrintRpt(GetPrinterName(5), RptName("SO52F0"), , Me.Caption, strSQL, strChooseString, , True, "Tmp0000.MDB", , , GiPaperLandscape)
                End If
            End If
      End If
      CloseRecordset rsTmp
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subPrint")
End Sub
Private Sub toExcel(ByVal strSQL As String, strCm As String)
  On Error GoTo ChkErr
    Dim rsExcel As New ADODB.Recordset
    Dim rsSubExcel(3) As New ADODB.Recordset
    Dim intFor As Integer
'      strsql = strsql & " Order by " & Mid(strOrder, 2)
      If Not GetRS(rsExcel, strSQL, cnn) Then Exit Sub
         '#7074 To Add a choice-condition parameter into excel print function by Kin 2015/09/09
      Call UseProperty(rsExcel, "���O�覡�����", "�Ĥ@��", "�s��/�W��,��D,�����X�p,�ꦬ�X�p" & strCm & ", ", _
                            , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , strChooseString)
      CloseRecordset rsExcel
      blnExcel = False
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subExcel")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If Len(gilCompCode.GetCodeNo) = 0 Then gilCompCode.SetFocus: strErrFile = "���q�O": GoTo 66
    If optAcceptTime Then
        If gdaRealDate1.GetValue = "" Then gdaRealDate1.SetFocus: strErrFile = "���O����_�l��": GoTo 66
        If gdaRealDate2.GetValue = "" Then gdaRealDate2.SetFocus: strErrFile = "���O����I���": GoTo 66
        If gdaRealDate2.GetValue(True) >= DateAdd("YYYY", 1, gdaRealDate1.Text) Then
            MsgBox "���O��������W�L�@�~", , "���~�T��"
            gdaRealDate2_GotFocus
            Exit Function
        End If
    End If
'    If optNode Then
'        If gymPrtSNo.Text = "" Then gymPrtSNo.SetFocus: strErrFile = "�L��~��": GoTo 66
'    End If
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
    Dim strTypeName As String
      strChoose = " A.CancelFlag=0 "
      strChooseString = ""
      strAddressField = ""
      strAddress = ""
       strFrom = "": strWhere = " Where"
    '�L��~��
      If gymPrtSNo.GetValue <> "" Then Call subAnd("A.PrtSNo >= '" & gymPrtSNo.GetValue & "' And A.PrtSNo < '" & gymPrtSNo.GetValue & "A'")
      
    '�ꦬ���
      If gdaRealDate1.GetValue <> "" Then Call subAnd("A.RealDate >= To_Date('" & gdaRealDate1.GetValue & "','YYYYMMDD')")
      If gdaRealDate2.GetValue <> "" Then Call subAnd("A.RealDate < To_Date('" & gdaRealDate2.GetValue & "','YYYYMMDD')+1")
    '�������
      If gdaShouldDate1.GetValue <> "" Then Call subAnd("A.ShouldDate >= To_Date('" & gdaShouldDate1.GetValue & "','YYYYMMDD')")
      If gdaShouldDate2.GetValue <> "" Then Call subAnd("A.ShouldDate < To_Date('" & gdaShouldDate2.GetValue & "','YYYYMMDD')+1")
    '�L��Ǹ�
      If txtPrtSNo1.Text <> "" Then Call subAnd("A.PrtSNo >= '" & Trim(txtPrtSNo1.Text) & "'")
      If txtPrtSNo2.Text <> "" Then Call subAnd("A.PrtSNo <= '" & Trim(txtPrtSNo2.Text) & "'")
    '�����s��
      If txtCircuitNo.Text = "" Then
         If optCircuitNo.Value = True Then Call subAnd("So014.CircuitNo Is Null")
      Else
         Call subAnd("'" & txtCircuitNo.Text & "'=SO014.CircuitNo")
      End If
    '�a�}�̾�
    If optNode Then
      Select Case True
             Case optInstAddress.Value
                  strAddressField = ",SO001.InstAddrNo AddrNo,SO001.InstAddress Address "
                  strFrom = strFrom & ",SO014 SO014"
                  strWhere = strWhere & " B.InstAddrNo=SO014.Addrno And "
                  strAddress = "�˾��a�}"
             Case optChargeAddress.Value
                  strAddressField = ",SO001.ChargeAddrNo AddrNo,SO001.ChargeAddress Address "
                  strFrom = strFrom & ",SO014 SO014"
                  strWhere = strWhere & " B.ChargeAddrNo=SO014.Addrno And "
                  strAddress = "���O�a�}"
'             Case optMailAddress.Value
'                  strAddressField = ",SO001.MailAddrNo AddrNo,SO001.MailAddress Address "
'                  strFrom = strFrom & ",SO014 SO014"
'                  strWhere = strWhere & " B.MailAddrNo=SO014.Addrno And "
'                  strAddress = "�l�H�a�}"
      End Select
     End If
    'GiList
      If gilCompCode.GetCodeNo <> "" Then Call subAnd("A.CompCode=" & gilCompCode.GetCodeNo)
      If gilServiceType.GetCodeNo <> "" Then Call subAnd("A.ServiceType='" & gilServiceType.GetCodeNo & "'")

    'GiMulti
      If gimCMCode.GetQryStr <> "" Then Call subAnd("A.CMCode " & gimCMCode.GetQryStr)
       'new
      If gimServCode.GetQryStr <> "" Then Call subAnd("SO014.ServCode " & gimServCode.GetQryStr)
      If gimAreaCode.GetQryStr <> "" Then Call subAnd("SO014.AreaCode " & gimAreaCode.GetQryStr)
      If gimBillType.GetQryStr <> "" Then Call subAnd("SubStr(A.BillNo,7,1) " & gimBillType.GetQryStr)
      If gimClctEn.GetQryStr <> "" Then Call subAnd("A.OLDClctEn " & gimClctEn.GetQryStr)
      If gimPWCode.GetQryStr <> "" Then Call subAnd("B.PWCode " & gimPWCode.GetQryStr)
      If gimClassCode.GetQryStr <> "" Then Call subAnd("B.ClassCode1 " & gimClassCode.GetQryStr)
      If gimCustStatus.GetQryStr <> "" Then Call subAnd("SO002.CustStatusCode " & gimCustStatus.GetQryStr)
      If gimStrtCode.GetQryStr <> "" Then Call subAnd("SO014.StrtCode " & gimStrtCode.GetQryStr)
      If gimCitemCode.GetQryStr <> "" Then Call subAnd("A.CitemCode " & gimCitemCode.GetQryStr)
     
      strChooseString = "�L��~��:" & subSpace(gymPrtSNo.GetValue(True)) & ";" & _
                        "���O���:" & subSpace(gdaRealDate1.GetValue(True)) & "~" & subSpace(gdaRealDate2.GetValue(True)) & ";" & _
                        "�������:" & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & _
                        "�L��Ǹ�:" & subSpace(Trim(txtPrtSNo1)) & "~" & subSpace(Trim(txtPrtSNo2)) & ";" & _
                        "�����s��:" & subSpace(txtCircuitNo) & ";" & _
                        "���q�O�@:" & subSpace(gilCompCode.GetDescription) & ";" & _
                        "�A�����O:" & subSpace(gilServiceType.GetDescription) & ";" & _
                        "���O����: " & subSpace(gimCitemCode.GetDispStr) & ";" & _
                        "���O�覡:" & subSpace(gimCMCode.GetDispStr) & ";" & _
                        "�A�Ȱϡ@:" & subSpace(gimServCode.GetDispStr) & ";" & _
                        "��F�ϡ@:" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                        "���O�H��:" & subSpace(gimClctEn.GetDispStr) & ";" & _
                        "�I�O�N�@:" & subSpace(gimPWCode.GetDispStr) & ";" & _
                        "�Ȥ����O:" & subSpace(gimClassCode.GetDispStr) & ";" & _
                        "�Ȥ᪬�A:" & subSpace(gimCustStatus.GetDispStr) & ";" & _
                        "��D�d��:" & subSpace(gimStrtCode.GetDispStr) & ";" & _
                        "������O:" & subSpace(gimBillType.GetDispStr) & ";" & _
                        "�a�}�̾�:" & strAddress
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub
'Private Sub txtCircuitNo_Change()
'    On Error Resume Next
'        If txtCircuitNo = "" Then
'            fraCircuitNo.Visible = True
'        Else
'            fraCircuitNo.Visible = False
'        End If
'End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO52F0A)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    If Alfa2 Then Call GetGlobal
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
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "���O�覡�N�X", "���O�覡�W��")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "�A�ȰϥN�X", "�A�ȰϦW��")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "��F�ϥN�X", "��F�ϦW��")
    Call SetgiMultiAddItem(gimBillType, "B,I,P,M,T", "���O��,�˾���,�������,���׳�,�{�ɦ��O��", "��ڥN��", "��ڦW��")
    Call SetgiMulti(gimPWCode, "CodeNo", "Description", "CD020", "�I�O�N�@�N�X", "�I�O�N�@�W��")
    Call SetgiMulti(gimClctEn, "EmpNo", "EmpName", "CM003", "���O���N��", "���O���W��")
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "�Ȥ����O�N�X", "�Ȥ����O�W��")
    Call SetgiMulti(gimCustStatus, "CodeNo", "Description", "CD035", "�Ȥ᪬�A�N�X", "�Ȥ᪬�A�W��")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "��D�N�X", "��D�W��")
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "���O���إN�X", "���O���ئW��")
    
    
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    'Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
    gcnGi.Execute "Drop View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO52F0A
End Sub

Private Sub gdaRealDate1_GotFocus()
  On Error Resume Next
  If gdaRealDate1.GetValue = "" Then gdaRealDate1.SetValue (RightDate)
End Sub

Private Sub gdaRealDate2_GotFocus()
  On Error Resume Next
    If gdaRealDate1.GetValue = "" Or gdaRealDate2.GetValue = "" Then gdaRealDate2.SetValue gdaRealDate1.GetValue
End Sub

Private Sub gdaRealDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaRealDate1, gdaRealDate2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimClassCode, gilServiceType.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
    
    gilServiceType.ListIndex = 1
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
    GiMultiFilter gimCMCode, gilServiceType.GetCodeNo
    GiMultiFilter gimPWCode, gilServiceType.GetCodeNo
    Call GiMultiFilter(gimCitemCode, gilServiceType.GetCodeNo)

  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilServiceType_Change")
End Sub

Private Function subCreateView() As Boolean
  Dim strView As String
  Dim strField As String, strField1 As String, strField2 As String
  Dim strGroup As String, intPara24 As Integer, strBillNo As String
'  strWhere = " Where "
'  strFrom = ""
  
    If InStr(1, strChoose, "B.") > 0 Or InStr(1, strWhere, "B.") > 0 Then
        strFrom = strFrom & ",SO001 B"
        strWhere = strWhere & " A.CustId=B.CustId AND "
    End If
    
    If InStr(1, strChoose, "SO002.") > 0 Then
        strFrom = strFrom & ",SO002 SO002"
        strWhere = strWhere & " A.CUSTID=SO002.CUSTID AND A.ServiceType=SO002.ServiceType AND "
    End If
    If optAcceptTime Then
        If InStr(1, strChoose, "SO014.") > 0 Then
            strFrom = strFrom & ",SO014 SO014"
            strWhere = strWhere & " A.AddrNo=SO014.AddrNo AND "
        End If
    End If
    On Error Resume Next
      gcnGi.Execute "Drop View " & strViewName
      
    On Error GoTo ChkErr
        strViewName = GetTmpViewName
        If optNode.Value Then
        
'            strView = "Create View " & strViewName & " As ( " & _
                      " SELECT SO014.ClctEn || SO014.ClctName as GroupName,SO014.StrtName,Count(A.CUSTID) RealCount ,A.CMCode,A.CMName " & _
                      " FROM SO033 A " & strFrom & strWhere & strChoose & IIf(strChoose = "", "", " And") & _
                      " A.UCCode Is Null " & _
                      " GROUP BY SO014.ClctEn || SO014.ClctName,SO014.StrtName,A.CMCode,A.CMName Union All " & _
                      " SELECT SO014.ClctEn || SO014.ClctName as GroupName,SO014.StrtName,Count(A.CUSTID) ShouldCount,-2 CMCode,'�����X�p' CMName " & _
                      " FROM SO033 A " & strFrom & strWhere & strChoose & _
                      " GROUP BY SO014.ClctEn || SO014.ClctName,SO014.StrtName Union All " & _
                      " SELECT SO014.ClctEn || SO014.ClctName as GroupName,SO014.StrtName,Count(A.CUSTID) ShouldCount,-1 CMCode,'�ꦬ�X�p' CMName " & _
                      " FROM SO033 A " & strFrom & strWhere & strChoose & IIf(strChoose = "", "", " And") & _
                      " UCCode Is Null " & _
                      " GROUP BY SO014.ClctEn || SO014.ClctName,SO014.StrtName Union All" & _
                      " SELECT SO014.ClctEn || SO014.ClctName as GroupName,SO014.StrtName,Count(A.CUSTID) RealCount ,A.CMCode,A.CMName " & _
                      " FROM SO034 A " & strFrom & strWhere & strChoose & IIf(strChoose = "", "", " And") & _
                      " A.UCCode Is Null " & _
                      " GROUP BY SO014.ClctEn || SO014.ClctName,SO014.StrtName,A.CMCode,A.CMName Union All " & _
                      " SELECT SO014.ClctEn || SO014.ClctName as GroupName,SO014.StrtName,Count(A.CUSTID) ShouldCount,-2 CMCode,'�����X�p' CMName " & _
                      " FROM SO034 A " & strFrom & strWhere & strChoose & _
                      " GROUP BY SO014.ClctEn || SO014.ClctName,SO014.StrtName Union All " & _
                      " SELECT SO014.ClctEn || SO014.ClctName as GroupName,SO014.StrtName,Count(A.CUSTID) ShouldCount,-1 CMCode,'�ꦬ�X�p' CMName " & _
                      " FROM SO034 A " & strFrom & strWhere & strChoose & IIf(strChoose = "", "", " And") & _
                      " UCCode Is Null " & _
                      " GROUP BY SO014.ClctEn || SO014.ClctName,SO014.StrtName)"
                      
            intPara24 = GetRsValue("Select Para24 From SO043 Where CompCode=" & gilCompCode.GetCodeNo)
            If intPara24 = 0 Then
                strBillNo = "A.BILLNO"
            Else
                strBillNo = "A.MediaBillNo"
            End If
            strField = "SO014.ClctEn || SO014.ClctName GroupName,SO014.StrtCode || SO014.StrtName StrtName,Count(A.CUSTID)  RealCount, COUNT(" & strBillNo & ") BillCount,A.CMCode ,A.CMName || ' / ���' CMName"
            strField1 = "SO014.ClctEn || SO014.ClctName GroupName,SO014.StrtCode || SO014.StrtName StrtName,Count(A.CUSTID) ShouldCount, COUNT(" & strBillNo & ") BillCount,-2 CMCode,'�����X�p / ���' CMName"
            strField2 = "SO014.ClctEn || SO014.ClctName GroupName,SO014.StrtCode || SO014.StrtName StrtName,Count(A.CUSTID) ShouldCount, COUNT(" & strBillNo & ") BillCount,-1 CMCode,'�ꦬ�X�p / ���' CMName"
            strGroup = " GROUP BY SO014.ClctEn || SO014.ClctName,SO014.StrtCode || SO014.StrtName ,A.CMCode,A.CMName"
            strView = "Create View " & strViewName & " As ( " & _
                      " SELECT " & strField & " FROM SO033 A " & strFrom & strWhere & strChoose & IIf(strChoose = "", "", " And") & " A.UCCode Is Null " & strGroup & " Union All " & _
                      " SELECT " & strField1 & " FROM SO033 A " & strFrom & strWhere & strChoose & strGroup & " Union All " & _
                      " SELECT " & strField2 & " FROM SO033 A " & strFrom & strWhere & strChoose & IIf(strChoose = "", "", " And") & " UCCode Is Null " & strGroup & " Union All" & _
                      " SELECT " & strField & " FROM SO034 A " & strFrom & strWhere & strChoose & IIf(strChoose = "", "", " And") & " A.UCCode Is Null " & strGroup & " Union All " & _
                      " SELECT " & strField1 & " FROM SO034 A " & strFrom & strWhere & strChoose & strGroup & " Union All " & _
                      " SELECT " & strField2 & " FROM SO034 A " & strFrom & strWhere & strChoose & IIf(strChoose = "", "", " And") & " UCCode Is Null " & strGroup & ")"
                      
                      
        Else
            If optAcceptTime Then
                strField = "to_char(A.RealDate,'MM')||'��'||to_char(A.RealDate,'DD')||'��'"
                strGroup = strField
            ElseIf optOldClctEn Then
                strField = "OldClctName"
                'strGroup = "rpad(OldClctEn,10,' ')||OldClctName"
                strGroup = "OldClctEn,OldClctName"
            End If
            strView = "Create View " & strViewName & " As ( " & _
                      "SELECT " & strField & " as GroupName,A.CMCode,A.CMName,'���B' as TypeName,sum(A.RealAmt) as RealAmt " & _
                      "From SO033 A " & strFrom & strWhere & strChoose & " Group By " & strGroup & ",A.CMCode,A.CMName  UNION All " & _
                      "SELECT " & strField & " as GroupName,A.CMCode,A.CMName,'����' as TypeName,Count(*) as RealAmt " & _
                      "From SO033 A " & strFrom & strWhere & strChoose & " Group By " & strGroup & ",A.CMCode,A.CMName  UNION All " & _
                      "SELECT " & strField & " as GroupName,999 as CMCode,'�X�p' As CMName,'���B' as TypeName,sum(A.RealAmt) as RealAmt " & _
                      "From SO033 A " & strFrom & strWhere & strChoose & " Group By " & strGroup & ",A.CMCode,A.CMName  UNION All " & _
                      "SELECT " & strField & " as GroupName,999 as CMCode,'�X�p' As CMName,'����' as TypeName,Count(*) as RealAmt " & _
                      "From SO033 A " & strFrom & strWhere & strChoose & " Group By " & strGroup & ",A.CMCode,A.CMName  UNION All " & _
                      "SELECT " & strField & " as GroupName,A.CMCode,A.CMName,'���B' as TypeName,sum(A.RealAmt) as RealAmt " & _
                      "From SO034 A " & strFrom & strWhere & strChoose & " Group By " & strGroup & ",A.CMCode,A.CMName  UNION All " & _
                      "SELECT " & strField & " as GroupName,A.CMCode,A.CMName,'����' as TypeName,Count(*) as RealAmt " & _
                      "From SO034 A " & strFrom & strWhere & strChoose & " Group By " & strGroup & ",A.CMCode,A.CMName  UNION All " & _
                      "SELECT " & strField & " as GroupName,999 as CMCode,'�X�p' As CMName,'���B' as TypeName,sum(A.RealAmt) as RealAmt " & _
                      "From SO034 A " & strFrom & strWhere & strChoose & " Group By " & strGroup & ",A.CMCode,A.CMName  UNION All " & _
                      "SELECT " & strField & " as GroupName,999 as CMCode,'�X�p' As CMName,'����' as TypeName,Count(*) as RealAmt " & _
                      "From SO034 A " & strFrom & strWhere & strChoose & " Group By " & strGroup & ",A.CMCode,A.CMName) "
        End If
        gcnGi.Execute strView
        SendSQL strView, True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function

Private Sub CreateTable(intCountR As Long)
  On Error Resume Next
    cnn.Execute "Drop Table SO52F0A"
  On Error GoTo ChkErr
    If optNode.Value Then
        cnn.Execute "Create Table SO52F0A (GROUPNAME Text(40),STRTNAME Text(200),REALCOUNT long,BillCOUNT long,CMCODE Long,CMNAME Text(255))"
    Else
        cnn.Execute "Create Table SO52F0A (GroupName Text(40),CMCode Long,CMName Text(40),TypeName Text(6),RealAmt long)"
    End If
    Call subInsertMDB("Select * From " & strViewName, intCountR)
'    Call subInsertMDB1("Select * From " & strViewName)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "CreateTable")
End Sub

Private Sub subInsertMDB(strSQL As String, lngCount As Long)
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
      Set rsTmp = gcnGi.Execute(strSQL)
    cnn.BeginTrans
    If optNode.Value Then
        Do While Not rsTmp.EOF
            cnn.Execute "Insert into SO52F0A (GroupName,StrtName,RealCount,BillCount,CMCode,CMName) Values(" & _
                        GetNullString(rsTmp("GroupName")) & "," & _
                        GetNullString(rsTmp("StrtName")) & "," & _
                        GetNullString(rsTmp("RealCount")) & "," & _
                        GetNullString(rsTmp("BillCount")) & "," & _
                        GetNullString(rsTmp("CMCode")) & "," & _
                        GetNullString(rsTmp("CMName")) & ")"
            lngCount = 1
            rsTmp.MoveNext
        Loop
    Else
        Do While Not rsTmp.EOF
            cnn.Execute "Insert into SO52F0A (GroupName,CMCode,CMName,TypeName,RealAmt) Values(" & _
                        GetNullString(rsTmp("GroupName")) & "," & _
                        GetNullString(rsTmp("CMCode")) & "," & _
                        GetNullString(rsTmp("CMName")) & "," & _
                        GetNullString(rsTmp("TypeName")) & "," & _
                        GetNullString(rsTmp("RealAmt")) & ")"
            lngCount = 1
            rsTmp.MoveNext
        Loop
    End If
    cnn.CommitTrans
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "subInsertMDB")
End Sub

Private Sub gdaShouldDate1_GotFocus()
  On Error Resume Next
    If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetValue (RightDate)
End Sub

Private Sub gdaShouldDate2_GotFocus()
  On Error GoTo ChkErr
     If gdaShouldDate1.GetValue = "" Or gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue gdaShouldDate1.GetValue
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaShouldDate2_GotFocus")
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
  On Error Resume Next
    Cancel = ChkDate2(gdaShouldDate1, gdaShouldDate2)
End Sub

Private Sub optAcceptTime_Click()
    If optAcceptTime.Value Then EnableMode False
End Sub

Private Sub optNode_Click()
    If optNode.Value Then EnableMode True
End Sub

Private Sub EnableMode(blnFlag As Boolean)
        FrmAdd.Enabled = blnFlag: optInstAddress.Enabled = blnFlag
        optChargeAddress.Enabled = blnFlag
        lblClctYM.Enabled = blnFlag: gymPrtSNo.Enabled = blnFlag
        lblPrtSNo.Enabled = blnFlag: txtPrtSNo1.Enabled = blnFlag: Label4.Enabled = blnFlag: txtPrtSNo2.Enabled = blnFlag
        frmNetNo.Enabled = blnFlag: txtCircuitNo.Enabled = blnFlag: optCircuitNo.Enabled = blnFlag: optAll.Enabled = blnFlag
        cmdExcel.Enabled = blnFlag

End Sub

Private Sub optOldClctEn_Click()
    If optOldClctEn.Value Then EnableMode False
    lblClctYM.Enabled = True: gymPrtSNo.Enabled = True
    lblPrtSNo.Enabled = True: txtPrtSNo1.Enabled = True: Label4.Enabled = True: txtPrtSNo2.Enabled = True
    
End Sub
