VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.17#0"; "csMulti.ocx"
Begin VB.Form frmSO5145A 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "���O���^���v�έp[SO5145A]"
   ClientHeight    =   5925
   ClientLeft      =   1950
   ClientTop       =   1530
   ClientWidth     =   7305
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO5145A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form14"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame Frame5 
      Caption         =   "�]�w�������B"
      Enabled         =   0   'False
      ForeColor       =   &H00800080&
      Height          =   525
      Left            =   4500
      TabIndex        =   34
      Top             =   4770
      Width           =   2745
      Begin VB.OptionButton OptShouldAmt 
         Caption         =   "�������B"
         Enabled         =   0   'False
         Height          =   225
         Left            =   1440
         TabIndex        =   41
         Top             =   210
         Width           =   1185
      End
      Begin VB.OptionButton OptOldAmt 
         Caption         =   "���������B"
         Enabled         =   0   'False
         Height          =   255
         Left            =   90
         TabIndex        =   40
         Top             =   210
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "�]�w���O�H��"
      Enabled         =   0   'False
      ForeColor       =   &H00800080&
      Height          =   555
      Left            =   4500
      TabIndex        =   33
      Top             =   4080
      Width           =   2745
      Begin VB.OptionButton OptClctName 
         Caption         =   "���O�H��"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   39
         Top             =   180
         Width           =   1245
      End
      Begin VB.OptionButton OptOldClctName 
         Caption         =   "�즬�O�H��"
         Enabled         =   0   'False
         Height          =   225
         Left            =   90
         TabIndex        =   38
         Top             =   210
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "�]�w����榡"
      ForeColor       =   &H00800080&
      Height          =   1185
      Left            =   2820
      TabIndex        =   32
      Top             =   4080
      Width           =   1455
      Begin VB.OptionButton OptPrint3 
         Caption         =   "����榡�T"
         Height          =   225
         Left            =   90
         TabIndex        =   37
         Top             =   900
         Width           =   1305
      End
      Begin VB.OptionButton OptPrint2 
         Caption         =   "����榡�G"
         Height          =   225
         Left            =   90
         TabIndex        =   36
         Top             =   600
         Width           =   1305
      End
      Begin VB.OptionButton OptPrint1 
         Caption         =   "����榡�@"
         Height          =   375
         Left            =   90
         TabIndex        =   35
         Top             =   210
         Value           =   -1  'True
         Width           =   1305
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�L��Ǹ�"
      ForeColor       =   &H00800080&
      Height          =   645
      Left            =   30
      TabIndex        =   31
      Top             =   4080
      Width           =   2655
      Begin VB.OptionButton optPrtSNoAll 
         Caption         =   "����"
         ForeColor       =   &H00800080&
         Height          =   405
         Left            =   1740
         TabIndex        =   16
         Top             =   180
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton optPrtSNoNo 
         Caption         =   "����"
         ForeColor       =   &H00800080&
         Height          =   405
         Left            =   975
         TabIndex        =   15
         Top             =   180
         Width           =   705
      End
      Begin VB.OptionButton optPrtSNoYes 
         Caption         =   "�w��"
         ForeColor       =   &H00800080&
         Height          =   405
         Left            =   210
         TabIndex        =   14
         Top             =   180
         Width           =   705
      End
   End
   Begin VB.CheckBox chkCancel 
      Caption         =   "�@�o��ڬO�_�@�ֲέp"
      Height          =   225
      Left            =   3480
      TabIndex        =   30
      Top             =   1050
      Width           =   3285
   End
   Begin VB.Frame Frame1 
      Height          =   765
      Left            =   0
      TabIndex        =   26
      Top             =   3240
      Width           =   7275
      Begin Gi_Multi.GiMulti gimSTCode 
         Height          =   345
         Left            =   90
         TabIndex        =   13
         Top             =   360
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   609
         ButtonCaption   =   "�u����]"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�ꦬ���B��0�A�B�u����]���U�C�ﶵ�A�h�C���ꦬ"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   120
         Width           =   4380
      End
   End
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   1350
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   661
      ButtonCaption   =   "��  �O  ��  ��"
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "�W���έp���G"
      Height          =   525
      Left            =   1950
      TabIndex        =   18
      Top             =   5340
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "����(&X)"
      Height          =   525
      Left            =   5850
      TabIndex        =   19
      Top             =   5340
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:�C�L"
      Default         =   -1  'True
      Height          =   525
      Left            =   270
      TabIndex        =   17
      Top             =   5340
      Width           =   1245
   End
   Begin Gi_Date.GiDate gdaRealDate2 
      Height          =   345
      Left            =   2220
      TabIndex        =   3
      Top             =   540
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
      Left            =   870
      TabIndex        =   2
      Top             =   540
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
   Begin Gi_Date.GiDate gdaShouldDate1 
      Height          =   345
      Left            =   870
      TabIndex        =   0
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   255
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
   Begin Gi_Date.GiDate gdaShouldDate2 
      Height          =   345
      Left            =   2220
      TabIndex        =   1
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      ForeColor       =   255
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
   Begin Gi_Multi.GiMulti gimBillType 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   2490
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   661
      ButtonCaption   =   "��   ��   ��   �O"
      DataType        =   2
      DIY             =   -1  'True
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   4290
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
      Left            =   4290
      TabIndex        =   7
      Top             =   600
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
   Begin CS_Multi.CSmulti gimClctEn 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   1710
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   661
      ButtonCaption   =   "��  �O  �H   ��"
   End
   Begin CS_Multi.CSmulti gimOldClctEn 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2100
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   661
      ButtonCaption   =   "�� �� �O �H ��"
   End
   Begin Gi_Date.GiDate GiCreateTime2 
      Height          =   345
      Left            =   2220
      TabIndex        =   5
      Top             =   990
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
   Begin Gi_Date.GiDate GiCreateTime1 
      Height          =   345
      Left            =   870
      TabIndex        =   4
      Top             =   990
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
   Begin CS_Multi.CSmulti gimCMCode 
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   2880
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   661
      ButtonCaption   =   "��  �O  ��  ��"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   2070
      TabIndex        =   29
      Top             =   1110
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�X����"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   30
      TabIndex        =   28
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
      Caption         =   "�A�����O"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3420
      TabIndex        =   25
      Top             =   660
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
      Caption         =   "��  �q  �O"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3420
      TabIndex        =   24
      Top             =   180
      Width           =   765
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2040
      TabIndex        =   23
      Top             =   180
      Width           =   105
   End
   Begin VB.Label lblShouldDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�������"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   0
      TabIndex        =   22
      Top             =   120
      Width           =   780
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      Height          =   195
      Left            =   2040
      TabIndex        =   21
      Top             =   630
      Width           =   105
   End
   Begin VB.Label lblRealDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�ꦬ���"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   30
      TabIndex        =   20
      Top             =   630
      Width           =   780
   End
End
Attribute VB_Name = "frmSO5145A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�ϥ�Table: SO033,SO034
Option Explicit
Dim StrTableName As String
Dim strChooseReal As String
Dim strChooseST As String
Dim strFormula  As String
Dim strClct As String
Dim strClct2 As String
Dim strAmt As String
Dim strBillOrMead As String
Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      If OptPrint2 = True Then
       Call PreviousRpt(GetPrinterName(5), RptName("SO5145", "1"), "���O���^���v�έp[SO5145A]")
      ElseIf OptPrint3 = True Then
        Call PreviousRpt(GetPrinterName(5), RptName("SO5145", "2"), "���O���^���v�έp[SO5145A]")
      Else
       Call PreviousRpt(GetPrinterName(5), RptName("SO5145"), "���O���^���v�έp[SO5145A]")
      End If
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click() '�C�L
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
      Screen.MousePointer = vbHourglass
            '����榡
    
    If OptPrint2.Value = True Or OptPrint3.Value = True Then
        If OptOldClctName.Value = True Then
             strClct = "OldClctEn"
             strClct2 = "OldClctName"
           Else
               strClct = "ClctEn"
               strClct2 = "ClctName"
        End If
        If OptOldAmt.Value = True Then
               strAmt = "OldAmt"
           Else
               strAmt = "ShouldAmt"
        End If
     End If

        cmdExit.SetFocus
        Me.Enabled = False
          ReadyGoPrint
          Call subChoose
          subCreateView
          CreateTable
          Set rsTmp = cnn.Execute("SELECT Count(*) as intCount From SO5145A2")
          If rsTmp("intCount") = 0 Then
              MsgNoRcd
              SendSQL , , True
          Else
              If OptPrint2.Value = True Then
                  strsql = "Select * From SO5145A2"
                  Call PrintRpt(GetPrinterName(5), RptName("SO5145", 1), , Me.Caption, strsql, strChooseString, , True, "Tmp0000.MDB", , GiPaperLandscape)
              ElseIf OptPrint1.Value = True Then
                  strsql = "Select * From SO5145A2"
                  Call PrintRpt(GetPrinterName(5), RptName("SO5145"), , Me.Caption, strsql, strChooseString, , True, "Tmp0000.MDB", , strGroupName, GiPaperLandscape)
              ElseIf OptPrint3.Value = True Then
                  strsql = "Select * From SO5145A2"
                  Call PrintRpt(GetPrinterName(5), RptName("SO5145", 2), , Me.Caption, strsql, strChooseString, , True, "Tmp0000.MDB", , strGroupName & strFormula, GiPaperLandscape)
              End If
          End If
        Me.Enabled = True
      Screen.MousePointer = vbDefault
    CloseRecordset rsTmp

  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: strErrFile = "���q�O": GoTo 66
    If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetFocus: strErrFile = "��������_�l��": GoTo 66
    If gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetFocus: strErrFile = "��������I���": GoTo 66
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
  Dim strPrtSNo As String
    strChoose = ""
    strChooseString = ""
    strChooseReal = ""
    strChooseST = ""
    
  '���
    If gdaRealDate1.GetValue <> "" Then Call subAnd2(strChooseReal, "RealDate >= To_Date('" & gdaRealDate1.GetValue & "','YYYYMMDD')")
    If gdaRealDate2.GetValue <> "" Then Call subAnd2(strChooseReal, "RealDate < To_Date('" & gdaRealDate2.GetValue & "','YYYYMMDD')+1")
    If gdaShouldDate1.GetValue <> "" Then Call subAnd("A.ShouldDate >= To_Date('" & gdaShouldDate1.GetValue & "','YYYYMMDD')")
    If gdaShouldDate2.GetValue <> "" Then Call subAnd("A.ShouldDate < To_Date('" & gdaShouldDate2.GetValue & "','YYYYMMDD')+1")
    If GiCreateTime1.GetValue <> "" Then Call subAnd("A.CreateTime >= To_Date('" & GiCreateTime1.GetValue & "','YYYYMMDD')")
    If GiCreateTime2.GetValue <> "" Then Call subAnd("A.CreateTime < To_Date('" & GiCreateTime2.GetValue & "','YYYYMMDD')+1")

  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("A.CompCode=" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("A.ServiceType='" & gilServiceType.GetCodeNo & "'")

  'GiMulti
    If gimCitemCode.GetQryStr <> "" Then Call subAnd("A.CitemCode " & gimCitemCode.GetQryStr)
    If gimClctEn.GetQryStr <> "" Then Call subAnd("A.ClctEn " & gimClctEn.GetQryStr)
    If gimOldClctEn.GetQryStr <> "" Then Call subAnd("A.OldClctEn " & gimOldClctEn.GetQryStr)
    If gimBillType.GetQryStr <> "" Then Call subAnd("subStr(A.BillNo,7,1) " & gimBillType.GetQryStr)
    If gimSTCode.GetQryStr <> "" Then Call subAnd2(strChooseST, "RealAmt=0 And STCode " & gimSTCode.GetQryStr)
    '2006/03/13 ���D��2193 �e������W�["���O�覡" -----Edit by Crystal for Debby
    If gimCMCode.GetQryStr <> "" Then Call subAnd("A.CMCode " & gimCMCode.GetQryStr)
  'CheckBox
    If chkCancel.Value = 0 Then subAnd ("A.cancelFlag=0")
    
  '�L��Ǹ�
    Select Case True
           Case optPrtSNoYes.Value '���L��Ǹ��~�L
                subAnd ("A.PrtSNo IS Not Null")
                strPrtSNo = "�w��"
           Case optPrtSNoNo.Value  '�L�L��Ǹ��~�L
                subAnd ("A.PrtSNo IS Null")
                strPrtSNo = "����"
           Case optPrtSNoAll.Value '���ަ��L�L��Ǹ��ҦL
                strPrtSNo = "����"
    End Select
    
      '�]�w����榡
    If OptPrint2.Value = True Or OptPrint3.Value = True Then
       Select Case True
          Case OptOldClctName.Value = True
'                  strFormula = "OldClctName='�즬�O�H��';CName='0'"
'                  strFormula = "CName='0'"
                  strFormula = "Type=0"
          Case OptClctName.Value = True
'                  strFormula = "ClctName='���O�H��';CName=1"
                  strFormula = "Type=1"
        End Select
'       Select Case True
'           Case OptOldAmt.Value = True
'                   strFormula = strFormula & ";OldAmt='���������B'"
'           Case OptShouldAmt.Value = True
'                   strFormula = strFormula & ";ShouldAmt='�������B'"
'        End Select
     End If
'
'    Else
'       strSumCust = "��D�N�X"
'       strFormula = "TitleName='��D'"
'       Select Case True
'              Case optGeneral.Value
'                   strType = "�@���"
'              Case optMdu.Value
'                   strType = "�j�Ӥ�"
'              Case optAll.Value
'                   strType = "����"
'       End Select
'   End If
                                               
    strChooseString = "�������: " & subSpace(gdaShouldDate1.GetOriginalValue) & "~" & subSpace(gdaShouldDate2.GetOriginalValue) & ";" & _
                      "�ꦬ���: " & subSpace(gdaRealDate1.GetOriginalValue) & "~" & subSpace(gdaRealDate2.GetOriginalValue) & ";" & _
                      "�X����: " & subSpace(GiCreateTime1.GetOriginalValue) & "~" & subSpace(GiCreateTime2.GetOriginalValue) & ";" & _
                      "���q�O  : " & subSpace(gilCompCode.GetDescription) & ";" & _
                      "�A�����O: " & subSpace(gilServiceType.GetDescription) & ";" & _
                      "�@�o��ڬO�_�@�ֲέp:" & IIf(chkCancel.Value = 1, "�O", "�_") & ";" & _
                      "���O����: " & subSpace(gimCitemCode.GetDispStr) & ";" & _
                      "���O�H��: " & subSpace(gimClctEn.GetDispStr) & ";" & _
                      "�즬�O�H��: " & subSpace(gimOldClctEn.GetDispStr) & ";" & _
                      "���O�覡: " & subSpace(gimCMCode.GetDispStr) & ";" & _
                      "������O: " & subSpace(gimBillType.GetDispStr) & ";" & _
                      "�u����]:" & subSpace(gimSTCode.GetDispStr) & ";" & _
                      "�L��Ǹ�:" & strPrtSNo
   Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
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

Private Sub cmdExit_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "���O���إN�X", "���O���ئW��")
    Call SetgiMulti(gimClctEn, "EmpNo", "EmpName", "CM003", "���O���N��", "���O���W��")
    Call SetgiMulti(gimOldClctEn, "EmpNo", "EmpName", "CM003", "�즬�O���N��", "�즬�O���W��")
    Call SetgiMultiAddItem(gimBillType, "B,I,P,M,T", "���O��,�˾���,�������,���׳�,�{�ɦ��O��", "��ڥN��", "��ڦW��")
    gimBillType.SetDispStr "���O��"
    gimBillType.SetQueryCode "B"
    Call SetgiMulti(gimSTCode, "CodeNo", "Description", "CD016", "�u����]�N�X", "�u����]�W��")
    '2006/03/13 ���D��2193 �e������W�["���O�覡" -----Edit by Crystal for Debby
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "���O�覡�N��", "���O�覡�W��")
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
  
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5145A)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  gcnGi.Execute "Drop View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5145A
End Sub

Private Sub gdaRealDate1_GotFocus()
  On Error GoTo ChkErr
    If gdaRealDate1.GetValue = "" Then gdaRealDate1.SetValue (RightDate)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaRealDate1_GotFocus")
End Sub

Private Sub gdaRealDate2_GotFocus()
  On Error GoTo ChkErr
    If gdaRealDate1.GetValue = "" Or gdaRealDate2.GetValue = "" Then gdaRealDate2.SetValue gdaRealDate1.GetValue
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaRealDate2_GotFocus")
End Sub

Private Sub gdaRealDate2_Validate(Cancel As Boolean)
  On Error Resume Next
    Cancel = ChkDate2(gdaRealDate1, gdaRealDate2)
End Sub

Private Sub gdaShouldDate1_GotFocus()
  On Error GoTo ChkErr
    If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetValue (RightDate)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdaShouldDate1_GotFocus")
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

Private Sub GiCreateTime1_GotFocus()
  On Error Resume Next
    If GiCreateTime1.GetValue = "" Then GiCreateTime1.SetValue (RightDate)
End Sub

Private Sub GiCreateTime2_GotFocus()
  On Error Resume Next
    If GiCreateTime1.GetValue = "" Or GiCreateTime2.GetValue = "" Then GiCreateTime2.SetValue (GiCreateTime1.GetValue)
End Sub

Private Sub GiCreateTime2_Validate(Cancel As Boolean)
  On Error Resume Next
    Cancel = ChkDate2(GiCreateTime1, GiCreateTime2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    GiMultiFilter gimClctEn, , gilCompCode.GetCodeNo
    GiMultiFilter gimOldClctEn, , gilCompCode.GetCodeNo
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
    GiMultiFilter gimSTCode, gilServiceType.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilServiceType_Change")
End Sub

Private Function subCreateView() As Boolean '�ؤ@�� SO033,SO034 View
  On Error Resume Next
    gcnGi.Execute "Drop View " & strViewName
  On Error GoTo ChkErr
  Dim strViewSql As String
  Dim rsSo043Para As New ADODB.Recordset
  Dim strParaSQL As String
  Dim lngPara As Long
  Dim strJoinBillNo As String
  '���D��2789 �����XSO043�O�_�ϥδC��h�b��B�z
  strParaSQL = "Select Para24 From So043 Where CompCode=" & gilCompCode.GetCodeNo & " And " _
                & "ServiceType='" & gilServiceType.GetCodeNo & "'"
  If Not GetRS(rsSo043Para, strParaSQL) Then Exit Function
  If Not rsSo043Para.EOF Then
    lngPara = rsSo043Para(0)
    If lngPara = 0 Then
        strBillOrMead = "BillNo"
    Else
        strBillOrMead = "MediaBillNo"
    End If
  Else
    Exit Function
  End If
  Set rsSo043Para = Nothing
  '���D��2789 �ھڰѼƦ�渹�δC�^
  strJoinBillNo = "A." & strBillOrMead
      strViewName = GetTmpViewName
      If OptPrint1.Value = True Then
             strViewSql = "Create View " & strViewName & " as (" & _
                             "SELECT A.ClctEn,A.ClctName,A.OldClctEn,A.OldClctName,A.Custid," & strJoinBillNo & ",A.ShouldAmt,A.OldAmt,A.RealAmt,A.RealDate,A.STCode,A.CancelFlag " & _
                                 "From SO033 A Where " & strChoose & _
                          " Union All " & _
                             "SELECT A.ClctEn,A.ClctName,A.OldClctEn,A.OldClctName,A.Custid," & strJoinBillNo & ",A.ShouldAmt,A.OldAmt,A.RealAmt,A.RealDate,A.STCode,A.CancelFlag " & _
                                 "From SO034 A Where " & strChoose & IIf(chkCancel.Value = 0, ")", "")
            ' 20040804 Editby Crystal for Debby ���D��1005 �[��SO035
            If chkCancel.Value = 1 Then
            strViewSql = strViewSql & " Union All " & _
                         "SELECT A.ClctEn,A.ClctName,A.OldClctEn,A.OldClctName,A.Custid," & strJoinBillNo & ",A.ShouldAmt,A.OldAmt,A.RealAmt,A.RealDate,A.STCode,A.CancelFlag " & _
                         "From SO035 A Where " & strChoose & ")"
            End If
     ElseIf OptPrint2.Value = True Then
            strViewSql = "Create View " & strViewName & " as (" & _
                             "SELECT A.ClctEn,A.ClctName,A.OldClctEn,A.OldClctName,A.Custid," & strJoinBillNo & ",A.ShouldAmt,A.OldAmt,A.RealAmt,A.RealDate,A.STCode,A.CancelFlag,A.OldPeriod,A.RealPeriod  " & _
                                 "From SO033 A Where " & strChoose & _
                          " Union All " & _
                             "SELECT A.ClctEn,A.ClctName,A.OldClctEn,A.OldClctName,A.Custid," & strJoinBillNo & ",A.ShouldAmt,A.OldAmt,A.RealAmt,A.RealDate,A.STCode,A.CancelFlag,A.OldPeriod,A.RealPeriod  " & _
                                 "From SO034 A Where " & strChoose & IIf(chkCancel.Value = 0, ")", "")
            ' 20040804 Editby Crystal for Debby ���D��1005 �[��SO035
            If chkCancel.Value = 1 Then
            strViewSql = strViewSql & " Union All " & _
                         "SELECT A.ClctEn,A.ClctName,A.OldClctEn,A.OldClctName,A.Custid," & strJoinBillNo & ",A.ShouldAmt,A.OldAmt,A.RealAmt,A.RealDate,A.STCode,A.CancelFlag,A.OldPeriod,A.RealPeriod  " & _
                         "From SO035 A Where " & strChoose & ")"
            End If
     ElseIf OptPrint3.Value = True Then
             strViewSql = "Create View " & strViewName & " as (" & _
                             "SELECT A.ClctEn,A.ClctName,A.OldClctEn,A.OldClctName,A.Custid," & strJoinBillNo & ",A.ShouldAmt,A.OldAmt,A.RealAmt,A.RealDate,A.STCode,A.CancelFlag " & _
                                 "From SO033 A Where " & strChoose & _
                          " Union All " & _
                             "SELECT A.ClctEn,A.ClctName,A.OldClctEn,A.OldClctName,A.Custid," & strJoinBillNo & ",A.ShouldAmt,A.OldAmt,A.RealAmt,A.RealDate,A.STCode,A.CancelFlag " & _
                                 "From SO034 A Where " & strChoose & IIf(chkCancel.Value = 0, ")", "")
            ' 20040804 Editby Crystal for Debby ���D��1005 �[��SO035
            If chkCancel.Value = 1 Then
            strViewSql = strViewSql & " Union All " & _
                         "SELECT A.ClctEn,A.ClctName,A.OldClctEn,A.OldClctName,A.Custid," & strJoinBillNo & ",A.ShouldAmt,A.OldAmt,A.RealAmt,A.RealDate,A.STCode,A.CancelFlag " & _
                         "From SO035 A Where " & strChoose & ")"
            End If
     End If
      gcnGi.Execute strViewSql
      SendSQL strViewSql
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function

Private Sub CreateTable()
  On Error Resume Next
    cnn.Execute "Drop Table SO5145A1"
    cnn.Execute "Drop Table SO5145A2"
  On Error GoTo ChkErr
    If OptPrint1.Value = True Then
           
        '����榡1
                cnn.Execute "Create Table SO5145A1 (ClctEn text(20),ClctName Text(40),OldClctEn text(20),OldClctName Text(40),ShouldCust Long,ShouldBillNo Long,ShouldAmt Long,RealCust Long,RealBillNo Long,RealAmt Long)"
                 '����
                 Call subInsertMDB("Select ClctEn,ClctName,OldClctEn,OldClctName,Count(Distinct Custid) as CustCount,Count(Distinct " & strBillOrMead & ") as BillCount,Sum(ShouldAmt) as SumAmt " & _
                                       "From " & strViewName & " Group by ClctEn,ClctName,OldClctEn,OldClctName", "Should")
                 '�ꦬ
                 Call subInsertMDB("Select ClctEn,ClctName,OldClctEn,OldClctName,Count(Distinct Custid) as CustCount,Count(Distinct " & strBillOrMead & ") as BillCount,Sum(RealAmt) as SumAmt " & _
                                       "From " & strViewName & _
                                       " Where " & IIf(strChooseReal = "", "", strChooseReal & " And ") & "(RealAmt<>0 " & IIf(strChooseST = "", "", " Or " & strChooseST) & ") Group by ClctEn,ClctName,OldClctEn,OldClctName", "Real")
        '����榡2
    ElseIf OptPrint2.Value = True Then
            cnn.Execute "Create Table SO5145A1 (ClctEn text(20),ClctName Text(40),OldClctEn text(20),OldClctName Text(40),ShouldAmt Long,OldAmt Long,RealPeriod Long,OldPeriod Long,RealAmt Long,AmtCount Long,RealAmtCount Long)"
               '����
                Call subInsertMDB("Select " & strClct & "," & strClct2 & "," & strAmt & ",RealAmt,OldPeriod,RealPeriod,Count(Distinct " & strAmt & " ) as AmtCount,Count(Distinct RealAmt) as RealAmtCount" & _
                                  " From " & strViewName & " Group by " & strClct & "," & strClct2 & "," & strAmt & ", RealAmt, OldPeriod, RealPeriod", "Should")
                                  
               '�ꦬ
                Call subInsertMDB("Select " & strClct & "," & strClct2 & "," & strAmt & ",RealAmt,OldPeriod,RealPeriod,Count(Distinct " & strAmt & " ) as AmtCount,Count(Distinct RealAmt) as RealAmtCount" & _
                                  " From " & strViewName & _
                                  " Where " & IIf(strChooseReal = "", "", strChooseReal & " And ") & "(RealAmt<>0 " & IIf(strChooseST = "", "", " Or " & strChooseST) & ") Group by " & strClct & "," & strClct2 & "," & strAmt & ", RealAmt, OldPeriod, RealPeriod", "Real")
    ElseIf OptPrint3.Value = True Then
        '����榡3
                cnn.Execute "Create Table SO5145A1 (ClctEn text(20),ClctName Text(40),ShouldCust Long,ShouldBillNo Long,ShouldAmt Long,RealCust Long,RealBillNo Long,RealAmt Long)"
                 '����
                 Call subInsertMDB("Select " & strClct & " ClctEn," & strClct2 & " ClctName,Count(Distinct Custid) as CustCount,Count(Distinct " & strBillOrMead & ") as BillCount,Sum(" & strAmt & ") as SumAmt " & _
                                       "From " & strViewName & " Group by " & strClct & "," & strClct2, "Should")
                 '�ꦬ
                 Call subInsertMDB("Select " & strClct & " ClctEn," & strClct2 & " ClctName" & ",Count(Distinct Custid) as CustCount,Count(Distinct " & strBillOrMead & ") as BillCount,Sum(RealAmt) as SumAmt " & _
                                       "From " & strViewName & _
                                       " Where " & IIf(strChooseReal = "", "", strChooseReal & " And ") & "(RealAmt<>0 " & IIf(strChooseST = "", "", " Or " & strChooseST) & ") Group by " & strClct & "," & strClct2, "Real")
          
    End If
    SumMDB
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "CreateTable")
End Sub

Private Sub subInsertMDB(strsql As String, strType As String)
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
  
    Set rsTmp = gcnGi.Execute(strsql)
    SendSQL strsql
    cnn.BeginTrans
    While Not rsTmp.EOF

         If OptPrint1.Value = True Then
               cnn.Execute "Insert Into SO5145A1 (ClctEn,ClctName,OldClctEn,OldClctName," & strType & "Cust," & strType & "BillNo," & strType & "Amt" & _
                             ") Values(" & _
                             GetNullString(rsTmp("ClctEn")) & "," & _
                             GetNullString(rsTmp("ClctName")) & "," & _
                             GetNullString(rsTmp("OldClctEn")) & "," & _
                             GetNullString(rsTmp("OldClctName")) & "," & _
                             GetNullString(rsTmp("CustCount")) & "," & _
                             GetNullString(rsTmp("BillCount")) & "," & _
                             GetNullString(rsTmp("SumAmt")) & ")"
        ElseIf OptPrint2.Value = True Then
                cnn.Execute "insert into SO5145A1(" & strClct & "," & strClct2 & "," & strAmt & ",RealAmt,OldPeriod,RealPeriod,AmtCount,RealAmtCount" & _
                                 ") Values (" & _
                             GetNullString(rsTmp(strClct)) & "," & _
                             GetNullString(rsTmp(strClct2)) & "," & _
                             GetNullString(rsTmp(strAmt), giLongV) & "," & _
                             GetNullString(rsTmp("RealAmt"), giLongV) & "," & _
                             GetNullString(rsTmp("OldPeriod"), giLongV) & "," & _
                             GetNullString(rsTmp("RealPeriod"), giLongV) & "," & _
                             GetNullString(rsTmp("AmtCount"), giLongV) & "," & _
                             GetNullString(rsTmp("RealAmtCount"), giLongV) & ")"
         ElseIf OptPrint3.Value = True Then
               cnn.Execute "Insert Into SO5145A1 (ClctEn,ClctName ," & strType & "Cust," & strType & "BillNo," & strType & "Amt" & _
                             ") Values(" & _
                             GetNullString(rsTmp("ClctEn")) & "," & _
                             GetNullString(rsTmp("ClctName")) & "," & _
                             GetNullString(rsTmp("CustCount")) & "," & _
                             GetNullString(rsTmp("BillCount")) & "," & _
                             GetNullString(rsTmp("SumAmt")) & ")"
        End If
        rsTmp.MoveNext
    Wend
    cnn.CommitTrans
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
 cnn.RollbackTrans
  Call ErrSub(Me.Name, "subInsertMDB")
End Sub

Private Sub SumMDB()
  On Error GoTo ChkErr
  Dim strsql As String
    cnn.BeginTrans
          If OptPrint1.Value = True Then
             strsql = "Select ClctEn,ClctName,OldClctEn,OldClctName," & _
                   "iif(isnull(sum(ShouldCust)),0,sum(ShouldCust)) as ShouldCust,iif(isnull(sum(ShouldBillNo)) ,0,sum(ShouldBillNo)) as ShouldBillNo,iif(isnull(sum(ShouldAmt)),0,sum(ShouldAmt)) as ShouldAmt," & _
                   "iif(isnull(sum(RealCust)),0,sum(RealCust)) as RealCust,iif(isnull(sum(RealBillNo)),0,sum(RealBillNo)) as RealBillNo,iif(isnull(sum(RealAmt)),0,sum(RealAmt)) as RealAmt," & _
                   "iif(isnull(sum(RealBillNo)),0,sum(RealBillNo))/iif(isnull(sum(ShouldBillNo)) ,0,sum(ShouldBillNo))* 100 as FinPercentage " & _
                 " Into SO5145A2 From SO5145A1 Group by ClctEn,ClctName,OldClctEn,OldClctName "
           ElseIf OptPrint2.Value = True Then
              strsql = "Select " & strClct & "," & strClct2 & "," & strAmt & ",RealAmt,oldperiod,realperiod" & _
                         " Into SO5145A2 From SO5145A1 Group by " & strClct & "," & strClct2 & "," & strAmt & ",RealAmt,oldperiod,realperiod "
          ElseIf OptPrint3.Value = True Then
             strsql = "Select ClctEn,ClctName," & _
                   "iif(isnull(sum(ShouldCust)),0,sum(ShouldCust)) as ShouldCustA,iif(isnull(sum(ShouldBillNo)) ,0,sum(ShouldBillNo)) as ShouldBillNoA,iif(isnull(sum(ShouldAmt)),0,sum(ShouldAmt)) as ShouldAmtA," & _
                   "iif(isnull(sum(RealCust)),0,sum(RealCust)) as RealCustA,iif(isnull(sum(RealBillNo)),0,sum(RealBillNo)) as RealBillNoA,iif(isnull(sum(RealAmt)),0,sum(RealAmt)) as RealAmtA," & _
                   "iif(isnull(sum(RealBillNo)),0,sum(RealBillNo))/iif(isnull(sum(ShouldBillNo)) ,0,sum(ShouldBillNo))* 100 as FinPercentage " & _
                 " Into SO5145A2 From SO5145A1 Group by ClctEn,ClctName "
          End If
        cnn.Execute strsql
        SendSQL strsql, True
    cnn.CommitTrans
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "SumMDB")
End Sub
'����榡�@
Private Sub OptPrint1_Click()
   If OptPrint1.Value = True Then
        Frame4.Enabled = False
        Frame5.Enabled = False
        OptOldClctName.Value = False
        OptOldAmt.Value = False
        OptClctName.Value = False
        OptShouldAmt.Value = False
        OptOldClctName.Enabled = False
        OptOldAmt.Enabled = False
        OptClctName.Enabled = False
        OptShouldAmt.Enabled = False
    End If
End Sub
'����榡�G
Private Sub OptPrint2_Click()
  If OptPrint2.Value = True Then
        Frame4.Enabled = True
        Frame5.Enabled = True
        OptOldClctName.Value = True
        OptOldAmt.Value = True
        OptOldClctName.Enabled = True
        OptOldAmt.Enabled = True
        OptClctName.Enabled = True
        OptShouldAmt.Enabled = True
  End If
End Sub

'����榡�T
Private Sub OptPrint3_Click()
  If OptPrint3.Value = True Then
        Frame4.Enabled = True
        Frame5.Enabled = True
        OptOldClctName.Value = True
        OptOldAmt.Value = True
        OptOldClctName.Enabled = True
        OptOldAmt.Enabled = True
        OptClctName.Enabled = True
        OptShouldAmt.Enabled = True
  End If
End Sub
