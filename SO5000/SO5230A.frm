VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO5230A 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�C�������P�ꦬ���`�� [SO5230A]"
   ClientHeight    =   6135
   ClientLeft      =   1950
   ClientTop       =   1530
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO5230A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame fraIVR 
      Height          =   585
      Left            =   7260
      TabIndex        =   42
      Top             =   5430
      Visible         =   0   'False
      Width           =   2655
      Begin VB.OptionButton optIVRAll 
         Caption         =   "����"
         Height          =   315
         Left            =   1740
         TabIndex        =   28
         Top             =   180
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optNoIVR 
         Caption         =   "�DIVR"
         Height          =   315
         Left            =   810
         TabIndex        =   27
         Top             =   180
         Width           =   885
      End
      Begin VB.OptionButton optIVR 
         Caption         =   "IVR"
         Height          =   315
         Left            =   180
         TabIndex        =   26
         Top             =   180
         Width           =   885
      End
   End
   Begin VB.Frame fraPrint 
      Height          =   555
      Left            =   4530
      TabIndex        =   41
      Top             =   5460
      Visible         =   0   'False
      Width           =   2685
      Begin VB.OptionButton optSimple 
         Caption         =   "�u��^��²��"
         Height          =   405
         Left            =   1020
         TabIndex        =   25
         Top             =   120
         Width           =   1905
      End
      Begin VB.OptionButton optDetail 
         Caption         =   "���Ӫ�"
         Height          =   405
         Left            =   60
         TabIndex        =   24
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox txtNote 
      Height          =   525
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  '�������b
      TabIndex        =   16
      Top             =   4770
      Width           =   2415
   End
   Begin CS_Multi.CSmulti gimStrtCode 
      Height          =   345
      Left            =   0
      TabIndex        =   14
      Top             =   3360
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   609
      ButtonCaption   =   "��  �D  �d  ��"
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "�W���έp���G"
      Height          =   525
      Left            =   1620
      TabIndex        =   22
      Top             =   5490
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "����(&X)"
      Height          =   525
      Left            =   3180
      TabIndex        =   23
      Top             =   5490
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:�C�L"
      Default         =   -1  'True
      Height          =   525
      Left            =   120
      TabIndex        =   21
      Top             =   5490
      Width           =   1245
   End
   Begin VB.Frame fraPage 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�����覡"
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   2490
      TabIndex        =   34
      Top             =   4560
      Width           =   5025
      Begin VB.OptionButton optStrtCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��D�s��"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3750
         TabIndex        =   20
         Top             =   330
         Width           =   1215
      End
      Begin VB.OptionButton optClctEn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "���O�H��"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2400
         TabIndex        =   19
         Top             =   330
         Width           =   1215
      End
      Begin VB.OptionButton optAreaCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��F��"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1290
         TabIndex        =   18
         Top             =   330
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optServCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�A�Ȱ�"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   330
         Width           =   885
      End
   End
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   345
      Left            =   0
      TabIndex        =   13
      Top             =   2970
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   609
      ButtonCaption   =   "��     �F     ��"
   End
   Begin Gi_Multi.GiMulti gimSTCode 
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   2190
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   609
      ButtonCaption   =   "�u  ��  ��  �]"
   End
   Begin Gi_Multi.GiMulti gimCMCode 
      Height          =   345
      Left            =   0
      TabIndex        =   9
      Top             =   1410
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   609
      ButtonCaption   =   "��  �O  ��  ��"
   End
   Begin Gi_Date.GiDate gdaShouldDate2 
      Height          =   345
      Left            =   2310
      TabIndex        =   1
      Top             =   90
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
      Left            =   960
      TabIndex        =   0
      Top             =   90
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
      Left            =   960
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
   Begin Gi_Date.GiDate gdaRealDate2 
      Height          =   345
      Left            =   2310
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
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   5670
      TabIndex        =   7
      Top             =   540
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
      Left            =   5670
      TabIndex        =   8
      Top             =   990
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
   Begin prjGiList.GiList gilClctEn 
      Height          =   315
      Left            =   5670
      TabIndex        =   6
      Top             =   90
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
      TabIndex        =   15
      Top             =   3750
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   609
      ButtonCaption   =   "��  ��  ��  ��"
      DataType        =   2
      ColumnOrder     =   -1  'True
   End
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   345
      Left            =   0
      TabIndex        =   10
      Top             =   1800
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   609
      ButtonCaption   =   "��  �O  ��  ��"
   End
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   345
      Left            =   0
      TabIndex        =   12
      Top             =   2580
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   609
      ButtonCaption   =   "�A     ��     ��"
   End
   Begin Gi_Multi.GiMulti gimBillType 
      Height          =   375
      Left            =   0
      TabIndex        =   38
      Top             =   4110
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   661
      ButtonCaption   =   "��  ��  ��  �O"
      DataType        =   2
      DIY             =   -1  'True
   End
   Begin Gi_Time.GiTime gdtPRTime1 
      Height          =   315
      Left            =   930
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
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
   Begin Gi_Time.GiTime gdtPRTime2 
      Height          =   315
      Left            =   2910
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
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
   Begin VB.Label lblPRT 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2760
      TabIndex        =   40
      Top             =   1050
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label lblResvTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�w�����"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   60
      TabIndex        =   39
      Top             =   1020
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
      Caption         =   "���O�Ƶ�"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   30
      TabIndex        =   37
      Top             =   4530
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
      Caption         =   "��  �q  �O"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4830
      TabIndex        =   36
      Top             =   630
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '�z��
      Caption         =   "�A�����O"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4830
      TabIndex        =   35
      Top             =   1050
      Width           =   780
   End
   Begin VB.Label lblClctEn 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "���O�H��"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4800
      TabIndex        =   33
      Top             =   180
      Width           =   780
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2130
      TabIndex        =   32
      Top             =   630
      Width           =   105
   End
   Begin VB.Label lblRealDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�ꦬ���"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   60
      TabIndex        =   31
      Top             =   630
      Width           =   780
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2130
      TabIndex        =   30
      Top             =   180
      Width           =   105
   End
   Begin VB.Label lblShouldDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�������"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   90
      TabIndex        =   29
      Top             =   180
      Width           =   780
   End
End
Attribute VB_Name = "frmSO5230A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�ϥ�Table: SO033 or SO034 or SO035 A,SO001,SO014
Option Explicit
Private strSO00XWhere As String
Private blnUseIVR As Boolean
Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    '�u��^��²�� [SO5230A2]
    Screen.MousePointer = vbHourglass
    If blnUseIVR And optSimple.Value Then
        Call PreviousRpt(GetPrinterName(5), RptName("SO5230", "2"), "�u��^��²�� [SO5230A2]")
    Else
        Call PreviousRpt(GetPrinterName(5), RptName("SO5230"), "�C�������P�ꦬ���`�� [SO5230A]")
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
          Call subPrint
        Me.Enabled = True
      Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: strErrFile = "���q�O": GoTo 66
    
    '#5000 �P�_�O�_�ϥ�IVR�^��\��,�p�G�O�~��j���w������n���� By Kin 2009/04/07
    '#5039 �N���\�ಾ��SO52J0
'    If blnUseIVR Then
'        If optSimple.Value Then
'            If (gdtPRTime1.GetValue = "") And (gdtPRTime2.GetValue = "") Then
'                gdtPRTime1.SetFocus
'                strErrFile = "�w���ɶ�"
'                GoTo 66
'            End If
'        End If
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
    Dim strpagetype As String
    Dim arrOrder() As String
    Dim varOrder As Variant
    Dim intSort As Integer
    strSO00XWhere = Empty

    strChoose = " WHERE A.CUSTID=SO001.CUSTID And A.CompCode = SO001.CompCode AND A.ADDRNO=SO014.ADDRNO And A.CompCode = SO014.CompCode "
    '#5000 �P�_�O�_��²��,�p�G²�����U�������� By Kin 2009/04/02
    '#5239 �NIVR�\�ಾ��SO52J0 By Kin 2009/04/23
'    If blnUseIVR Then
'        If optDetail.Value Then
'            strChoose = strChoose & " And A.ShouldAmt<>A.RealAmt And A.RealDate is not null And A.cancelFlag=0 "
'        Else
'            strChoose = " WHERE A.CUSTID=SO001.CUSTID And A.CompCode = SO001.CompCode AND A.ADDRNO=SO014.ADDRNO And A.CompCode = SO014.CompCode "
'        End If
'    Else
        strChoose = strChoose & " And A.ShouldAmt<>A.RealAmt And A.RealDate is not null And A.cancelFlag=0 "
'    End If
    strChooseString = ""
    strpagetype = ""
    
  '���
    If gdaRealDate1.GetValue <> "" Then Call subAnd("A.RealDate >= To_Date('" & gdaRealDate1.GetValue & "','YYYYMMDD')")
    If gdaRealDate2.GetValue <> "" Then Call subAnd("A.RealDate < To_Date('" & gdaRealDate2.GetValue & "','YYYYMMDD')+1")
    If gdaShouldDate1.GetValue <> "" Then Call subAnd("A.ShouldDate >= To_Date('" & gdaShouldDate1.GetValue & "','YYYYMMDD')")
    If gdaShouldDate2.GetValue <> "" Then Call subAnd("A.ShouldDate < To_Date('" & gdaShouldDate2.GetValue & "','YYYYMMDD')+1")
    
    '#5000 �W�[�P�_�w���ɶ� By Kin 2009/04/02
    '#5239 �NIVR�\�ಾ��SO52J0 By Kin 2009/04/23
'    If blnUseIVR Then
'        If gdtPRTime1.GetValue <> "" Then
'            Call subAnd2(strSO00XWhere, "D.ResvTime>=To_Date(" & gdtPRTime1.GetValue & ",'yyyymmddhh24mi')")
'        End If
'        If gdtPRTime2.GetValue <> "" Then
'            Call subAnd2(strSO00XWhere, "D.ResvTime<To_Date(" & gdtPRTime2.GetValue & ",'yyyymmddhh24mi')+1")
'        End If
'    End If
  'GiMulti
    
    If gimAreaCode.GetQryStr <> "" Then Call subAnd("A.AreaCode " & gimAreaCode.GetQryStr)
    If gimCitemCode.GetQryStr <> "" Then Call subAnd("A.CitemCode " & gimCitemCode.GetQryStr)
    If gimCMCode.GetQryStr <> "" Then Call subAnd("A.CMCode " & gimCMCode.GetQryStr)
    If gimSTCode.GetQryStr <> "" Then Call subAnd("A.STCode " & gimSTCode.GetQryStr)
    If gimServCode.GetQryStr <> "" Then Call subAnd("A.ServCode " & gimServCode.GetQryStr)
    If gimStrtCode.GetQryStr <> "" Then Call subAnd("A.StrtCode " & gimStrtCode.GetQryStr)
    '#5000 �W�[�L�o���� By Kin 2009/04/02
    ''#5239 �NIVR�\�ಾ��SO52J0 By Kin 2009/04/23
    If gimBillType.GetQryStr <> "" Then
        Call subAnd("SubStr(A.BillNo,7,1) " & gimBillType.GetQryStr)
'        If blnUseIVR And optSimple.Value Then
'            Call subAnd2(strSO00XWhere, " SubStr(D.SNO,7,1)" & gimBillType.GetQryStr)
'        End If
'    Else
'        If blnUseIVR And optSimple Then
'            Call subAnd("SubStr(A.BillNo,7,1) In('I','M','P')")
'        End If
    End If
    
  'GiList
    If gilClctEn.GetCodeNo <> "" Then Call subAnd("A.ClctEn ='" & gilClctEn.GetCodeNo & "'")
    
    '#5000 ��u��SO00X���q�O������ By Kin 2009/04/08
    If gilCompCode.GetCodeNo <> "" Then
        Call subAnd("A.CompCode =" & gilCompCode.GetCodeNo)
        'Call subAnd2(strSO00XWhere, "D.CompCode=" & gilCompCode.GetCodeNo)
    End If
    
    '#5000 �u��]�n��ServiceType By Kin 2009/04/08
    If gilServiceType.GetCodeNo <> "" Then
        Call subAnd("A.ServiceType ='" & gilServiceType.GetCodeNo & "'")
        '#5239 �NIVR�\�ಾ��SO52J0 By Kin 2009/04/23
'        If blnUseIVR And optSimple.Value Then
'            Call subAnd2(strSO00XWhere, " D.ServiceType ='" & gilServiceType.GetCodeNo & "'")
'        End If
    End If
    
    '#5000 �P�_�O�_�ϥ�²��,�M��A�P�_�u��O�_�ϥ�IVR�^�� By Kin 2009/04/07
    '#5239 �NIVR�\�ಾ��SO52J0 By Kin 2009/04/23
'    If blnUseIVR Then
'        If optSimple Then
'            If optIVR.Value Then Call subAnd2(strSO00XWhere, " D.IVRFLAG > 0 ")
'            If optNoIVR.Value Then Call subAnd2(strSO00XWhere, " D.IVRFLAG = 0 ")
'        End If
'    End If
    
    If txtNote <> "" Then Call subAnd("A.Note like '%" & txtNote & "%'")
    Select Case True
           Case optAreaCode.Value '��F��
                strGroupName = "GroupName={V.AreaCode};strGroupName={V.AreaName}"
                strpagetype = "��F��"
           Case optClctEn.Value '���O�H��
                strGroupName = "GroupName={V.ClctEn};strGroupName={V.ClctName}"
                strpagetype = "���O�H��"
           Case optServCode.Value '�A�Ȱ�
                strGroupName = "GroupName={V.ServCode};strGroupName={V.ServName}"
                strpagetype = "�A�Ȱ�"
           Case optStrtCode.Value '��D
                strGroupName = "GroupName={V.StrtCode};;strGroupName={V.StrtName}"
                strpagetype = "��D"
    End Select

  '�ƧǤ覡
    If gimOrder.GetColumnOrderCode <> "" Then
      arrOrder = Split(gimOrder.GetColumnOrderCode, ",")
      intSort = 0
      For Each varOrder In arrOrder
          Select Case arrOrder(intSort)
                 Case "A"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.CustId}"
                 Case "B"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.ClctEn}"
                 Case "C"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.BillNo}"
                 Case "D"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.RealDate}"
                 Case "E"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.AddrSort}"
                 Case "F"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.ClassCode}"
                 Case "G"
                      strGroupName = strGroupName & ";Sort" & intSort & "={V.Note2}"
          End Select
          intSort = intSort + 1
      Next
    End If
    strChooseString = "�������: " & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & _
                      "�ꦬ���: " & subSpace(gdaRealDate1.GetValue(True)) & "~" & subSpace(gdaRealDate2.GetValue(True)) & ";" & _
                      "���O�H��: " & subSpace(gilClctEn.GetDescription) & ";" & _
                      "���q�O  : " & subSpace(gilCompCode.GetDescription) & ";" & _
                      "�A�����O: " & subSpace(gilServiceType.GetDescription) & ";" & _
                      "���O�覡: " & subSpace(gimCMCode.GetDispStr) & ";" & _
                      "���O����: " & subSpace(gimCitemCode.GetDispStr) & ";" & _
                      "�A�Ȱϡ@: " & subSpace(gimServCode.GetDispStr) & ";" & _
                      "��F�ϡ@: " & subSpace(gimAreaCode.GetDispStr) & ";" & _
                      "��D�W��: " & subSpace(gimStrtCode.GetDispStr) & ";" & _
                      "���O�Ƶ�: " & subSpace(txtNote) & ";" & _
                      IIf(blnUseIVR, "", "�����覡: " & strpagetype & ";") & _
                      "�ƧǤ覡: " & subSpace(gimOrder.GetColumnOrderDspStr) & ";"
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub subPrint()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    '�إ�View
    Call subCreateView
  On Error GoTo ChkErr
    Set rsTmp = gcnGi.Execute("Select Count(*) as intCount From " & strViewName & " Where RowNum=1")
    If rsTmp("intCount") = 0 Then
        MsgNoRcd
        SendSQL , , True
    Else
       strSQL = "SELECT * FROM " & strViewName & " V"
       '#5000 �P�_�O�_�ϥ�IVR�^��,�h�����@�� By Kin 2009/04/07
       '#5239 �NIVR�\�ಾ��SO52J0 By Kin 2009/04/23
'        If optSimple And blnUseIVR Then
'            Call PrintRpt(GetPrinterName(5), "SO5230A2.rpt", , "�u��^��²�� [SO5230A2]", strsql, strChooseString, , True, , , , GiPaperLandscape)
'        Else
            Call PrintRpt(GetPrinterName(5), "SO5230A.rpt", , "�C�������P�ꦬ���`�� [SO5230A]", strSQL, strChooseString, , True, , , strGroupName, GiPaperLandscape)
'        End If
    End If
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subPrint")
End Sub

Private Sub Form_Initialize()
  On Error GoTo ChkErr
    'blnUseIVR = Val(GetRsValue("Select IVRCALL From " & GetOwner & "SO041", gcnGi)) = 1
    blnUseIVR = False
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Initialize")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGim
    Call subGil
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
    If Not blnUseIVR Then
        fraIVR.Visible = False
        gdtPRTime1.Visible = False
        gdtPRTime2.Visible = False
        fraPrint.Visible = False
        lblResvTime.Visible = False
        lblPRT.Visible = False
    Else
        If optDetail.Value Then
            gdtPRTime1.Enabled = False
            gdtPRTime2.Enabled = False
            fraIVR.Enabled = False
        End If
    End If
    
'    If Alfa Then
'        gdtPRTime1.SetValue "200901010000"
'        gdtPRTime2.SetValue "200901010000"
'        optSimple.Value = True
'    End If
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
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "��F�ϥN�X", "��F�ϦW��")
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "���O���إN�X", "���O���ئW��")
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "���O�覡�N�X", "���O�覡�W��")
    Call SetgiMulti(gimSTCode, "CodeNo", "Description", "CD016", "�u����]�N�X", "�u����]�W��")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "�A�ȰϥN�X", "�A�ȰϦW��")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "��D�s��", "��D�W��")
    
    Call SetgiMultiAddItem(gimOrder, "A,B,C,D,E,F,G", "�Ȥ�s��,���O���N��,���O�渹,���O���,�a�},�Ȥ����O,���O�Ƶ�")
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    'Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
    Call SetgiList(gilClctEn, "EmpNo", "EmpName", "CM003")
    Call SetgiMultiAddItem(gimBillType, "B,I,P,M,T", "���O��,�˾���,�������,���׳�,�{�ɦ��O��")
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub
  
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5230A)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  gcnGi.Execute "Droup View " & strViewName
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5230A
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
    If gdaRealDate1.GetValue = "" Or gdaRealDate2.GetValue = "" Then gdaRealDate2.SetValue (gdaRealDate1.GetValue)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaRealDate2_GotFocus")
End Sub

Private Sub gdaRealDate2_Validate(Cancel As Boolean)
  On Error Resume Next
    If gdaRealDate2.GetValue < gdaRealDate1.GetValue Then MsgBox "�I���������j��_�l���": gdaShouldDate2.SetFocus
End Sub

Private Sub gdaShouldDate1_GotFocus()
  On Error GoTo ChkErr
    If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetValue (RightDate)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaShouldDate1_GotFocus")
End Sub

Private Sub gdaShouldDate2_GotFocus()
  On Error Resume Next
    If gdaShouldDate1.GetValue = "" Or gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue (gdaShouldDate1.GetValue)
   Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaShouldDate2_GotFocus")
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaShouldDate1, gdaShouldDate2)
End Sub

Private Function subCreateView() As Boolean
  Dim strView As String
  Dim i, j As Integer
  Dim strQryCnt As String
  Dim strWhere As String
    On Error Resume Next
      gcnGi.Execute "Drop View " & strViewName
    On Error GoTo ChkErr
    
    strViewName = GetTmpViewName
    strQryCnt = Empty
    '#5000 �W�[²��(�n��SO007,SO008,SO009) By Kin 2009/04/02
    '#5239 �NIVR�\�ಾ��SO52J0 By Kin 2009/04/23
    If (optDetail.Value) Or (Not blnUseIVR) Then
        strView = "Create View " & strViewName & " as (" & _
                      "SELECT A.CUSTID,A.BILLNO,A.CITEMNAME,A.REALSTARTDATE,A.REALSTOPDATE,A.SHOULDDATE,A.SHOULDAMT,A.REALDATE,A.REALAMT,A.UCNAME,SO001.CUSTNAME,SO001.TEL1,SO014.ADDRESS,A.AreaCode,SO014.AreaName,A.ClctEn,A.ClctName,A.ServCode,SO014.ServName,A.StrtCode,SO014.StrtName,A.AddrNo,A.ClassCode,SO014.AddrSort,A.Note,SubStrB(A.Note,1,255) Note2,STCode,STName " & _
                          "From SO001,SO014,SO034 A " & strChoose & _
                  " UNION All " & _
                      "SELECT A.CUSTID,A.BILLNO,A.CITEMNAME,A.REALSTARTDATE,A.REALSTOPDATE,A.SHOULDDATE,A.SHOULDAMT,A.REALDATE,A.REALAMT,A.UCNAME,SO001.CUSTNAME,SO001.TEL1,SO014.ADDRESS,A.AreaCode,SO014.AreaName,A.ClctEn,A.ClctName,A.ServCode,SO014.ServName,A.StrtCode,SO014.StrtName,A.AddrNo,A.ClassCode,SO014.AddrSort,A.Note,SubStrB(A.Note,1,255) Note2,STCode,STName " & _
                          "From SO001,SO014,SO033 A " & strChoose & _
                  " UNION All " & _
                      "SELECT A.CUSTID,A.BILLNO,A.CITEMNAME,A.REALSTARTDATE,A.REALSTOPDATE,A.SHOULDDATE,A.SHOULDAMT,A.REALDATE,A.REALAMT,A.UCNAME,SO001.CUSTNAME,SO001.TEL1,SO014.ADDRESS,A.AreaCode,SO014.AreaName,A.ClctEn,A.ClctName,A.ServCode,SO014.ServName,A.StrtCode,SO014.StrtName,A.AddrNo,A.ClassCode,SO014.AddrSort,A.Note,SubStrB(A.Note,1,255) Note2,STCode,STName " & _
                          "From SO001,SO014,SO035 A " & strChoose & ")"
'    Else
'        For i = 0 To 2
'            For j = 0 To 2
'                If strQryCnt <> Empty Then strQryCnt = strQryCnt & " Union All "
'                strQryCnt = strQryCnt & " Select  D.SNO FSNO,Nvl(Sum(A.ShouldAmt),0) FAMT,NULL NSNO,0 NAMT,NULL ESNO,0 EAMT,D.WORKEREN1,D.WORKERNAME1 " & _
'                    " From ( Select A.* From SO001,SO014,SO03" & i + 3 & " A " & strChoose & _
'                    " ) A,SO00" & j + 7 & " D Where A.BillNo(+)=D.SNO  And D.FinTime Is Not Null And " & strSO00XWhere & _
'                    " GROUP BY D.WORKEREN1,D.WORKERNAME1,D.SNO " & _
'                    " UNION  " & _
'                    " Select NULL FSNO,0 FAMT, D.SNO NSNO,NVL(SUM(A.SHOULDAMT),0) NAMT,NULL ESNO,0 EAMT,D.WORKEREN1,D.WORKERNAME1 " & _
'                    " From ( Select A.* From SO001,SO014,SO03" & i + 3 & " A " & strChoose & _
'                    " ) A,SO00" & j + 7 & " D Where A.BillNo(+)=D.SNO AND D.IVRDataMatch=2 And D.FinTime Is Not Null And " & strSO00XWhere & _
'                    " GROUP BY D.WORKEREN1,D.WORKERNAME1,D.SNO " & _
'                    " UNION  " & _
'                    " Select NULL FSNO,0 FAMT,NULL NSNO,0 NAMT,D.SNO ESNO,NVL(SUM(A.SHOULDAMT),0) EAMT,D.WORKEREN1,D.WORKERNAME1 " & _
'                    " From ( Select A.* From SO001,SO014,SO03" & i + 3 & " A " & strChoose & _
'                    " ) A,SO00" & j + 7 & " D Where A.BillNo(+)=D.SNO AND D.IVRDataMatch=1 And D.FinTime Is Not Null And " & strSO00XWhere & _
'                    " GROUP BY D.WORKEREN1,D.WORKERNAME1,D.SNO"
'
'            Next
'        Next i
'        strQryCnt = "Select COUNT(DISTINCT FSNO) CNT,SUM(FAMT) FAMT ," & _
'                    "COUNT(DISTINCT NSNO) NSNO,SUM(NAMT) NAMT ," & _
'                    " COUNT(DISTINCT ESNO) ESNO,SUM(EAMT) EAMT ,WORKEREN1,WORKERNAME1 FROM (" & _
'                strQryCnt & ") Group By WORKEREN1,WORKERNAME1 "
'        strView = "Create View " & strViewName & " As(" & strQryCnt & ")"

    End If
    gcnGi.Execute strView
    
    SendSQL strView, True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function

Private Sub gdtPRTime2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdtPRTime1, gdtPRTime2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
    GiListFilter gilClctEn, , gilCompCode.GetCodeNo
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
    GiMultiFilter gimCMCode, gilServiceType.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilServiceType_Change")
End Sub

Private Sub optDetail_Click()
    On Error Resume Next
    gimBillType.Clear
    Call SetgiMultiAddItem(gimBillType, "B,I,P,M,T", "���O��,�˾���,�������,���׳�,�{�ɦ��O��")
    gimOrder.Enabled = True
    gdtPRTime1.SetValue ""
    gdtPRTime2.SetValue ""
    If blnUseIVR Then
        gdtPRTime1.Enabled = False
        gdtPRTime2.Enabled = False
        fraIVR.Enabled = False
    End If
    lblResvTime.ForeColor = &HFF0000
    fraPage.Enabled = True
End Sub

Private Sub optSimple_Click()
    On Error Resume Next
    '#5000 �ϥ�IVR�^��\�ण���\�Ŀ������O By Kin 2009/04/07
    If blnUseIVR Then
        gimBillType.Clear
        'SetgiMultiAddItem(gimBillType, "B,I,P,M,T", "���O��,�˾���,�������,���׳�,�{�ɦ��O��")
        Call SetgiMultiAddItem(gimBillType, "I,P,M", "�˾���,�������,���׳�")
        gimOrder.Clear
        gimOrder.Enabled = False
        gdtPRTime1.Enabled = True
        gdtPRTime2.Enabled = True
        fraIVR.Enabled = True
        lblResvTime.ForeColor = &HFF&
        fraPage.Enabled = False
    End If
End Sub
