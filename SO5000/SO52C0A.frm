VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.17#0"; "csMulti.ocx"
Begin VB.Form frmSO52C0A 
   BorderStyle     =   1  '單線固定
   Caption         =   "每月每日實收月費分佈表[SO52C0A]"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   Icon            =   "SO52C0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Frame4 
      Caption         =   "網路編號"
      Height          =   735
      Left            =   30
      TabIndex        =   36
      Top             =   1380
      Width           =   7665
      Begin VB.TextBox mskCircuitNo 
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   90
         TabIndex        =   7
         Top             =   240
         Width           =   4605
      End
      Begin VB.OptionButton optAll 
         Caption         =   "全部"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6660
         TabIndex        =   9
         Top             =   270
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton optCircuitNo 
         Caption         =   "只印空白網路編號"
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   4830
         TabIndex        =   8
         Top             =   240
         Width           =   1845
      End
   End
   Begin CS_Multi.CSmulti gimStrtCode 
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   4890
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   661
      ButtonCaption   =   "街  道  範  圍"
   End
   Begin VB.Frame fraPage 
      Caption         =   "小計依據"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   0
      TabIndex        =   28
      Top             =   5730
      Width           =   7755
      Begin VB.OptionButton optClctAreaCode 
         Caption         =   "收費區"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   2550
         TabIndex        =   21
         Top             =   330
         Width           =   915
      End
      Begin VB.OptionButton optNothing 
         Caption         =   "不分"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   6600
         TabIndex        =   24
         Top             =   330
         Width           =   915
      End
      Begin VB.OptionButton optServCode 
         Caption         =   "服務區"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   300
         TabIndex        =   19
         Top             =   330
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optAreaCode 
         Caption         =   "行政區"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1425
         TabIndex        =   20
         Top             =   330
         Width           =   915
      End
      Begin VB.OptionButton optClctEn 
         Caption         =   "收費人員"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   5145
         TabIndex        =   23
         Top             =   330
         Width           =   1245
      End
      Begin VB.OptionButton optStrtCode 
         Caption         =   "街道編號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   3690
         TabIndex        =   22
         Top             =   330
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1800
      TabIndex        =   26
      Top             =   6780
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6150
      TabIndex        =   27
      Top             =   6780
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      TabIndex        =   25
      Top             =   6780
      Width           =   1245
   End
   Begin Gi_Date.GiDate gdaShouldDate2 
      Height          =   345
      Left            =   2370
      TabIndex        =   3
      Top             =   510
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
   Begin Gi_Date.GiDate gdaShouldDate1 
      Height          =   345
      Left            =   930
      TabIndex        =   2
      Top             =   510
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
   Begin Gi_Multi.GiMulti gimAreaCode 
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   4500
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   661
      ButtonCaption   =   "行      政      區"
   End
   Begin Gi_Multi.GiMulti gimServCode 
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   4110
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   661
      ButtonCaption   =   "服      務      區"
   End
   Begin Gi_Multi.GiMulti gimStatusCode 
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   3720
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   661
      ButtonCaption   =   "客  戶  狀  態"
   End
   Begin Gi_Multi.GiMulti gimBillType 
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   2940
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   661
      ButtonCaption   =   "單  據  類  別"
      DataType        =   2
      DIY             =   -1  'True
      Exception       =   -1  'True
   End
   Begin prjGiList.GiList gilClctEn 
      Height          =   315
      Left            =   4680
      TabIndex        =   6
      Top             =   960
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
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
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2160
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   661
      ButtonCaption   =   "收  費  方  式"
   End
   Begin Gi_Date.GiDate gdaRealDate2 
      Height          =   345
      Left            =   2370
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
   Begin Gi_Date.GiDate gdaRealDate1 
      Height          =   345
      Left            =   930
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
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   4680
      TabIndex        =   5
      Top             =   510
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
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
      Height          =   315
      Left            =   4680
      TabIndex        =   4
      Top             =   90
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F5Corresponding =   -1  'True
   End
   Begin CS_Multi.CSmulti gimMduId 
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   5280
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   661
      ButtonCaption   =   "大  樓  編  號"
   End
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   2550
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   661
      ButtonCaption   =   "收  費  項  目"
   End
   Begin CS_Multi.CSmulti gimClassCode 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   3330
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   661
      ButtonCaption   =   "客  戶  類  別"
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3870
      TabIndex        =   35
      Top             =   150
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "服務類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3870
      TabIndex        =   34
      Top             =   570
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2160
      TabIndex        =   33
      Top             =   180
      Width           =   105
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "收費人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3870
      TabIndex        =   32
      Top             =   1020
      Width           =   780
   End
   Begin VB.Label lblShouldDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "應收日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   90
      TabIndex        =   31
      Top             =   600
      Width           =   780
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2160
      TabIndex        =   30
      Top             =   600
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "實收日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   210
      Width           =   780
   End
End
Attribute VB_Name = "frmSO52C0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table:SO033 or SO034 A,SO014(,SO001)(,SO002)
Option Explicit
'Dim strST As String
Dim strGroupField As String
Dim strFrom As String

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
   Screen.MousePointer = vbHourglass
     Call PreviousRpt(GetPrinterName(5), RptName("SO52C0"), "每月每日實收月費分佈表 [SO52C0A]")
   Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo ChkErr
  Dim lngCount As Long
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    Screen.MousePointer = vbHourglass
      cmdExit.SetFocus
      Me.Enabled = False
        ReadyGoPrint
        Call subChoose
        If Not subCreateView Then Exit Sub
        If Not subInsertMDB(lngCount) Then Exit Sub
        If lngCount = 0 Then
           MsgNoRcd
           SendSQL , , True
        Else
            Call subPrint
        End If
      Me.Enabled = True
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gdaRealDate1.GetValue = "" Then gdaRealDate1.SetFocus: strErrFile = "收費日期起始日": GoTo 66
    If gdaRealDate2.GetValue = "" Then gdaRealDate2.SetFocus: strErrFile = "收費日期截止日": GoTo 66
    If gdaRealDate2.GetValue(True) >= Format(DateAdd("m", 1, gdaRealDate1.Text), "YYYY/MM/DD") Then
        MsgBox "收費期間不能超過一個月", , "錯誤訊息"
        gdaRealDate2_GotFocus
        Exit Function
    End If
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
  Dim strpagetype As String '記錄小計依據
  Dim strTM As String '取出收費月份
  Dim blnSO001 As Boolean
  Dim blnSO002 As Boolean
  Dim blnSO014 As Boolean
    strFrom = "A"
    'strChoose = "A.AddrNo=SO014.AddrNo And A.cancelFlag=0 " & IIf(strST = "", "", "And nvl(A.STCode,-1)  in(" & strST & ",-1)")
    strChoose = "A.cancelFlag=0 And Nvl(A.STCode,-1) NOT IN (Select CodeNo From CD016 Where RefNo =1) And A.UCCode Is Null "
    strChooseString = ""
    strGroupName = ""
    
    '日期
      If gdaRealDate1.GetValue <> "" Then Call subAnd("A.RealDate >= To_Date('" & gdaRealDate1.GetValue & "','YYYYMMDD')")
      If gdaRealDate2.GetValue <> "" Then Call subAnd("A.RealDate < To_Date('" & gdaRealDate2.GetValue & "','YYYYMMDD')+1")
      If gdaShouldDate1.GetValue <> "" Then Call subAnd("A.ShouldDate >= To_Date('" & gdaShouldDate1.GetValue & "','YYYYMMDD')")
      If gdaShouldDate2.GetValue <> "" Then Call subAnd("A.ShouldDate < To_Date('" & gdaShouldDate2.GetValue & "','YYYYMMDD')+1")
      
    '網路編號
      If mskCircuitNo.Text <> "" Then
          Call subAnd("SO014.CircuitNo='" & mskCircuitNo.Text & "'")
          blnSO014 = True
      Else
          If optCircuitNo.Value Then Call subAnd("SO014.CircuitNo is null "): blnSO014 = True
      End If
      
    'GiList
      If gilClctEn.GetCodeNo <> "" Then Call subAnd("A.ClctEn='" & gilClctEn.GetCodeNo & "'")
      If gilCompCode.GetCodeNo <> "" Then Call subAnd("A.CompCode=" & gilCompCode.GetCodeNo)
      If gilServiceType.GetCodeNo <> "" Then Call subAnd("A.ServiceType='" & gilServiceType.GetCodeNo & "'")
    
    'GiMulti
      If gimCMCode.GetQryStr <> "" Then Call subAnd("A.CMCode " & gimCMCode.GetQryStr)
      If gimCitemCode.GetQryStr <> "" Then Call subAnd("A.CitemCode " & gimCitemCode.GetQryStr)
      If gimBillType.GetQryStr <> "" Then Call subAnd("SubStr(A.BillNo,7,1) " & gimBillType.GetQryStr)
      If gimClassCode.GetQryStr <> "" Then Call subAnd("SO001.ClassCode1 " & gimClassCode.GetQryStr): blnSO001 = True
      If gimStatusCode.GetQryStr <> "" Then Call subAnd("SO002.CustStatusCode " & gimStatusCode.GetQryStr): blnSO002 = True
      If gimServCode.GetQryStr <> "" Then Call subAnd("SO001.ServCode " & gimServCode.GetQryStr): blnSO001 = True
      If gimAreaCode.GetQryStr <> "" Then Call subAnd("SO014.AreaCode " & gimAreaCode.GetQryStr): blnSO014 = True
      If gimStrtCode.GetQryStr <> "" Then Call subAnd("SO014.StrtCode " & gimStrtCode.GetQryStr): blnSO014 = True

    '大樓編號
      If gimMduId.GetQryStr <> "" Then
          Call subAnd("A.MduId " & gimMduId.GetQryStr)
      Else
          If gimMduId.GetDispStr <> "" Then subAnd ("A.MduId is Not Null")
      End If
      
    '取出收費月份
      strTM = Format(gdaRealDate1.GetValue(True), "yyyy/mm")
      
    '串SO001
      If blnSO001 Then
          strFrom = strFrom & ",SO001"
          strChoose = "A.CustId=SO001.CustId And " & strChoose
      End If
    '串SO002
      If blnSO002 Then
          strFrom = strFrom & ",SO002"
          strChoose = "A.CustId=SO002.CustId And A.ServiceType=SO002.ServiceType And " & strChoose
      End If

    '串SO014
      If blnSO014 Then
          strFrom = strFrom & ",SO014"
          strChoose = "A.AddrNo=SO014.AddrNo And " & strChoose
      End If

      
    '小計依據
        Select Case True
               Case optServCode.Value
                    strFrom = strFrom & ",CD002"
                    strChoose = "A.ServCode=CD002.CodeNo(+) And " & strChoose
                    strGroupName = "ReportType=True;YYYYMM='" & strTM & "'"
                    strGroupField = "CD002.Description"
                    strpagetype = "服務區"
                   
               Case optAreaCode.Value
                    strFrom = strFrom & ",CD001"
                    strChoose = "A.AreaCode=CD001.CodeNo(+) And " & strChoose
                    strGroupName = "ReportType=True;YYYYMM='" & strTM & "'"
                    strGroupField = "CD001.Description"
                    strpagetype = "行政區"
                   
               Case optClctAreaCode.Value
                    strFrom = strFrom & ",CD040"
                    strChoose = "A.ClctAreaCode=CD040.CodeNo(+) And " & strChoose
                    strGroupName = "ReportType=True;YYYYMM='" & strTM & "'"
                    strGroupField = "CD040.Description"
                    strpagetype = "收費區"
                   
               Case optStrtCode.Value
                    strFrom = strFrom & ",CD017"
                    strChoose = "A.StrtCode=CD017.CodeNo(+) And " & strChoose
                    strGroupName = "ReportType=True;YYYYMM='" & strTM & "'"
                    strGroupField = "CD017.Description"
                    strpagetype = "街道編號"
                    
               Case optClctEn.Value
                    strGroupName = "ReportType=True;YYYYMM='" & strTM & "'"
                    strGroupField = "A.clctName"
                    strpagetype = "收費人員"
                    
               Case optNothing.Value
                    strGroupName = "ReportType=False;YYYYMM='" & strTM & "'"
                    strGroupField = "1"
                    strpagetype = "不分"
         End Select

      strChooseString = "收費日期: " & subSpace(gdaRealDate1.GetValue(True)) & "~" & subSpace(gdaRealDate2.GetValue(True)) & ";" & _
                        "應收日期: " & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & _
                        "收費人員:" & subSpace(gilClctEn.GetDescription) & ";" & _
                        "公司別　:" & subSpace(gilCompCode.GetDescription) & ";" & _
                        "服務類別: " & subSpace(gilServiceType.GetDescription) & ";" & _
                        "綱路編號: " & subSpace(mskCircuitNo.Text) & ";" & _
                        "收費方式: " & subSpace(gimCMCode.GetDispStr) & ";" & _
                        "收費項目:" & subSpace(gimCitemCode.GetDispStr) & ";" & _
                        "單據類別: " & subSpace(gimBillType.GetDispStr) & ";" & _
                        "客戶類別:" & subSpace(gimClassCode.GetDispStr) & ";" & _
                        "客戶狀態:" & subSpace(gimStatusCode.GetDispStr) & ";" & _
                        "服務區　:" & subSpace(gimServCode.GetDispStr) & ";" & _
                        "行政區　:" & subSpace(gimAreaCode.GetDispStr) & ";" & _
                        "街道範圍:" & subSpace(gimStrtCode.GetDispStr) & ";" & _
                        "大樓編號: " & subSpace(gimMduId.GetDispStr) & ";" & _
                        "小計依據: " & strpagetype

   Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub subPrint()
  On Error GoTo ChkErr
    Dim lngCount As Long
      strsql = "Select A.*,B.* From So52C0A2 A,So52C0A3 B Where A.GroupName=B.GroupName"
      Call PrintRpt(GetPrinterName(5), RptName("SO52C0"), , "每月每日實收月費分佈表 [SO52C0A]", strsql, strChooseString, , True, "TMP0000.MDB", , strGroupName, GiPaperLandscape)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subPrint")
End Sub

Private Function subInsertMDB(ByRef lngCount As Long) As Boolean
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim rsTest As New ADODB.Recordset
    Dim strWhatDay As String
      If Not GetRS(rsTmp, "Select GroupName,To_Char(RealDate,'DD') RealDate,Count(Distinct Custid) CountCustId,Count(Distinct BillNo) CountBillNo ,sum(RealAmt) RealAmt From " & strViewName & " Group By GroupName,To_Char(RealDate,'DD')") Then Exit Function
      'SElect To_Char(RealDate,'DD') RealDate,Count(Distinct Custid) CountCustId,Count(Distinct BillNo) ,sum(RealAmt) From TMP_A164449So52C0A Group By ServArea,To_Char(RealDate,'DD')"
      '第一次的MDB
      On Error Resume Next
      cnn.Execute "Drop Table So52C0A1"
      cnn.Execute "Drop Table So52C0A2"
      cnn.Execute "Drop Table So52C0A3"
      cnn.Execute "Drop Table So52C0A21"
      cnn.Execute "Drop Table So52C0A31"
      cnn.BeginTrans

      On Error GoTo ChkErr
      cnn.Execute "Create Table So52C0A1 (" & GetMonthField & ", " & _
              "GroupName Text(40) ,RealDate Text(2))"
      Do While Not rsTmp.EOF
          cnn.Execute "Insert Into So52C0A1 (GroupName,RealDate " & _
              ",CountCustId" & rsTmp("RealDate") & _
              ",CountBillNo" & rsTmp("RealDate") & _
              ",RealAmt" & rsTmp("RealDate") & ") Values ('" & _
              IIf(IsNull(rsTmp(0)), "空白", rsTmp(0)) & "'," & _
              GetNullString(rsTmp(1)) & "," & _
              GetNullString(rsTmp(2)) & "," & _
              GetNullString(rsTmp(3)) & "," & _
              GetNullString(rsTmp(4)) & ")"
          DoEvents
          rsTmp.MoveNext
      Loop
     cnn.CommitTrans

      '第二次
      On Error GoTo ChkErr
      cnn.Execute "Select GroupName," & GetMonthField2 & " Into So52C0A2 From So52C0A1 Group By GroupName "
      cnn.Execute "Create Index I_So52C0A2_GroupName On So52C0A2(GroupName)"
      cnn.Execute "Select * Into So52C0A3 From So52C0A2"
      If Not addMonthField Then Exit Function
      cnn.Execute "Create Index I_So52C0A3_GroupName On So52C0A3(GroupName)"
      
      cnn.Execute "Select " & GetMonthField2 & " Into So52C0A21 From So52C0A2 "
      cnn.Execute "Select " & GetMonthField2 & " Into So52C0A31 From So52C0A3 "
      lngCount = rsTmp.RecordCount
      CloseRecordset rsTmp
      subInsertMDB = True
    Exit Function
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "subInsertMDB")
End Function

Private Function GetMonthField() As String
  On Error Resume Next
    Dim lngCount As Long
    Dim strDefineField As String
    Dim strRealAmtField As String
    Dim strCountBillNoField As String
    Dim strCountCustIdField As String
      For lngCount = 1 To 31
          strRealAmtField = strRealAmtField & ",RealAmt" & Format(lngCount, "00") & " Long"
          strCountBillNoField = strCountBillNoField & ",CountBillNo" & Format(lngCount, "00") & " Long"
          strCountCustIdField = strCountCustIdField & ",CountCustId" & Format(lngCount, "00") & " Long"
      Next
      GetMonthField = Mid(strRealAmtField & strCountBillNoField & strCountCustIdField, 2)
End Function

Private Function GetMonthField2() As String
  On Error Resume Next
    Dim lngCount As Long
    Dim strDefineField As String
    Dim strRealAmtField As String
    Dim strCountBillNoField As String
    Dim strCountCustIdField As String
      For lngCount = 1 To 31
          strRealAmtField = strRealAmtField & ",IIF(ISNULL(Sum(RealAmt" & Format(lngCount, "00") & ")),0,Sum(RealAmt" & Format(lngCount, "00") & "))  as RealAmt" & Format(lngCount, "00")
          strCountBillNoField = strCountBillNoField & ",IIF(ISNULL(Sum(CountBillNo" & Format(lngCount, "00") & ")),0, Sum(CountBillNo" & Format(lngCount, "00") & ")) as CountBillNo" & Format(lngCount, "00")
          strCountCustIdField = strCountCustIdField & ",IIF(ISNULL(Sum(CountCustId" & Format(lngCount, "00") & ")),0, Sum(CountCustId" & Format(lngCount, "00") & " )) as CountCustId" & Format(lngCount, "00")
      Next
      GetMonthField2 = Mid(strRealAmtField & strCountBillNoField & strCountCustIdField, 2)
End Function

Private Function addMonthField() As Boolean
  On Error GoTo ChkErr
    Dim lngCount As Long
    Dim strDefineField As String
    Dim strRealAmtField As String
    Dim strCountBillNoField As String
    Dim strCountCustIdField As String
        For lngCount = 2 To 31
            strRealAmtField = "  RealAmt" & Format(lngCount, "00") & " = RealAmt" & Format(lngCount, "00") & " + RealAmt" & Format(lngCount - 1, "00")
            strCountBillNoField = ", CountBillNo" & Format(lngCount, "00") & " = CountBillNo" & Format(lngCount, "00") & " + CountBillNo" & Format(lngCount - 1, "00")
            strCountCustIdField = ", CountCustId" & Format(lngCount, "00") & "= CountCustId" & Format(lngCount, "00") & " + CountCustId" & Format(lngCount - 1, "00")
            cnn.Execute "Update So52C0A3 Set " & strRealAmtField & strCountBillNoField & strCountCustIdField
        Next
        addMonthField = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "addMonthField")
End Function

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGim
    Call subGil
    'gilCompCode.ListIndex = 1
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
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱")
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱")
    Call SetgiMultiAddItem(gimBillType, "B,I,P,M,T", "收費單,裝機單,停拆移機單,維修單,臨時收費單", "單據代號", "單據名稱")
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    Call SetgiMulti(gimStatusCode, "CodeNo", "Description", "CD035", "客戶狀態代碼", "客戶狀態名稱")
    Call SetgiMulti(gimServCode, "CodeNo", "Description", "CD002", "服務區代碼", "服務區名稱")
    Call SetgiMulti(gimAreaCode, "CodeNo", "Description", "CD001", "行政區代碼", "行政區名稱")
    Call SetgiMulti(gimStrtCode, "CodeNo", "Description", "CD017", "街道代碼", "街道名稱")
    Call SetgiMulti(gimMduId, "MduId", "Name", "SO017", "大樓編號", "大樓名稱")
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    Call SetgiList(gilClctEn, "EmpNo", "EmpName", "CM003")
    'Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub
  
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO52C0A)
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
  ReleaseCOM frmSO52C0A
End Sub

Private Sub gdaRealDate1_GotFocus()
  On Error Resume Next
  If gdaRealDate1.GetValue = "" Then gdaRealDate1.SetValue (Format(RightDate, "YYYYMM") & "01")
End Sub

Private Sub gdaRealDate2_GotFocus()
  On Error Resume Next
    If gdaRealDate1.GetValue <> "" Then
      gdaRealDate2.SetValue GetLastDate(gdaRealDate1.GetValue(True))
    Else
      gdaRealDate2.SetValue ""
    End If
End Sub

Private Sub gdaRealDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaRealDate1, gdaRealDate2)
End Sub

Private Sub gdaShouldDate1_GotFocus()
  On Error Resume Next
  If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetValue (RightDate)
End Sub

Private Sub gdaShouldDate2_GotFocus()
  On Error Resume Next
  If gdaShouldDate1.GetValue = "" Or gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue (gdaShouldDate1.GetValue)
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaShouldDate1, gdaShouldDate2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    
    GiMultiFilter gimServCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimAreaCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimStrtCode, , gilCompCode.GetCodeNo
    GiMultiFilter gimMduId, , gilCompCode.GetCodeNo
    GiListFilter gilClctEn, , gilCompCode.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Change")
End Sub

Private Sub gilCompCode_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then
        MsgMustBe ("公司別")
        Cancel = True
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Sub gilServiceType_Change()
  On Error GoTo ChkErr
    GiMultiFilter gimCitemCode, gilServiceType.GetCodeNo
    GiMultiFilter gimClassCode, gilServiceType.GetCodeNo
    GiMultiFilter gimCMCode, gilServiceType.GetCodeNo
    'strST = GetCode("CD016", "<>1", gilServiceType.GetCodeNo, , , False)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Sub mskCircuitNo_Validate(Cancel As Boolean)
  On Error Resume Next
    If mskCircuitNo.Text = "" Then
        optCircuitNo.Enabled = True
        optAll.Enabled = True
    Else
        optCircuitNo.Enabled = False
        optAll.Enabled = False
    End If
End Sub

Private Function subCreateView() As Boolean
  On Error Resume Next
      gcnGi.Execute "Drop View " & strViewName
  On Error GoTo ChkErr
    Dim strView As String
    strViewName = GetTmpViewName
    strView = "Create View " & strViewName & " as (" & _
                  "SELECT " & strGroupField & " GroupName,A.REALDATE,A.REALAMT,A.BillNo,A.CustId " & _
                      " From SO033 " & strFrom & " Where " & strChoose & _
              " Union All " & _
                  "SELECT " & strGroupField & " GroupName,A.REALDATE,A.REALAMT,A.BillNo,A.CustId " & _
                      " From SO034 " & strFrom & " Where " & strChoose & ")"
    gcnGi.Execute strView
    SendSQL strView, True
    subCreateView = True
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function
