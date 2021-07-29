VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.17#0"; "csMulti.ocx"
Begin VB.Form frmSO52G0A 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶繳款習性分析報表 [SO52G0A]"
   ClientHeight    =   7275
   ClientLeft      =   1950
   ClientTop       =   1530
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO52G0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form17"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Frame1 
      Caption         =   "地址依據"
      ForeColor       =   &H00FF0000&
      Height          =   765
      Left            =   3930
      TabIndex        =   43
      Top             =   60
      Width           =   3735
      Begin VB.OptionButton optMailAddress 
         Caption         =   "郵寄地址"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2460
         TabIndex        =   4
         Top             =   360
         Width           =   1125
      End
      Begin VB.OptionButton optChargeAddress 
         Caption         =   "收費地址"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1290
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optInstAddress 
         Caption         =   "裝機地址"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdLabel 
      Caption         =   "郵寄標籤"
      Height          =   525
      Left            =   3270
      TabIndex        =   26
      Top             =   6630
      Width           =   1395
   End
   Begin VB.Frame fraPreCMCode 
      Caption         =   "前次收費條件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   30
      TabIndex        =   36
      Top             =   4620
      Width           =   7635
      Begin Gi_Multi.GiMulti gimPreCMCode 
         Height          =   375
         Left            =   90
         TabIndex        =   22
         Top             =   750
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   661
         ButtonCaption   =   "前次收費方式"
      End
      Begin Gi_Multi.GiMulti gimPreBillType 
         Height          =   375
         Left            =   90
         TabIndex        =   23
         Top             =   1110
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   661
         ButtonCaption   =   "單  據  類  別"
         DataType        =   2
         DIY             =   -1  'True
         Exception       =   -1  'True
      End
      Begin Gi_Date.GiDate gdaPreRealDate1 
         Height          =   345
         Left            =   1080
         TabIndex        =   18
         Top             =   315
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
      Begin Gi_Date.GiDate gdaPreRealDate2 
         Height          =   345
         Left            =   2490
         TabIndex        =   19
         Top             =   315
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
      Begin prjNumber.GiNumber gnbPre 
         Height          =   315
         Left            =   4350
         TabIndex        =   20
         Top             =   330
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         AllowZero       =   0   'False
      End
      Begin prjNumber.GiNumber gnbSame 
         Height          =   315
         Left            =   5640
         TabIndex        =   21
         Top             =   330
         Visible         =   0   'False
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         AllowZero       =   0   'False
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "連續"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5190
         TabIndex        =   42
         Top             =   390
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "次相同收費方式"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6180
         TabIndex        =   41
         Top             =   390
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "次"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4920
         TabIndex        =   40
         Top             =   390
         Width           =   195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "統計前"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3720
         TabIndex        =   39
         Top             =   390
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "~"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2280
         TabIndex        =   38
         Top             =   390
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "實收日期"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   210
         TabIndex        =   37
         Top             =   390
         Width           =   780
      End
   End
   Begin VB.Frame fraCMCode 
      Caption         =   "本次收費條件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2325
      Left            =   30
      TabIndex        =   30
      Top             =   2160
      Width           =   7635
      Begin VB.Frame Frame2 
         Caption         =   "收費單狀態"
         ForeColor       =   &H00FF0000&
         Height          =   705
         Left            =   120
         TabIndex        =   33
         Top             =   270
         Width           =   3255
         Begin VB.OptionButton optAll 
            Caption         =   "全部"
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   2220
            TabIndex        =   10
            Top             =   300
            Value           =   -1  'True
            Width           =   705
         End
         Begin VB.OptionButton optNoReal 
            Caption         =   "未收"
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1170
            TabIndex        =   9
            Top             =   300
            Width           =   705
         End
         Begin VB.OptionButton optReal 
            Caption         =   "己收"
            ForeColor       =   &H00800080&
            Height          =   315
            Left            =   240
            TabIndex        =   8
            Top             =   300
            Width           =   705
         End
      End
      Begin Gi_Multi.GiMulti gimCMCode 
         Height          =   345
         Left            =   90
         TabIndex        =   16
         Top             =   1530
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   609
         ButtonCaption   =   "本次收費方式"
      End
      Begin CS_Multi.CSmulti gimClctEn 
         Height          =   375
         Left            =   90
         TabIndex        =   15
         Top             =   1200
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   661
         ButtonCaption   =   "收  費  人   員"
      End
      Begin Gi_Date.GiDate gdaRealDate1 
         Height          =   345
         Left            =   4500
         TabIndex        =   13
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         ForeColor       =   8388736
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
      Begin Gi_Date.GiDate gdaRealDate2 
         Height          =   345
         Left            =   5910
         TabIndex        =   14
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         ForeColor       =   8388736
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
      Begin Gi_Date.GiDate gdaShouldDate2 
         Height          =   345
         Left            =   5910
         TabIndex        =   12
         Top             =   270
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
         Left            =   4500
         TabIndex        =   11
         Top             =   270
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
      Begin Gi_Multi.GiMulti gimBillType 
         Height          =   375
         Left            =   90
         TabIndex        =   17
         Top             =   1860
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   661
         ButtonCaption   =   "單  據  類  別"
         DataType        =   2
         DIY             =   -1  'True
         Exception       =   -1  'True
      End
      Begin VB.Label lblShouldDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "應收日期"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3630
         TabIndex        =   35
         Top             =   360
         Width           =   780
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "~"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5700
         TabIndex        =   34
         Top             =   360
         Width           =   105
      End
      Begin VB.Label lblRealDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "實收日期"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   3630
         TabIndex        =   32
         Top             =   810
         Width           =   780
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "~"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   5700
         TabIndex        =   31
         Top             =   810
         Width           =   105
      End
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   525
      Left            =   1680
      TabIndex        =   25
      Top             =   6630
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   525
      Left            =   6330
      TabIndex        =   27
      Top             =   6630
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      Height          =   525
      Left            =   120
      TabIndex        =   24
      Top             =   6630
      Width           =   1245
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   900
      TabIndex        =   0
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
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   900
      TabIndex        =   1
      Top             =   480
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
   Begin Gi_Multi.GiMulti gimStatusCode 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   930
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "客  戶  狀  態"
   End
   Begin CS_Multi.CSmulti gimClassCode 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1290
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "客  戶  類  別"
   End
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1650
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   661
      ButtonCaption   =   "收  費  項  目"
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   60
      TabIndex        =   29
      Top             =   150
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "服務類別"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   60
      TabIndex        =   28
      Top             =   540
      Width           =   780
   End
End
Attribute VB_Name = "frmSO52G0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table: SO033 or SO034 A,SO001,SO002
Option Explicit
Dim strPreChoose As String
Dim strSO002Choose As String
Dim blnLabel As Boolean
Dim strAddressField As String
Dim strWhere As String

Private Sub cmdLabel_Click()
  On Error GoTo ChkErr
    blnLabel = True
    cmdPrint_Click
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      Call PreviousRpt(GetPrinterName(5), RptName("SO52G0"), Me.Name)
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
          Call subInsertMDB
          Call subPrint
        Me.Enabled = True
      Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: MsgMustBe ("公司別"): Exit Function
    If gnbPre.Text = "" Or Val(gnbPre.Text) = 0 Then gnbPre.SetFocus: MsgMustBe ("統計前幾次"): Exit Function
    'If gnbSame.Text = "" Or Val(gnbSame.Text) = 0 Then gnbSame.SetFocus: MsgMustBe ("連續幾次相同收費方式"): Exit Function
  IsDataOk = True
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub subChoose()
  On Error GoTo ChkErr
    Dim strRealStatus As String
    Dim strAddress As String
    strChoose = ""
    strPreChoose = ""
    strChooseString = ""
    strSO002Choose = ""
    strAddressField = ""
    strWhere = ""
  '共同條件-GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("A.CompCode=" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("A.ServiceType='" & gilServiceType.GetCodeNo & "'")
  
  '共同條件-GiMulti
    If gimClassCode.GetQryStr <> "" Then Call subAnd("A.ClassCode " & gimClassCode.GetQryStr)
    If gimStatusCode.GetQryStr <> "" Then Call subAnd2(strSO002Choose, "SO002.CustStatusCode " & gimStatusCode.GetQryStr)
    strPreChoose = strChoose
    If gimCitemCode.GetQryStr <> "" Then Call subAnd("A.CitemCode " & gimCitemCode.GetQryStr)
    
  '地址依據
    Select Case True
           Case optInstAddress.Value
                strAddressField = ",SO001.InstAddrNo AddrNo,SO001.InstAddress Address "
                strWhere = "SO001.InstAddrNo=SO014.Addrno"
                strAddress = "裝機地址"
           Case optChargeAddress.Value
                strAddressField = ",SO001.ChargeAddrNo AddrNo,SO001.ChargeAddress Address "
                strWhere = "SO001.ChargeAddrNo=SO014.Addrno"
                strAddress = "收費地址"
           Case optMailAddress.Value
                strAddressField = ",SO001.MailAddrNo AddrNo,SO001.MailAddress Address "
                strWhere = "SO001.MailAddrNo=SO014.Addrno"
                strAddress = "郵寄地址"
    End Select
    
  '本次條件-日期
    If gdaShouldDate1.GetValue <> "" Then Call subAnd("A.ShouldDate >= To_Date('" & gdaShouldDate1.GetValue & "','YYYYMMDD')")
    If gdaShouldDate2.GetValue <> "" Then Call subAnd("A.ShouldDate < To_Date('" & gdaShouldDate2.GetValue & "','YYYYMMDD')+1")
    If gdaRealDate1.GetValue <> "" Then Call subAnd("A.RealDate >= To_Date('" & gdaRealDate1.GetValue & "','YYYYMMDD')")
    If gdaRealDate2.GetValue <> "" Then Call subAnd("A.RealDate < To_Date('" & gdaRealDate2.GetValue & "','YYYYMMDD')+1")
    
  '本次條件-GiMulti
    If gimClctEn.GetQryStr <> "" Then Call subAnd("A.ClctEn " & gimClctEn.GetQryStr)
    If gimCMCode.GetQryStr <> "" Then Call subAnd("A.CMCode " & gimCMCode.GetQryStr)
    If gimBillType.GetQryStr <> "" Then Call subAnd("SubStr(A.BillNo,7,1) " & gimBillType.GetQryStr)
    
    Select Case True
           Case optReal.Value
                strRealStatus = "己收"
                Call subAnd("A.UCCode Is Null And Nvl(A.CancelFlag,0)=0")
                If gdaRealDate1.GetValue = "" Or gdaRealDate2.GetValue = "" Then Call subAnd("RealDate Is not Null")
           Case optNoReal.Value
                Call subAnd("A.UCCode Is Not Null And Nvl(A.CancelFlag,0)=0")
                Call subAnd("RealDate Is Null")
                strRealStatus = "未收"
           Case optAll.Value
                Call subAnd("Nvl(A.CancelFlag,0)=0")
                strRealStatus = "全部"
    End Select

    
  '前次條件-日期
    If gdaPreRealDate1.GetValue <> "" Then Call subAnd2(strPreChoose, "A.RealDate >= To_Date('" & gdaPreRealDate1.GetValue & "','YYYYMMDD')")
    If gdaPreRealDate2.GetValue <> "" Then Call subAnd2(strPreChoose, "A.RealDate < To_Date('" & gdaPreRealDate2.GetValue & "','YYYYMMDD')+1")
    If Not optNoReal.Value Then
        If gdaPreRealDate1.GetValue = "" Or gdaPreRealDate2.GetValue = "" Then Call subAnd2(strPreChoose, "A.RealDate Is not Null")
    End If
  '前次條件-GiMulti
    If gimPreBillType.GetQryStr <> "" Then Call subAnd2(strPreChoose, "SubStr(A.BillNo,7,1) " & gimPreBillType.GetQryStr)
    If gimPreCMCode.GetQryStr <> "" Then Call subAnd2(strPreChoose, "A.CMCode " & gimPreCMCode.GetQryStr)

    strChooseString = "公司別  :" & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別:" & subSpace(gilServiceType.GetDescription) & ";" & _
                      "地址依據:" & strAddress & ";" & _
                      "客戶狀態:" & subSpace(gimStatusCode.GetDispStr) & ";" & _
                      "客戶類別:" & subSpace(gimClassCode.GetDispStr) & ";" & _
                      "收費項目: " & subSpace(gimCitemCode.GetDispStr) & ";" & _
                      "本次收費狀態:" & strRealStatus & ";" & _
                      "本次應收日期: " & subSpace(gdaShouldDate1.GetValue(True)) & "~" & subSpace(gdaShouldDate2.GetValue(True)) & ";" & _
                      "本次實收日期: " & subSpace(gdaRealDate1.GetValue(True)) & "~" & subSpace(gdaRealDate2.GetValue(True)) & ";" & _
                      "本次收費人員: " & subSpace(gimClctEn.GetDispStr) & ";" & _
                      "本次收費方式: " & subSpace(gimCMCode.GetDispStr) & ";" & _
                      "本次單據類別: " & subSpace(gimBillType.GetDispStr) & ";" & _
                      "統計前　" & gnbPre.Text & "　次" & ";" & _
                      "前次實收日期: " & subSpace(gdaPreRealDate1.GetValue(True)) & "~" & subSpace(gdaPreRealDate2.GetValue(True)) & ";" & _
                      "前次單據類別: " & subSpace(gimPreBillType.GetDispStr) & ";" & _
                      "前次收費方式: " & subSpace(gimPreCMCode.GetDispStr)
                      
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub subPrint()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
  Dim strSubQry1 As String
  Dim strSubQry2 As String
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    Set rsTmp = cnn.Execute("Select Count(*) as intCount From SO52G0A ")
    If rsTmp("intCount") = 0 Then
        MsgNoRcd
        SendSQL , , True
    Else
        If blnLabel Then
            strsql = "SELECT * FROM SO52G0ALabel M"
            Call PrintRpt(GetPrinterName(5), RptName("SO52G0", "Label"), , Me.Name, strsql, "", , True, "Tmp0000.Mdb")
        Else
            strsql = "SELECT * FROM SO52G0A M"
            strSubQry1 = "Select * From SO52G0A1 A"
            strSubQry2 = "Select * From SO52G0A2 A"
            Call PrintRpt(GetPrinterName(5), RptName("SO52G0"), , Me.Name, strsql, strChooseString, , True, "Tmp0000.Mdb", , , , , strSubQry1, strSubQry2)
        End If
    End If
    CloseRecordset rsTmp
    blnLabel = False
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subPrint")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGim
    Call subGil
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
    gnbPre.Text = 1
    'gnbSame.Text = 1
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
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    Call SetgiMulti(gimStatusCode, "CodeNo", "Description", "CD035", "客戶狀態代碼", "客戶狀態名稱")
    Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱")
  
    Call SetgiMulti(gimClctEn, "EmpNo", "EmpName", "CM003", "收費員代號", "收費員名稱")
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱")
    Call SetgiMultiAddItem(gimBillType, "B,I,P,M,T", "收費單,裝機單,停拆移機單,維修單,臨時收費單")
    
    Call SetgiMulti(gimPreCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱")
    Call SetgiMultiAddItem(gimPreBillType, "B,I,P,M,T", "收費單,裝機單,停拆移機單,維修單,臨時收費單")
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub
  
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO52G0A)
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
  ReleaseCOM frmSO52G0A
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
    Cancel = ChkDate2(gdaRealDate1, gdaRealDate2)
End Sub

Private Sub gdaPreRealDate1_GotFocus()
  On Error GoTo ChkErr
    If gdaPreRealDate1.GetValue = "" Then gdaPreRealDate1.SetValue (RightDate)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaPreRealDate1_GotFocus")
End Sub

Private Sub gdaPreRealDate2_GotFocus()
  On Error GoTo ChkErr
    If gdaPreRealDate1.GetValue = "" Or gdaPreRealDate2.GetValue = "" Then gdaPreRealDate2.SetValue (gdaPreRealDate1.GetValue)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaPreRealDate2_GotFocus")
End Sub

Private Sub gdaPreRealDate2_Validate(Cancel As Boolean)
  On Error Resume Next
    Cancel = ChkDate2(gdaPreRealDate1, gdaPreRealDate2)
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
    If gdaShouldDate1.GetValue = "" Or gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue (gdaShouldDate1.GetValue)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gdaShouldDate2_GotFocus")
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
  On Error Resume Next
    Cancel = ChkDate2(gdaShouldDate1, gdaShouldDate2)
End Sub

Private Sub subInsertMDB()
  On Error GoTo ChkErr
    Dim strTmpSQL As String
    Dim strField As String
    Dim rsMDB As New ADODB.Recordset
    Dim rsPrevious As New ADODB.Recordset
    Dim intPer As Long
    Dim strPreSQL As String
    Dim strIndexSO033 As String
    Dim strIndexSO034 As String
    
    'Insert Table "SO52G0A" 全部
      If blnLabel Then
          strAddressField = strAddressField & ",SO014.ZipCode "
          strWhere = ") A,SO001,SO014" & IIf(strSO002Choose = "", " Where A.Custid=SO001.Custid And " & strWhere, ",SO002 Where A.Custid=SO001.Custid And A.Custid=SO002.Custid And A.ServiceType=SO002.ServiceType And " & strWhere & " And " & strSO002Choose)
      Else
          strAddressField = strAddressField & ",Null ZipCode "
          strWhere = ") A,SO001" & IIf(strSO002Choose = "", " Where A.Custid=SO001.Custid ", ",SO002 Where A.Custid=SO001.Custid And A.Custid=SO002.Custid And A.ServiceType=SO002.ServiceType And " & strSO002Choose)
      End If
      
      strField = "A.Custid,A.BillNo,A.CitemCode,A.CitemName,A.CMCode,A.CMName,A.RealDate,A.OldStopDate,A.ServiceType"
      strTmpSQL = "Select " & strField & ",A.PreBillNo,A.PreCMCode,A.PreCMName,SO001.CustName,SO001.Tel1,SO001.Tel2,SO001.Tel3" & strAddressField & _
                  "From(" & _
                          "Select " & strField & ",Null PreBillNo,0 PreCMCode,Null PreCMName From SO033 A Where " & strChoose & _
                          " Union All " & _
                          "Select " & strField & ",Null PreBillNo,0 PreCMCode,Null PreCMName From SO034 A Where " & strChoose & strWhere
                          
      Set rsMDB = gcnGi.Execute(strTmpSQL)
      Call CreateMDBTable(rsMDB, "SO52G0A", cnn)
      SendSQL strTmpSQL, True
      'lngCustId = 0: intCitemCode = 0
      strIndexSO033 = GetUseIndexStr("SO033", "CustId")
      strIndexSO034 = GetUseIndexStr("SO034", "CustId")
      Do While Not rsMDB.EOF
          If optReal.Value Then
          strPreSQL = "Select " & strIndexSO033 & " Custid,BillNo,CMCode,CMName,RealDate,RealStopDate,ServiceType " & _
                          "From SO033 A Where A.CustId = " & rsMDB("CustId") & " And A.CitemCode = " & rsMDB("CitemCode") & " And A.RealDate < to_date(" & Format(rsMDB("RealDate"), "yyyymmdd") & ",'yyyymmdd') And " & strPreChoose & _
                    " Union All " & _
                      "Select  " & strIndexSO034 & " Custid,BillNo,CMCode,CMName,RealDate,RealStopDate,ServiceType " & _
                          "From SO034 A Where A.CustId = " & rsMDB("CustId") & " And A.CitemCode = " & rsMDB("CitemCode") & " And A.RealDate < to_date(" & Format(rsMDB("RealDate"), "yyyymmdd") & ",'yyyymmdd') And " & strPreChoose
          Else
          strPreSQL = "Select " & strIndexSO033 & " Custid,BillNo,CMCode,CMName,RealDate,RealStopDate,ServiceType " & _
                          "From SO033 A Where A.CustId = " & rsMDB("CustId") & " And A.CitemCode = " & rsMDB("CitemCode") & " And " & strPreChoose & _
                    " Union All " & _
                      "Select  " & strIndexSO034 & " Custid,BillNo,CMCode,CMName,RealDate,RealStopDate,ServiceType " & _
                          "From SO034 A Where A.CustId = " & rsMDB("CustId") & " And A.CitemCode = " & rsMDB("CitemCode") & " And " & strPreChoose
          
          End If
          If strSO002Choose <> "" Then
              strPreSQL = "Select A.BillNo,A.CMCode,A.CMName,A.RealDate,A.RealStopDate From(" & strPreSQL & ") A,SO002 Where A.Custid=SO002.Custid And A.ServiceType=SO002.ServiceType And " & strSO002Choose & " Order By RealDate Desc,RealStopDate Desc"
          Else
              strPreSQL = strPreSQL & " Order By RealDate Desc,RealStopDate Desc"
          End If
          
          If Not GetRS(rsPrevious, strPreSQL) Then Exit Sub
                                                                        
          'If Not rsPrevious.EOF Then
          intPer = 1
          Do While Not rsPrevious.EOF
              If intPer = Val(gnbPre.Text) Then
                  cnn.BeginTrans
                  cnn.Execute "INSERT INTO SO52G0A (Custid,BillNo,CitemCode,CitemName,CMCode,CMName,RealDate,OldStopDate,PREBILLNO,PRECMCODE,PRECMNAME,CustName,Tel1,Tel2,Tel3,AddrNo,Address,ZipCode) " & _
                              "VALUES (" & _
                              GetNullString(rsMDB("Custid"), giLongV) & "," & _
                              GetNullString(rsMDB("BillNo")) & "," & _
                              GetNullString(rsMDB("CitemCode"), giLongV) & "," & _
                              GetNullString(rsMDB("CitemName")) & "," & _
                              GetNullString(rsMDB("CMCode"), giLongV) & "," & _
                              GetNullString(rsMDB("CMName")) & "," & _
                              GetNullString(rsMDB("RealDate"), giDateV, giAccessDb) & "," & _
                              GetNullString(rsMDB("OldStopDate"), giDateV, giAccessDb) & "," & _
                              GetNullString(rsPrevious("BillNo")) & "," & _
                              GetNullString(rsPrevious("CMCode")) & "," & _
                              GetNullString(rsPrevious("CMName")) & "," & _
                              GetNullString(rsMDB("CustName")) & "," & _
                              GetNullString(rsMDB("Tel1")) & "," & _
                              GetNullString(rsMDB("Tel2")) & "," & _
                              GetNullString(rsMDB("Tel3")) & "," & _
                              GetNullString(rsMDB("AddrNo")) & "," & _
                              GetNullString(rsMDB("Address")) & "," & _
                              GetNullString(rsMDB("ZipCode")) & ")"
                  cnn.CommitTrans
                  Exit Do
                  DoEvents
              End If
             intPer = intPer + 1
             rsPrevious.MoveNext
             DoEvents
          Loop
          'End If
          rsMDB.MoveNext
          DoEvents
      Loop

      If blnLabel Then
          On Error Resume Next
            cnn.Execute "Drop Table SO52G0ALabel"
            If Err.Number = -2147217900 Then MsgBox "無法刪除〔 SO52G0ALabel ] 資料表 ", vbCritical, "錯誤..": Exit Sub
          On Error GoTo ChkErr
          cnn.Execute "Select Custid,CustName,Tel1,Address,ZipCode Into SO52G0ALabel From SO52G0A Group by Custid,CustName,Tel1,Address,ZipCode"
      Else
          On Error Resume Next
            cnn.Execute "Drop Table SO52G0A1"
            cnn.Execute "Drop Table SO52G0A2"
            If Err.Number = -2147217900 Then MsgBox "無法刪除〔 SO52G0A1 或 SO52G0A2 ] 資料表 ", vbCritical, "錯誤..": Exit Sub
            
          On Error GoTo ChkErr
          'Insert Table "SO52G0A1" 收費方式本次與前次相同
            cnn.Execute "Select Custid,CitemCode,CitemName,CMCode,CMName,PreCMCode Into SO52G0A1 From SO52G0A Where CMCode=PreCMCode"
          'Insert Table "SO52G0A2" 收費方式本次與前次不相同
            cnn.Execute "Select Custid,CitemCode,CitemName,CMCode,CMName,PreCMCode,PreCMName Into SO52G0A2 From SO52G0A Where CMCode<>PreCMCode"
      End If
      CloseRecordset rsMDB
      CloseRecordset rsPrevious
  Exit Sub
ChkErr:
'  cnn.RollbackTrans
  Call ErrSub(Me.Name, "subInsertMDB")
End Sub
        
Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
    GiMultiFilter gimClctEn, , gilCompCode.GetCodeNo
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
    GiMultiFilter gimCMCode, gilServiceType.GetCodeNo
    GiMultiFilter gimPreCMCode, gilServiceType.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilServiceType_Change")
End Sub

'Private Sub gimNoCMCode_Change()
'  On Error GoTo ChkErr
'    If gimNoCMCode.GetQueryCode <> "" Then
'        gimCMCode.Clear
'        gimCMCode.Enabled = False
'        gimPreCMCode.Clear
'        gimPreCMCode.Enabled = False
'    Else
'        gimCMCode.Enabled = True
'        gimPreCMCode.Enabled = True
'    End If
'  Exit Sub
'ChkErr:
'  Call ErrSub(Me.Name, "gimNoCMCode_Change")
'End Sub

Private Sub optAll_Click()
  On Error Resume Next
    gdaRealDate1.Enabled = Not optAll.Value
    gdaRealDate2.Enabled = Not optAll.Value
    gdaRealDate1.SetValue ""
    gdaRealDate2.SetValue ""
End Sub

Private Sub optNoReal_Click()
  On Error Resume Next
    gdaRealDate1.Enabled = Not optNoReal.Value
    gdaRealDate2.Enabled = Not optNoReal.Value
    gdaRealDate1.SetValue ""
    gdaRealDate2.SetValue ""
End Sub

Private Sub optReal_Click()
  On Error Resume Next
    gdaRealDate1.Enabled = optReal.Value
    gdaRealDate2.Enabled = optReal.Value
End Sub
