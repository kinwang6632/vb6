VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.0#0"; "GiTime.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#1.17#0"; "csMulti.ocx"
Begin VB.Form frmSO55A0A 
   BorderStyle     =   1  '單線固定
   Caption         =   "申請轉帳記錄報表 [SO55A0A]"
   ClientHeight    =   6315
   ClientLeft      =   1725
   ClientTop       =   1500
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO55A0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form46"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.TextBox txtStopYM2 
      Height          =   345
      Left            =   7200
      MaxLength       =   6
      TabIndex        =   10
      Top             =   450
      Width           =   735
   End
   Begin VB.TextBox txtStopYM1 
      Height          =   345
      Left            =   6030
      MaxLength       =   6
      TabIndex        =   9
      Top             =   450
      Width           =   735
   End
   Begin VB.TextBox txtCustId 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2505
      TabIndex        =   28
      Top             =   4665
      Visible         =   0   'False
      Width           =   6315
   End
   Begin VB.Frame fraPage 
      Caption         =   "群組小計"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   4950
      TabIndex        =   36
      Top             =   1440
      Width           =   4035
      Begin VB.OptionButton optPWCode 
         Caption         =   "付費意願"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2640
         TabIndex        =   15
         Top             =   240
         Width           =   1155
      End
      Begin VB.OptionButton optBankCode 
         Caption         =   "銀行別"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   330
         TabIndex        =   13
         Top             =   270
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optIntroID 
         Caption         =   "介紹人"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1440
         TabIndex        =   14
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      Height          =   525
      Left            =   330
      TabIndex        =   25
      Top             =   5730
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   525
      Left            =   7410
      TabIndex        =   27
      Top             =   5730
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   525
      Left            =   2010
      TabIndex        =   26
      Top             =   5730
      Width           =   1395
   End
   Begin Gi_Multi.GiMulti gimCMCode 
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   2745
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   661
      ButtonCaption   =   "收  費  方  式"
   End
   Begin Gi_Multi.GiMulti gimCardCode 
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   3105
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   661
      ButtonCaption   =   "信 用 卡 種 類"
   End
   Begin Gi_Date.GiDate gdaSnactionDate2 
      Height          =   375
      Left            =   2580
      TabIndex        =   5
      Top             =   975
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
   Begin Gi_Date.GiDate gdaSnactionDate1 
      Height          =   375
      Left            =   1140
      TabIndex        =   4
      Top             =   975
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
   Begin Gi_Date.GiDate gdaPropDate2 
      Height          =   375
      Left            =   2580
      TabIndex        =   3
      Top             =   510
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
   Begin Gi_Date.GiDate gdaPropDate1 
      Height          =   375
      Left            =   1140
      TabIndex        =   2
      Top             =   510
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
      Left            =   6030
      TabIndex        =   8
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
   Begin Gi_Time.GiTime gdtAcceptTime1 
      Height          =   345
      Left            =   1140
      TabIndex        =   0
      Top             =   90
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      ForeColor       =   16711680
      BackColor       =   16777215
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
   Begin Gi_Time.GiTime gdtAcceptTime2 
      Height          =   345
      Left            =   3210
      TabIndex        =   1
      Top             =   90
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   609
      ForeColor       =   16711680
      BackColor       =   16777215
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
      TabIndex        =   19
      Top             =   3480
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   661
      ButtonCaption   =   "介  紹  媒  介"
   End
   Begin Gi_Multi.GiMulti gimIntroId 
      Height          =   330
      Left            =   0
      TabIndex        =   24
      Top             =   5310
      Visible         =   0   'False
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   582
      ButtonCaption   =   "介      紹     人"
   End
   Begin Gi_Multi.GiMulti gimBankCode 
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   3855
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   661
      ButtonCaption   =   "銀  行  名  稱"
   End
   Begin Gi_Multi.GiMulti gimOrder 
      Height          =   330
      Left            =   15
      TabIndex        =   23
      Top             =   4950
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   582
      ButtonCaption   =   "排  序  方  式"
      DataType        =   2
      ColumnOrder     =   -1  'True
   End
   Begin CS_Multi.CSmulti gimAcceptEn 
      Height          =   375
      Left            =   15
      TabIndex        =   16
      Top             =   2400
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   661
      ButtonCaption   =   "受  理  人  員"
   End
   Begin Gi_Multi.GiMulti gimStatusCode 
      Height          =   345
      Left            =   0
      TabIndex        =   21
      Top             =   4230
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   609
      ButtonCaption   =   "客  戶  狀  態"
   End
   Begin Gi_Multi.GiMulti gimPWCode 
      Height          =   345
      Left            =   15
      TabIndex        =   22
      Top             =   4590
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   609
      ButtonCaption   =   "付  費  意  願"
   End
   Begin Gi_Date.GiDate gdaStopDate2 
      Height          =   375
      Left            =   2580
      TabIndex        =   7
      Top             =   1410
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
   Begin Gi_Date.GiDate gdaStopDate1 
      Height          =   375
      Left            =   1140
      TabIndex        =   6
      Top             =   1410
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
   Begin Gi_Date.GiDate gdaAuthStop1 
      Height          =   375
      Left            =   1140
      TabIndex        =   11
      Top             =   1860
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
   Begin Gi_Date.GiDate gdaAuthStop2 
      Height          =   375
      Left            =   2580
      TabIndex        =   12
      Top             =   1860
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2370
      TabIndex        =   44
      Top             =   1980
      Width           =   105
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "授權截止日"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   90
      TabIndex        =   43
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6930
      TabIndex        =   42
      Top             =   510
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "停用日期"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   41
      Top             =   1500
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2370
      TabIndex        =   40
      Top             =   1500
      Width           =   105
   End
   Begin VB.Label lblEx 
      AutoSize        =   -1  'True
      Caption         =   "(格式:MMYYYY,例: 042003)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   5610
      TabIndex        =   39
      Top             =   900
      Width           =   2715
   End
   Begin VB.Label lblCustId 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "客戶編號(以,分隔 以-為區間)"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   60
      TabIndex        =   38
      Top             =   4245
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblStopYM 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "信用卡有效期限"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4620
      TabIndex        =   37
      Top             =   540
      Width           =   1365
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5160
      TabIndex        =   35
      Top             =   180
      Width           =   765
   End
   Begin VB.Label lbl3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2370
      TabIndex        =   34
      Top             =   600
      Width           =   105
   End
   Begin VB.Label lblPropDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "申請日期"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   33
      Top             =   600
      Width           =   780
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2370
      TabIndex        =   32
      Top             =   1065
      Width           =   105
   End
   Begin VB.Label lblSnactionDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "核准日期"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   31
      Top             =   1065
      Width           =   780
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3060
      TabIndex        =   30
      Top             =   180
      Width           =   105
   End
   Begin VB.Label lblAcceptTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "受理日期"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   29
      Top             =   180
      Width           =   780
   End
End
Attribute VB_Name = "frmSO55A0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Table:SO106 A(,SO001 B)
Option Explicit
Dim strChooseInst As String
Dim strChooseService As String
Dim strChoosePR As String
Dim strWhere As String
Dim strOrderName As String
Dim strChoose2 As String
Dim strFrom As String

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
    Call PreviousRpt(GetPrinterName(5), RptName("SO55A0"), Me.Caption)
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
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 66
    IsDataOk = True
  Exit Function
66:
  MsgMustBe (strErrFile)
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub subPrint()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
  Dim strSubQry1 As String
  
     '問題集2807 彙總的數量與明細不符，將COUNT(*)改為 COUNT(DISTINCT A.CUSTID)
     If optBankCode.Value = True Then
        strSubQry1 = "Select A.BankCode GroupCode1,A.BankName GroupName1,A.CardCode GroupCode2,A.CardName GroupName2," & _
                     "Count(Distinct A.Custid) intCnt,'銀行別' Name1,'信用卡別' Name2 From SO106 A ,SO001 B, SO002 C Where A.CustId = B.CustId AND A.CompCode=B.CompCode And A.CustId=C.CustId(+) And A.CompCode=C.CompCode(+)" & _
                      IIf(strChoose <> "", " And ", "") & strChoose & IIf(strChoose2 <> "", " And ", "") & strChoose2 & " Group By A.BankCode,A.BankName,A.CardCode,A.CardName"
     ElseIf optIntroID.Value = True Then
        strSubQry1 = "Select A.IntroID GroupCode1,A.IntroName GroupName1,A.CMCode GroupCode2,A.CMName GroupName2," & _
                     "Count(Distinct A.Custid) intCnt,'介紹人員' Name1,'扣款種類' Name2 From SO106 A ,SO001 B,SO002 C Where A.CustId = B.CustId AND A.CompCode=B.CompCode And A.CustId=C.CustId(+) And A.CompCode=C.CompCode(+)" & _
                     IIf(strChoose <> "", " And ", "") & strChoose & IIf(strChoose2 <> "", " And ", "") & strChoose2 & " Group By A.IntroID,A.IntroName,A.CMCode,A.CMName"
     ElseIf optPWCode.Value = True Then
        strSubQry1 = " Select ' ' GroupCode1,' ' GroupName1,B.PWCode GroupCode2,B.PWName GroupName2," & _
                     "Count(Distinct A.Custid) intCnt,' ' Name1,'付費意願' Name2 FROM SO106 A,SO001 B,SO002 C Where A.CustId = B.CustId AND A.CompCode=B.CompCode " & _
                     " AND A.CustId=C.CustId(+) And A.CompCode=C.CompCode(+)" & IIf(strChoose <> "", " And ", "") & strChoose & IIf(strChoose2 <> "", " AND ", "") & strChoose2 & _
                     " GROUP BY B.PWCode,B.PWName"
     End If
     Call subCreateView(strSubQry1)
     strSubQry1 = "SELECT * From " & strViewName & " V"
     If InStr(strChoose2, "C.") > 0 Then
        strFrom = ",SO002 C "
        strWhere = strWhere & " AND A.CustId=C.CustId(+) And A.CompCode=C.CompCode(+) "
     End If
      
     strsql = "Select A.CustId,A.AccountName,B.Tel1,A.PropDate,A.SnactionDate,A.CardName," & _
              "A.StopYM,A.MediaName,A.IntroName,A.ACHSN From SO106 A, SO001 B " & strFrom & _
              strWhere & " And " & strChoose & IIf(strChoose2 <> "", " AND ", "") & strChoose2 & " Order By A.CustId"
     Set rsTmp = gcnGi.Execute("Select Count(*) As intCount From So106 A,So001 B " & strFrom & strWhere & " And " & strChoose & IIf(strChoose2 <> "", " AND ", "") & strChoose2)
        
     If rsTmp("intCount") = 0 Then
        MsgNoRcd
        SendSQL , , True
     Else
        Call PrintRpt(GetPrinterName(5), RptName("SO55A0"), , Me.Caption, strsql, strChooseString, , True, , , strOrderName, GiPaperLandscape, , strSubQry1)
     End If
     Call CloseRecordset(rsTmp)

  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subPrint")
End Sub

Private Function subIsNull(var1 As Variant)
  On Error GoTo ChkErr
    If varType(var1) = vbNull Then
        subIsNull = "Null"
    Else
        subIsNull = var1
    End If
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subIsNull")
End Function

Private Sub subChoose()
  On Error GoTo ChkErr
    Dim strsql As String
    Dim strYM1 As String    '信用卡有效期限起
    Dim strYM2 As String    '信用卡有效期限迄
    
    strWhere = "Where A.CustId=B.CustId And A.CompCode=B.CompCode "
    strChoose = ""
    strChoose2 = ""
    strChooseString = ""
    strFrom = ""
  '日期
    If gdtAcceptTime1.GetValue <> "" Then Call subAnd("A.AcceptTime >= To_Date('" & gdtAcceptTime1.GetValue & "','YYYYMMDDHH24Miss')")
    If gdtAcceptTime2.GetValue <> "" Then Call subAnd("A.AcceptTime <= To_Date('" & gdtAcceptTime2.GetValue & "','YYYYMMDDHH24Miss')")
    If gdaPropDate1.GetValue <> "" Then Call subAnd("A.PropDate >= To_Date('" & gdaPropDate1.GetValue & "','YYYYMMDD')")
    If gdaPropDate2.GetValue <> "" Then Call subAnd("A.PropDate < To_Date('" & gdaPropDate2.GetValue & "','YYYYMMDD')+1")
    If gdaSnactionDate1.GetValue <> "" Then Call subAnd("A.SnactionDate >= To_Date('" & gdaSnactionDate1.GetValue & "','YYYYMMDD')")
    If gdaSnactionDate2.GetValue <> "" Then Call subAnd("A.SnactionDate < To_Date('" & gdaSnactionDate2.GetValue & "','YYYYMMDD')+1")
    If gdaStopDate1.GetValue <> "" Then Call subAnd("A.StopDate >= To_Date('" & gdaStopDate1.GetValue & "','YYYYMMDD')")
    If gdaStopDate2.GetValue <> "" Then Call subAnd("A.StopDate < To_Date('" & gdaStopDate2.GetValue & "','YYYYMMDD')+1")
    If gdaAuthStop1.GetValue <> "" Then Call subAnd("A.AuthorizeStopDate>=To_Date('" & gdaAuthStop1.GetValue & "','YYYYMMDD')")
    If gdaAuthStop2.GetValue <> "" Then Call subAnd("A.AuthorizeStopDate<To_Date('" & gdaAuthStop2.GetValue & "','YYYYMMDD')+1")
  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("A.CompCode=" & gilCompCode.GetCodeNo)

  'GiMulti
    If gimBankCode.GetQryStr <> "" Then Call subAnd("A.BankCode " & gimBankCode.GetQryStr)
    If gimAcceptEn.GetQryStr <> "" Then Call subAnd("A.AcceptEn " & gimAcceptEn.GetQryStr)
    If gimCMCode.GetQryStr <> "" Then Call subAnd("A.CMCode " & gimCMCode.GetQryStr)
    If gimCardCode.GetQryStr <> "" Then Call subAnd("A.CardCode " & gimCardCode.GetQryStr)
    If gimMediaCode.GetQryStr <> "" Then Call subAnd("A.MediaCode " & gimMediaCode.GetQryStr)
    If gimIntroId.Visible = True And gimIntroId.GetQryStr <> "" Then Call subAnd("A.IntroId " & gimIntroId.GetQryStr)
    If gimStatusCode.GetQryStr <> "" Then Call subAnd2(strChoose2, "C.CustStatusCode " & gimStatusCode.GetQryStr)
    If gimPWCode.GetQryStr <> "" Then Call subAnd2(strChoose2, "B.PWCode " & gimPWCode.GetQryStr)
  '有效期限
  '==============================================================================================
    '2006/11/30 By Tommy 幫KIN修改1446問題集，因為原本在串條件時有限期限的判斷有問題
    '122005會大於112006，所以將期限轉為年月並補上0
    'If txtStopYM1.Text <> "" Then subAnd(A.StopYM>=txtStopYM1.Text)
    'If txtStopYM2.Text <> "" Then subAnd(A.StopYM<=txtStopYM2.Text)
    If txtStopYM1.Text <> "" Then
        strYM1 = Right(txtStopYM1.Text, 4) & Format(Left(txtStopYM1.Text, IIf(Len(txtStopYM1.Text) = 5, 1, 2)), "00")
        Call subAnd("SubStr(A.StopYM,Decode(Length(A.STOPYM),5,2,6,3)) || TRIM(TO_CHAR(SubStr(A.stopYM,1,Decode(Length(A.StopYM),5,1,6,2)),'00'))>=" & strYM1)
    End If
    
    If txtStopYM2.Text <> "" Then
        strYM2 = Right(txtStopYM2.Text, 4) & Format(Left(txtStopYM2.Text, IIf(Len(txtStopYM2.Text) = 5, 1, 2)), "00")
        Call subAnd("SubStr(A.StopYM,Decode(Length(A.STOPYM),5,2,6,3)) || TRIM(TO_CHAR(SubStr(A.stopYM,1,Decode(Length(A.StopYM),5,1,6,2)),'00'))<=" & strYM2)
    End If
  '==============================================================================================
    
    
      '排序方式
    SetGiMultiOrder

    strChooseString = "受理日期:" & subSpace(gdtAcceptTime1.GetValue(True)) & "~" & subSpace(gdtAcceptTime2.GetValue(True)) & ";" & _
                      "申請日期:" & subSpace(gdaPropDate1.GetValue(True)) & "~" & subSpace(gdaPropDate2.GetValue(True)) & ";" & _
                      "核准日期:" & subSpace(gdaSnactionDate1.GetValue(True)) & "~" & subSpace(gdaSnactionDate2.GetValue(True)) & ";" & _
                      "停用日期:" & subSpace(gdaStopDate1.GetValue(True)) & "~" & subSpace(gdaStopDate2.GetValue(True)) & ";" & _
                      "信用卡有效期限:" & subSpace(txtStopYM1.Text) & "~" & subSpace(txtStopYM2.Text) & ";" & _
                      "公司別　:" & subSpace(gilCompCode.GetDescription) & ";" & _
                      "受理人員:" & subSpace(gimAcceptEn.GetDispStr) & ";" & _
                      "收費方式:" & subSpace(gimCMCode.GetDispStr) & ";" & _
                      "信用卡種類:" & subSpace(gimCardCode.GetDispStr) & ";" & _
                      "介紹媒介:" & subSpace(gimMediaCode.GetDispStr) & ";" & _
                      "介紹人　:" & subSpace(gimIntroId.GetDispStr) & ";" & _
                      "銀行名稱:" & subSpace(gimBankCode.GetDispStr) & ";" & _
                      "客戶狀態:" & subSpace(gimStatusCode.GetDispStr) & ";" & _
                      "付費意願:" & subSpace(gimPWCode.GetDispStr) & ";" & _
                      "排序方式:" & subSpace(gimOrder.GetColumnOrderDspStr) & ";" & _
                      "授權截止日:" & subSpace(gdaAuthStop1.GetValue(True)) & "~" & subSpace(gdaAuthStop2.GetValue(True))
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO55A0A)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Public Sub SetGiMultiOrder()
  On Error GoTo ChkErr
    Dim intSort As Integer
    Dim arrOrder() As String
    Dim varOrder As Variant
      intSort = 0: strOrderName = ""
      arrOrder = Split(gimOrder.GetColumnOrderCode, ",")
      For Each varOrder In arrOrder
        Select Case arrOrder(intSort)
               Case "A"
                    strOrderName = strOrderName & ";Sort" & intSort & "={B.Custid}"
               Case "B"
                    strOrderName = strOrderName & ";Sort" & intSort & "={A.AcceptTime}"
               Case "C"
                    strOrderName = strOrderName & ";Sort" & intSort & "={A.PropDate}"
               Case "D"
                    strOrderName = strOrderName & ";Sort" & intSort & "={A.SnactionDate}"
               Case "E"
                    strOrderName = strOrderName & ";Sort" & intSort & "={A.StopDate}"
               Case "F"
                    strOrderName = strOrderName & ";Sort" & intSort & "={A.BankCode}"
               Case "G"
                    strOrderName = strOrderName & ";Sort" & intSort & "={B.PWCode}"
        End Select
        intSort = intSort + 1
      Next
  Exit Sub
ChkErr:
    ErrSub Me.Name, "SetGiMultiOrder"
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
    Call SetgiMulti(gimBankCode, "CodeNo", "Description", "CD018", "銀行代碼", "銀行名稱")
    Call SetgiMulti(gimAcceptEn, "EmpNo", "EmpName", "CM003", "受理人員代碼", "受理人員名稱")
'    Call SetgiMulti(gimUpdEn, "EmpNo", "EmpName", "CM003", "受理人員代碼", "受理人員名稱")
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱")
    Call SetgiMulti(gimCardCode, "CodeNo", "Description", "CD037", "信用卡種類代碼", "信用卡種類名稱")
    Call SetgiMulti(gimMediaCode, "CodeNo", "Description", "CD009", "介紹媒介代碼", "介紹媒介名稱")
    Call SetgiMulti(gimIntroId, "IntroId", "NameP", "So013", "介紹人代碼", "介紹人姓名")
    Call SetgiMultiAddItem(gimOrder, "A,B,C,D,E,F,G", "客戶編號,受理日,申請日,核准日,停用日期,銀行別,付費意願")
    Call SetgiMulti(gimStatusCode, "CodeNo", "Description", "CD035", "客戶狀態代碼", "客戶狀態名稱")
    Call SetgiMulti(gimPWCode, "CodeNo", "Description", "CD020", "付費意願代碼", "付費意願名稱")
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
  ReleaseCOM frmSO55A0A
End Sub

Private Sub gdtAcceptTime1_GotFocus()
  On Error Resume Next
  If gdtAcceptTime1.GetValue = "" Then gdtAcceptTime1.SetValue (RightDate)
End Sub

Private Sub gdtAcceptTime2_GotFocus()
  On Error Resume Next
  If gdtAcceptTime1.GetValue = "" Or gdtAcceptTime2.GetValue = "" Then gdtAcceptTime2.SetValue GetDayLastTime(gdtAcceptTime1.GetValue(True))
End Sub

Private Sub gdtAcceptTime2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdtAcceptTime1, gdtAcceptTime2)
End Sub

Private Sub gdaPropDate1_GotFocus()
  On Error Resume Next
  If gdaPropDate1.GetValue = "" Then gdaPropDate1.SetValue (RightDate)
End Sub

Private Sub gdaPropDate2_GotFocus()
  On Error Resume Next
  If gdaPropDate1.GetValue = "" Or gdaPropDate2.GetValue = "" Then gdaPropDate2.SetValue (gdaPropDate1.GetValue)
End Sub

Private Sub gdaPropDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaPropDate1, gdaPropDate2)
End Sub

Private Sub gdaSnactionDate1_GotFocus()
  On Error Resume Next
  If gdaSnactionDate1.GetValue = "" Then gdaSnactionDate1.SetValue (RightDate)
End Sub

Private Sub gdaSnactionDate2_GotFocus()
  On Error Resume Next
  If gdaSnactionDate1.GetValue = "" Or gdaSnactionDate2.GetValue = "" Then gdaSnactionDate2.SetValue (gdaSnactionDate1.GetValue)
End Sub

Private Sub gdaSnactionDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaSnactionDate1, gdaSnactionDate2)
End Sub
Private Sub gdaStopDate1_GotFocus()
  On Error Resume Next
  If gdaStopDate1.GetValue = "" Then gdaStopDate1.SetValue (RightDate)
End Sub

Private Sub gdaStopDate2_GotFocus()
  On Error Resume Next
  If gdaStopDate1.GetValue = "" Or gdaStopDate2.GetValue = "" Then gdaStopDate2.SetValue (gdaStopDate1.GetValue)
End Sub

Private Sub gdaStopDate2_Validate(Cancel As Boolean)
  On Error Resume Next
  Cancel = ChkDate2(gdaStopDate1, gdaStopDate2)
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    GiMultiFilter gimAcceptEn, , gilCompCode.GetCodeNo
    GiMultiFilter gimIntroId, , gilCompCode.GetCodeNo
'    GiMultiFilter gimBankCode, , gilCompCode.GetCodeNo, , " PrgName='CREDITCARDCITIBANK'"

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

Private Sub gimMediaCode_Change()
  On Error GoTo ChkErr
    If gimMediaCode.GetQueryCode = "" Then gimIntroId.Visible = False: Exit Sub
    Select Case InStr(gimMediaCode.GetQueryCode, ",")
           Case Is > 1
                gimIntroId.Visible = False
                lblCustId.Visible = False
                txtCustId.Visible = False
           Case 0
                'Call ChkRefNo
                ChangIntroVisible Me, gimMediaCode.GetQryStr, gilCompCode.GetCodeNo
    End Select
  Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "gimMediaCode_Change")
End Sub

'Private Sub ChkRefNo()
'  On Error GoTo ChkErr
'    Dim rsCD009 As New ADODB.Recordset
'    Dim strSql As String
'    strSql = "Select RefNo From CD009 Where CodeNo" & gimMediaCode.GetQryStr
'    Call GetRS(rsCD009, strSql)
'    Select Case rsCD009("refno")
'           Case 1
'                lblCustId.Visible = True
'                txtCustId.Visible = True
'                gimIntroId.Visible = False
'           Case 2
'                Call SetgiMulti(gimIntroId, "EmpNo", "EmpName", "CM003", "介紹人代碼", "介紹人姓名", "Where CompCode=" & gilCompCode.GetCodeNo)
'                lblCustId.Visible = False
'                txtCustId.Visible = False
'                gimIntroId.Visible = True
'           Case 3
'                Call SetgiMulti(gimIntroId, "CodeNo", "Description", "CD010", "介紹人代碼", "介紹人姓名", "Where CompCode=" & gilCompCode.GetCodeNo)
'                lblCustId.Visible = False
'                txtCustId.Visible = False
'                gimIntroId.Visible = True
'           Case Else
'                lblCustId.Visible = False
'                txtCustId.Visible = False
'                gimIntroId.Visible = False
'    End Select
'  Exit Sub
'ChkErr:
'  Call ErrSub(Me.Name, "ChkRefNo")
'End Sub

Private Sub subCreateView(ByVal vSQL As String)
  Dim strView As String
   
    On Error Resume Next
      gcnGi.Execute "Drop View " & strViewName
    On Error GoTo ChkErr
        strViewName = GetTmpViewName
        strView = "Create View " & strViewName & " as (" & vSQL & ")"
        gcnGi.Execute strView
        SendSQL strView, True
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subCreateView"
End Sub
