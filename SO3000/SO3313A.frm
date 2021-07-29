VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO3313A 
   BorderStyle     =   1  '單線固定
   Caption         =   "信用卡自動扣款[SO3313A]"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   Icon            =   "SO3313A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   8340
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CheckBox chkFubonIntegrate 
      Caption         =   "財金平台格式"
      Height          =   285
      Left            =   6000
      TabIndex        =   24
      Top             =   1830
      Value           =   1  '核取
      Width           =   1485
   End
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   375
      Left            =   30
      TabIndex        =   23
      Top             =   2250
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   661
      ButtonCaption   =   "收費項目"
   End
   Begin VB.ComboBox cboBillHeadFmt 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1860
      Style           =   2  '單純下拉式
      TabIndex        =   2
      Top             =   630
      Width           =   4065
   End
   Begin VB.Frame fraChinese 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   1470
      Left            =   5970
      TabIndex        =   18
      Top             =   255
      Width           =   2295
      Begin VB.CheckBox chkChinaBank 
         Caption         =   "以中華商銀格式出單"
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
         Left            =   195
         TabIndex        =   6
         Top             =   255
         Width           =   2055
      End
      Begin VB.Frame fraTCB 
         Enabled         =   0   'False
         Height          =   795
         Left            =   180
         TabIndex        =   19
         Top             =   510
         Width           =   1950
         Begin VB.OptionButton optNotTCB 
            Caption         =   "他行卡"
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
            Height          =   180
            Left            =   120
            TabIndex        =   8
            Top             =   510
            Width           =   1035
         End
         Begin VB.OptionButton optTCB 
            Caption         =   "中華商銀自行卡"
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
            Left            =   120
            TabIndex        =   7
            Top             =   210
            Value           =   -1  'True
            Width           =   1710
         End
      End
   End
   Begin VB.ComboBox cboCreditCardType 
      Height          =   300
      ItemData        =   "SO3313A.frx":0442
      Left            =   1860
      List            =   "SO3313A.frx":0444
      Style           =   2  '單純下拉式
      TabIndex        =   3
      Top             =   1005
      Width           =   4065
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6180
      TabIndex        =   12
      Top             =   3450
      Width           =   1410
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   750
      TabIndex        =   11
      Top             =   3450
      Width           =   1320
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2025
      TabIndex        =   9
      Top             =   2895
      Width           =   4965
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7035
      TabIndex        =   10
      Top             =   2910
      Width           =   435
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   300
      Left            =   1860
      TabIndex        =   0
      Top             =   255
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth2       =   3000
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilBankCode 
      Height          =   300
      Left            =   1860
      TabIndex        =   4
      Top             =   1380
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth2       =   3000
      F2Corresponding =   -1  'True
   End
   Begin MSComDlg.CommonDialog comdPath 
      Left            =   5490
      Top             =   3540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   300
      Left            =   2040
      TabIndex        =   1
      Top             =   630
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F2Corresponding =   -1  'True
   End
   Begin Gi_Date.GiDate gdaRealDate 
      Height          =   315
      Left            =   1860
      TabIndex        =   5
      Top             =   1800
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
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
   Begin VB.Label lblAccount 
      AutoSize        =   -1  'True
      Caption         =   "（若空白，則以扣帳日為入帳日）"
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
      Left            =   2985
      TabIndex        =   22
      Top             =   1875
      Width           =   2955
   End
   Begin VB.Label lblDate1 
      AutoSize        =   -1  'True
      Caption         =   "收款入帳日期"
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
      Left            =   615
      TabIndex        =   21
      Top             =   1890
      Width           =   1170
   End
   Begin VB.Label lblBillHeadFmt 
      AutoSize        =   -1  'True
      Caption         =   "多帳戶產生依據設定"
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
      Left            =   30
      TabIndex        =   20
      Top             =   690
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "信用卡別"
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
      Left            =   1005
      TabIndex        =   17
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "資料來源路徑"
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
      Left            =   750
      TabIndex        =   16
      Top             =   2970
      Width           =   1170
   End
   Begin VB.Label lblServiceType 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1035
      TabIndex        =   15
      Top             =   690
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "公司別"
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
      Left            =   1200
      TabIndex        =   14
      Top             =   315
      Width           =   585
   End
   Begin VB.Label lblBankCode 
      AutoSize        =   -1  'True
      Caption         =   "信用卡扣款銀行"
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
      Left            =   420
      TabIndex        =   13
      Top             =   1455
      Width           =   1365
   End
End
Attribute VB_Name = "frmSO3313A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strPrgName As String        '轉帳程式名稱
Private rsSo3314 As New ADODB.Recordset
Private strFilePath As String       '記錄路徑名稱
Private objAgencyIn As Object
Private objAction As Object
Dim lngBankCode As Integer
Public Enum INCreditCardType      ''  2003/10/29  判斷所用的是那一定的銀行
      INCardType_default = 0  '' 預設為 聯合與中華  東森用的
      INCardType_chinatrust = 1   '' 中國信託 目前是 新視波
      INCardType_hsbc = 2  '' 匯豐銀行
      INCardType_chinatrustNew = 3
      INCardType_NewWeb = 4   '藍新科技
      INCardType_UnionBath = 5
      INCardType_Fubon = 6
End Enum
Public e_CreditCardType As INCreditCardType
Private rsBillHeadFmt As New ADODB.Recordset
Private intPara24 As Integer

Private Sub cboCreditCardType_Click()
  On Error GoTo chkErr
    Select Case cboCreditCardType.ItemData(cboCreditCardType.ListIndex)
        Case 0
            e_CreditCardType = CardType_default
            chkChinaBank.Value = 0
            fraChinese.Visible = True
            Call SetgiList(gilBankCode, "CodeNo", "Description", "CD018", , , , , , , _
                                        " Where UPPER(PrgName) = '" & "CREDITCARDUNION'  AND  COMPCODE =" & gilCompCode.GetCodeNo & "   AND StopFlag <>1")
        Case 1
            e_CreditCardType = CardType_chinatrust
            chkChinaBank.Value = 0
            fraChinese.Visible = False
            Call SetgiList(gilBankCode, "CodeNo", "Description", "CD018", , , , , , , _
                                       " Where UPPER(PrgName) = '" & "CREDITCARDCHINATRUST" & "' AND  COMPCODE =" & gilCompCode.GetCodeNo & "   AND StopFlag <>1")
        Case 2
            e_CreditCardType = INCardType_hsbc
            chkChinaBank.Value = 0
            fraChinese.Visible = False
            Call SetgiList(gilBankCode, "CodeNo", "Description", "CD018", , , , , , , _
                                        " Where UPPER(PrgName) = '" & "CREDITCARDHSBC" & "' AND  COMPCODE =" & gilCompCode.GetCodeNo & "   AND StopFlag <>1")
        
        '#3697 增加一組中國信託新格式 By Kin 2007/12/31
        Case 3
            e_CreditCardType = INCardType_chinatrustNew
            chkChinaBank.Value = 0
            fraChinese.Visible = False
            Call SetgiList(gilBankCode, "CodeNo", "Description", "CD018", , , , , , , _
                                        " Where UPPER(PrgName) = '" & "CREDITCARDTRUSTNEW" & "' AND  COMPCODE =" & gilCompCode.GetCodeNo & "   AND StopFlag <>1")
                                        
        '#5494 增加一組藍新科技 By Kin 2010/02/02
        Case 4
            e_CreditCardType = INCardType_NewWeb
            chkChinaBank.Value = 0
            fraChinese.Visible = False
            Call SetgiList(gilBankCode, "CodeNo", "Description", "CD018", , , , , , , _
                                            " Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "' And CompCode=" & gilCompCode.GetCodeNo & " AND STOPFLAG <>1")
        Case 5
        '#5609 增加一組聯合信用卡批次格式 By Kin 2010/06/07
            e_CreditCardType = INCardType_UnionBath
            chkChinaBank.Value = 0
            fraChinese.Visible = True
            Call SetgiList(gilBankCode, "CodeNo", "Description", "CD018", , , , , , , _
                                        " Where UPPER(PrgName) = '" & "CREDITCARDUNION1'  AND  COMPCODE =" & gilCompCode.GetCodeNo & "   AND StopFlag <>1")
        '#7765 By Kin 2018/05/09
        Case 6
            e_CreditCardType = INCardType_Fubon
            chkChinaBank.Value = 0
            fraChinese.Visible = False
            Call SetgiList(gilBankCode, "CodeNo", "Description", "CD018", , , , , , , _
                                        " Where UPPER(PrgName) = '" & "CREDITCARDFUBON'  AND  COMPCODE =" & gilCompCode.GetCodeNo & "   AND StopFlag <>1")
    End Select
    Exit Sub
chkErr:
   Call ErrSub(Me.Name, "cboCreditCardType_Click")
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
        Unload Me
    Exit Sub
End Sub

Private Sub cmdOk_Click()
  On Error GoTo chkErr
    Dim strCitemCode As String
    Dim strRealDate As String

    If Len(gdaRealDate.GetValue) > 0 Then
        strRealDate = gdaRealDate.GetValue
    Else
        strRealDate = ""
    End If
    '#7049 改用CD068A.CitemCode By Kin 2015/07/23
    If cboBillHeadFmt.Visible Then
        strCitemCode = " IN (Select CitemCode From " & GetOwner & "CD068A Where BillHeadFmt='" & cboBillHeadFmt.Text & "') "
    Else
        strCitemCode = gimCitemCode.GetQryStr
    End If

    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    
    GetPrgName
    Set objAgencyIn = CreateObject("CreditCardIn.clsInterface")
    
    With objAgencyIn
        .FilePath = strFilePath                          'INI檔案路徑
        .prgName = strPrgName                            '轉帳程式名稱
        .SourcePath = txtPath.Text                       '原始檔案路徑
        .UpdEn = garyGi(0)                               '異動人員
        .Updname = garyGi(1)                             '#4388 測試不OK,忘記加了這行,導致會將異動人員清空 By Kin 2009/05/13
        .uFubonIntegrate = chkFubonIntegrate.Value = 1
        '#4022 如果啟動多媒体不傳入服務別 By Kin 2008/07/22
        If intPara24 <> 1 And gilServiceType.GetCodeNo <> "" Then
            .ServiceType = gilServiceType.GetCodeNo & ""     '服務類別
        End If
        .CompCode = gilCompCode.GetCodeNo & ""
        .BankId = gilBankCode.GetCodeNo & ""
        .pOwnerName = GetOwner
        
         ''2003/10/30  多一個判斷  是否為中國信託
        Select Case e_CreditCardType
            Case 0
                If chkChinaBank.Value = 1 Then
                    .BankName = "中華商銀入帳"
                Else
                    .BankName = gilBankCode.GetDescription & ""
                End If
            Case 1
                .BankName = "中國信託信用卡入帳"
            Case 2
                .BankName = "匯豐銀行信用卡入帳"
                
            '#3697 增加一組中國信託新格式 By Kin 2007/12/31
            Case 3
                .BankName = "中國信託信用卡(新格式)入帳"
            Case 4
                .BankName = "藍新科技信用卡入帳"
            '#5609 增加一個聯合信用卡批次入帳 By Kin 2010/06/09
            Case 5
                .BankName = "聯合信用卡批次入帳"
            Case 6
                .BankName = "台北富邦入帳"
        End Select
        Set .ugcnGi = gcnGi
        If Len(strPrgName & "") = 0 Then
            MsgBox "程式名稱無設或無使用權限！！", vbExclamation, "提示"
        Else
                '' 2003/10/30因為不止一家信用卡整合進來，所以直接寫死使用這個元件
                ''Set objAction = .InitObject(strPrgName & "")
            Set objAction = .InitObject("CREDITCARDUNION")
            Select Case e_CreditCardType
                '#5609 增加聯合信用卡批次格式 By Kin 2010/06/09
                Case 0, 5, 6
                    If chkChinaBank.Value = 1 Then
                        Call objAction.Action(True, optTCB.Value, CLng(e_CreditCardType), strCitemCode, strRealDate)
                    Else
                        Call objAction.Action(False, False, CLng(e_CreditCardType), strCitemCode, strRealDate)
                    End If
                Case 1
                    Call objAction.Action(False, False, CLng(e_CreditCardType), strCitemCode, strRealDate)
                Case 2
                    Call objAction.Action(False, False, 2, strCitemCode, strRealDate)
                '#3697 增加一組中國信託新格式 By Kin 2007/12/31
                Case 3
                    Call objAction.Action(False, False, CLng(e_CreditCardType), strCitemCode, strRealDate)
                '#5494 增加一組藍新科技 By Kin 2010/02/02
                Case 4
                    Call objAction.Action(False, False, CLng(e_CreditCardType), strCitemCode, strRealDate)
               
                
            End Select
            Set objAction = Nothing
        End If
    End With
    '若更新失敗，則關閉程式
    If objAgencyIn.UpdState = False Then Exit Sub
    Call DefineRs
    Call ScrToRcd
    Set objAgencyIn = Nothing
        
   Exit Sub
chkErr:
   Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Sub ScrToRcd()
  On Error Resume Next
    With rsSo3314
       .AddNew
       .Fields("BankCode") = gilBankCode.GetCodeNo
       .Fields("FilePath") = txtPath.Text
       If Dir(strFilePath & "\" & strPrgName & "In1.log") <> "" Then
          Kill strFilePath & "\" & strPrgName & "In1.log"
       End If
          .Save strFilePath & "\" & strPrgName & "In1.log"
    End With
  Exit Sub

End Sub

Private Sub cmdPath_Click()
  On Error GoTo chkErr
    With comdPath
        .FileName = txtPath.Text
        If e_CreditCardType = INCardType_NewWeb Then
            .Filter = "藍新科技回覆檔(*.*R)|*.*R"
        ElseIf e_CreditCardType = INCardType_UnionBath Then
            .Filter = "聯合信用卡批次回覆檔(*.RSP)|*.RSP"
        ElseIf e_CreditCardType = INCardType_Fubon Then
            .Filter = "台北富邦回覆檔(*.DAT)|*.DAT"
        Else
            .Filter = "文字檔 (*.txt;*.dat)|*.txt;*.dat|所有檔案 (*.*)|*.*"
        End If
        
        .InitDir = ReadGICMIS1("ErrLogPath")
        .Action = 1
        txtPath.Text = .FileName
    End With
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdPath_Click")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo chkErr
    Dim strErrMsg As String
    If Len(gilCompCode.GetCodeNo & "") = 0 Then
        strErrMsg = "公司別"
        gilCompCode.SetFocus
        GoTo 66
    End If
    '#4022 沒有啟動多媒体號戶才要檢查服務別 By Kin 2008/07/22
    If intPara24 = 0 Then
        If Len(gilServiceType.GetCodeNo & "") = 0 Then
            strErrMsg = "服務類別"
            gilServiceType.SetFocus
            GoTo 66
        End If
    End If
    If Len(gilBankCode.GetCodeNo) = 0 Then
        If Not chkChinaBank.Value = 1 Then
            strErrMsg = "轉帳代收銀行"
            gilBankCode.SetFocus
            GoTo 66
        End If
    End If
    If Len(txtPath) = 0 Then MsgBox "請輸入檔名", vbExclamation, "提示": txtPath.SetFocus: Exit Function
    If Dir(Trim(txtPath.Text)) = "" Then MsgBox "代收資料不存在，請重新選取！", vbExclamation, "訊息！": Exit Function
    
    '**********************************************************************************************************************************************************
    '#3697 中國信託新格式　必須輸入扣款日期 By Kin 2008/01/02
    If e_CreditCardType = INCardType_chinatrustNew And gdaRealDate.GetValue = Empty Then
        MsgBox "中國信託新格式必須輸入扣款日期！", vbInformation, "訊息"
        gdaRealDate.SetFocus
        Exit Function
    End If
    '***********************************************************************************************************************************************************
    
    '*****************************************************************
    '#5494 藍新格式要判斷副檔名最後一個字元是否為R By Kin 2010/02/03
    If e_CreditCardType = INCardType_NewWeb Then
        If InStrRev(txtPath, ".") <= 0 Then
            MsgBox "格式不符！", vbInformation, "訊息"
            Exit Function
        Else
            If UCase(Right(txtPath.Text, 1)) <> "R" Then
                MsgBox "格式不符！", vbInformation, "訊息"
                Exit Function
            End If
        End If
    End If
    '*****************************************************************
    '*****************************************************************
    '#5609 增加聯合信用卡格式判斷 By Kin 2010/06/09
    If e_CreditCardType = INCardType_UnionBath Then
        If InStrRev(txtPath, ".") <= 0 Then
            MsgBox "格式不符！", vbInformation, "訊息"
            Exit Function
        Else
            If UCase(Right(txtPath.Text, 3)) <> "RSP" Then
                MsgBox "格式不符！", vbInformation, "訊息"
                Exit Function
            End If
        End If
    End If
    '****************************************************************
    
    IsDataOk = True
    Exit Function
66:
    Call MsgMustBe(strErrMsg)
   Exit Function
chkErr:
    Call ErrSub(Me.Name, "IsdataOK")
End Function


Private Sub Form_Activate()
  On Error GoTo chkErr
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    gilCompCode.SetFocus
    strSQL = "Select PrgName From " & GetOwner & "CD018 Where UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'"
    If rsTmp.State = adStateOpen Then rsTmp.Close
'    rsTmp.CursorLocation = adUseClient
'    rsTmp.Open strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
    If Not GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then Exit Sub
    If rsTmp.EOF Then
        MsgBox "請設定銀行代碼中轉帳程式名稱!", vbExclamation, "警告"
        Unload Me
    End If
  Exit Sub
chkErr:
  ErrSub Me.Name, "Form_Activate"
End Sub

Private Sub Form_Initialize()
    Dim rsCD018TCB As New ADODB.Recordset
    Call GetRS(rsCD018TCB, "SELECT CODENO FROM " & GetOwner & "CD018 WHERE  BANKID ='" & "TCB" & "'", gcnGi, adUseClient, adOpenKeyset)
    If rsCD018TCB.EOF Or rsCD018TCB.BOF Then
        lngBankCode = -1
    Else
        lngBankCode = Val(rsCD018TCB("CodeNO") & "")
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If Not ChkGiList(KeyCode) Then Exit Sub
    If KeyCode = vbKeyEscape Then Unload Me
    If KeyCode = vbKeyF2 Then
        If gdaRealDate.GetValue <> Empty Then
            If Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "") Then
                gdaRealDate.SetFocus
                Exit Sub
            End If
        End If
        If cmdOK.Enabled = True Then cmdOK.Value = True
    End If
         
    
End Sub

Private Sub Form_Load()
  On Error GoTo chkErr
    Dim intPara23  As Integer
    
    Screen.MousePointer = vbDefault
    intPara24 = Val(GetRsValue("Select Para24 From " & GetOwner & "SO043", gcnGi))
    Call subGiList
    Call getDefault
    SetCreditCardTypeCbo
    Screen.MousePointer = vbDefault
    chkChinaBank.Enabled = GetUserPriv("SO3313", "SO33131")
    'gilServiceType.ListIndex = 1
    SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱", , True
    CboAddItem
    intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043") & "")
'    If intPara23 = 2 Then gimCitemCode.Enabled = False
    If intPara23 = 2 Then gimCitemCode.IsReadOnly = True
    '#4022 判斷是否有啟用多媒体帳戶,如果有要隱藏服務別 By Kin 2008/07/22
    If intPara24 = 1 Then
        gilServiceType.Visible = False
        lblServiceType.Visible = False
        lblBillHeadFmt.Visible = True
        cboBillHeadFmt.Visible = True
        gimCitemCode.Visible = True
        '#7049 啟用多媒體帳戶 要把收費項目鎖掉 By Kin 2015/07/15
        'gimCitemCode.Enabled = False
        gimCitemCode.IsReadOnly = True
    Else
        gilServiceType.Visible = True
        lblServiceType.Visible = True
        lblBillHeadFmt.Visible = False
        cboBillHeadFmt.Visible = False
        gimCitemCode.Visible = False
    End If
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub DefineRs()
On Error GoTo chkErr
  If rsSo3314.State = adStateOpen Then rsSo3314.Close
    With rsSo3314
       .LockType = adLockOptimistic
       .Fields.Append "BankCode", adBSTR, 8
       .Fields.Append "FilePath", adBSTR, 50
       .Open
    End With
  Exit Sub
chkErr:
  ErrSub Me.Name, "DefineRS"
End Sub

Private Sub GetInitData()
  On Error Resume Next
    strFilePath = ReadGICMIS1("ErrLogPath")
    If Dir(strFilePath & "\" & GetPrgName & "In1.log") = "" Then
        txtPath.Text = ""
    Else
        If rsSo3314.State = adStateOpen Then rsSo3314.Close
        rsSo3314.Open strFilePath & "\" & strPrgName & "In1.log"
        gilBankCode.SetCodeNo rsSo3314("BankCode").Value & ""
        gilBankCode.Query_CodeNo
        txtPath.Text = rsSo3314("FilePath").Value & ""
    End If
End Sub

Private Function GetPrgName() As String
  On Error GoTo chkErr
     Dim rsTmp As New ADODB.Recordset
     Dim strSQL As String
         If chkChinaBank.Value = 1 Then
             strSQL = "Select PrgName From " & GetOwner & "CD018 Where CodeNo = " & lngBankCode
         Else
             strSQL = "Select PrgName From " & GetOwner & "CD018 Where CodeNo = " & gilBankCode.GetCodeNo
         End If
        Call OpenRecordset(rsTmp, strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False)
        
            If Len(rsTmp("PrgName") & "") = 0 Then
                strPrgName = ""
                MsgBox "請設定信用卡自動扣款銀行代碼之程式名稱!", vbCritical, "錯誤"
            Else
                strPrgName = rsTmp("PrgName")
                GetPrgName = strPrgName
            End If
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "GetPrgName")
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   If rsSo3314.State = adStateOpen Then
      rsSo3314.Close
      Set rsSo3314 = Nothing
   End If
   Call ReleaseCOM(Me)
End Sub

Private Sub gdaRealDate_Validate(Cancel As Boolean)
    On Error Resume Next
        '#3411 如果未輸入則不檢查 By Kin 2008/06/02
        If gdaRealDate.GetValue = "" Then Exit Sub
        Cancel = Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "")
End Sub

Private Sub gilBankCode_Change()
  On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
       If Len(gilBankCode.GetCodeNo) > 0 And Len(gilBankCode.GetDescription) > 0 Then
           GetInitData
       End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "gilBankCode_Change")
End Sub

Private Sub subGiList()
On Error GoTo chkErr
   SetgiList gilBankCode, "CodeNo", "Description", "CD018", , , , , 3, 24
   SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12
   SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 3, 12
  Exit Sub
chkErr:
  ErrSub Me.Name, "subGiList"
End Sub

Private Sub gilCompCode_Change()

On Error GoTo chkErr
   
    If ChgComp(gilCompCode, "SO2300", "(SO2310)") = False Then Exit Sub
    subGiList
    If Len(gilCompCode.GetCodeNo & "") = 0 Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 99
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
'    GiListFilter gilBankCode, , gilCompCode.GetCodeNo
'    gilBankCode.Filter = gilBankCode.Filter & IIf(gilBankCode.Filter = "", " Where ", " And ") & " UPPER(SUBSTR(PrgName,1,10)) = '" & "CREDITCARD" & "'"
'    gilBankCode.Filter = gilBankCode.Filter & " AND NOT CODENO =" & lngBankCode
    SetCreditCardTypeCbo
  Exit Sub
99:
  MsgMustBe (strErrFile)
  Exit Sub
chkErr:
  ErrSub Me.Name, "gilCompCode_Change"
End Sub

Private Sub gilServiceType_Change()
On Error GoTo chkErr
Dim strWhere As String
           
    If Len(gilServiceType.GetCodeNo & "") = 0 Then
       strWhere = ""
    Else
       strWhere = " Where ServiceType = '" & gilServiceType.GetCodeNo & "'"
    End If
  ''  gilPTCode.Filter = strWhere
  Exit Sub
chkErr:
  ErrSub Me.Name, "gilServiceType_Change"
End Sub

Private Sub getDefault()
  On Error GoTo chkErr
     gilCompCode.SetCodeNo garyGi(9)
     gilCompCode.Query_Description
  Exit Sub
chkErr:
  ErrSub Me.Name, "getDefault"
End Sub

Private Sub chkChinaBank_Click()
If chkChinaBank.Value = 1 Then
   gilBankCode.Clear
   gilBankCode.Enabled = 0
   fraTCB.Enabled = True
Else
   gilBankCode.Enabled = 1
   fraTCB.Enabled = False
End If
End Sub
Private Sub SetCreditCardTypeCbo()
  On Error GoTo chkErr
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lbn
    Dim lngIndex As Long
    For lngIndex = 0 To cboCreditCardType.ListCount - 1
        cboCreditCardType.RemoveItem (0)
    Next
    lngIndex = 0
    strSQL = "SELECT DISTINCT PrgName FROM " & _
                    GetOwner & "CD018 WHERE UPPER(SUBSTR(PrgName,1,10)) = '" & _
                    "CREDITCARD" & "'  AND COMPCODE =" & gilCompCode.GetCodeNo
                    
    Call GetRS(rsTmp, strSQL, gcnGi, adUseClient, adOpenDynamic)
    Dim blnAdd As Boolean
    blnAdd = False
    Do While Not (rsTmp.EOF Or rsTmp.BOF)
        Select Case UCase(rsTmp("PrgName"))
            Case "CREDITCARDUNION"
                cboCreditCardType.AddItem (" 聯合/中華商銀 信用卡")
                cboCreditCardType.ItemData(lngIndex) = 0
                blnAdd = True
            Case "CREDITCARDCHINATRUST"
                cboCreditCardType.AddItem (" 中國信託 信用卡")
                cboCreditCardType.ItemData(lngIndex) = 1
                blnAdd = True
            Case "CREDITCARDHSBC"
                cboCreditCardType.AddItem (" 匯豐銀行 信用卡")
                cboCreditCardType.ItemData(lngIndex) = 2
                blnAdd = True
                
            '#3697增加一組中國信託新格式 By Kin 2007/12/31
            Case "CREDITCARDTRUSTNEW"
                cboCreditCardType.AddItem ("中國信託 信用卡新格式")
                cboCreditCardType.ItemData(lngIndex) = 3
                blnAdd = True
            '#5494 增加一組藍新科技 By Kin 2010/02/02
            Case "CREDITCARDNEWEB"
                cboCreditCardType.AddItem ("藍新科技 信用卡")
                cboCreditCardType.ItemData(lngIndex) = 4
                blnAdd = True
            '#5609 增加一組聯合信用卡批次格式 By Kin 2010/06/07
            Case "CREDITCARDUNION1"
                cboCreditCardType.AddItem (" 聯合/中華商銀 批次")
                cboCreditCardType.ItemData(lngIndex) = 5
                blnAdd = True
            Case "CREDITCARDFUBON"
            '#7765 By Kin 2018/05/09
                cboCreditCardType.AddItem (" 台北富邦")
                cboCreditCardType.ItemData(lngIndex) = 6
                blnAdd = True
         End Select
         If blnAdd = True Then lngIndex = lngIndex + 1
         blnAdd = False
         rsTmp.MoveNext
     Loop
     
    rsTmp.Close
    Set rsTmp = Nothing
    If cboCreditCardType.ListCount > 0 Then cboCreditCardType.ListIndex = 0
    Screen.MousePointer = vbDefault
Exit Sub
chkErr:
  ErrSub Me.Name, "SetCreditCardTypeCbo"
End Sub

Private Sub CboAddItem()
    On Error Resume Next
        'If cboBillHeadFmt.ListCount > 0 Then Exit Sub
        If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr From " & GetOwner & "CD068") Then Exit Sub
        Do While Not rsBillHeadFmt.EOF
            cboBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
            rsBillHeadFmt.MoveNext
        Loop
End Sub
Private Sub cboBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
    Dim strCD019SQL As String
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
         '#7049 改用CD068A.CitemCode By Kin 2015/07/09
        strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                                    " And CD068.BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1"
        'If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        If Not GetRS(rsCD019, strCD019SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        'If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        gimCitemCode.Filter = gimCitemCode.Filter
        gimCitemCode.SetQueryCode strCitemCode
End Sub

