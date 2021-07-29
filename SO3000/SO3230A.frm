VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO3230A 
   BorderStyle     =   1  '單線固定
   Caption         =   "收費單統計表 [SO3230A]"
   ClientHeight    =   5235
   ClientLeft      =   1950
   ClientTop       =   1530
   ClientWidth     =   8235
   FillColor       =   &H80000012&
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3230A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin CS_Multi.CSmulti gimCitemCode 
      Height          =   315
      Left            =   90
      TabIndex        =   32
      Top             =   1740
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   556
      ButtonCaption   =   "收費項目"
      IsReadOnly      =   -1  'True
   End
   Begin VB.ComboBox cboBillHeadFmt 
      Height          =   315
      Left            =   2070
      Style           =   2  '單純下拉式
      TabIndex        =   8
      Top             =   1410
      Visible         =   0   'False
      Width           =   6015
   End
   Begin Gi_Multi.GiMulti gimCustStatusCode 
      Height          =   345
      Left            =   90
      TabIndex        =   9
      Top             =   2100
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   609
      ButtonCaption   =   "客戶狀態"
      FontSize        =   9
      FontName        =   "新細明體"
   End
   Begin Gi_Multi.GiMulti gimOrder 
      Height          =   345
      Left            =   90
      TabIndex        =   13
      Top             =   3540
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   609
      ButtonCaption   =   "排序方式"
      Enabled         =   0   'False
      ColumnOrder     =   -1  'True
      FontSize        =   9
      FontName        =   "新細明體"
   End
   Begin CS_Multi.CSmulti gimClctEn 
      Height          =   345
      Left            =   90
      TabIndex        =   11
      Top             =   2820
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   609
      ButtonCaption   =   "收費人員"
   End
   Begin VB.CheckBox chkNoPrint 
      Caption         =   "未列印資料"
      Height          =   195
      Left            =   5100
      TabIndex        =   7
      Top             =   990
      Value           =   1  '核取
      Width           =   1515
   End
   Begin VB.CheckBox chkHavePrint 
      Caption         =   "已列印資料"
      Height          =   195
      Left            =   5100
      TabIndex        =   6
      Top             =   570
      Value           =   1  '核取
      Width           =   1515
   End
   Begin VB.CheckBox chkIncClose 
      Caption         =   "包含日結資料"
      Height          =   195
      Left            =   2940
      TabIndex        =   3
      Top             =   570
      Width           =   1515
   End
   Begin MSMask.MaskEdBox mskPrtSNO1 
      Height          =   345
      Left            =   1050
      TabIndex        =   4
      Top             =   960
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "999999999999"
      PromptChar      =   " "
   End
   Begin VB.Frame fraPrintChoose 
      BackColor       =   &H00E0E0E0&
      Caption         =   "列印內容"
      ForeColor       =   &H00800080&
      Height          =   585
      Left            =   90
      TabIndex        =   29
      Top             =   4050
      Width           =   3375
      Begin VB.OptionButton optSummarics 
         BackColor       =   &H00E0E0E0&
         Caption         =   "彙總表"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   1635
         TabIndex        =   15
         Top             =   270
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton optDetail 
         BackColor       =   &H00E0E0E0&
         Caption         =   "明細表"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Frame fraCalculate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "收費狀況"
      ForeColor       =   &H00800080&
      Height          =   585
      Left            =   3570
      TabIndex        =   28
      Top             =   4050
      Width           =   4575
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00E0E0E0&
         Caption         =   "全部"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   3330
         TabIndex        =   19
         Top             =   270
         Value           =   1  '核取
         Width           =   705
      End
      Begin VB.CheckBox chkUCCode 
         BackColor       =   &H00E0E0E0&
         Caption         =   "未收"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   270
         Value           =   1  '核取
         Width           =   765
      End
      Begin VB.CheckBox chkRealDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "已收"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   1230
         TabIndex        =   17
         Top             =   270
         Value           =   1  '核取
         Width           =   765
      End
      Begin VB.CheckBox chkCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "作廢"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   2070
         TabIndex        =   18
         Top             =   270
         Value           =   1  '核取
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   4710
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   495
      Left            =   6750
      TabIndex        =   22
      Top             =   4710
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   495
      Left            =   1800
      TabIndex        =   21
      Top             =   4710
      Width           =   1395
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1050
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
      Left            =   5100
      TabIndex        =   1
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
   Begin Gi_YM.GiYM gymPrtYM 
      Height          =   345
      Left            =   1050
      TabIndex        =   2
      Top             =   510
      Width           =   1035
      _ExtentX        =   1826
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
   Begin MSMask.MaskEdBox mskPrtSNO2 
      Height          =   345
      Left            =   2940
      TabIndex        =   5
      Top             =   960
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   609
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "999999999999"
      PromptChar      =   " "
   End
   Begin Gi_Multi.GiMulti gimCMCode 
      Height          =   345
      Left            =   90
      TabIndex        =   10
      Top             =   2460
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   609
      ButtonCaption   =   "收費方式"
      FontSize        =   9
      FontName        =   "新細明體"
   End
   Begin CS_Multi.CSmulti gimOldClctEn 
      Height          =   345
      Left            =   90
      TabIndex        =   12
      Top             =   3180
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   609
      ButtonCaption   =   "原收費人員"
   End
   Begin VB.Label lblBillHeadFmt 
      AutoSize        =   -1  'True
      Caption         =   "多帳戶產生依據設定"
      Height          =   195
      Left            =   240
      TabIndex        =   31
      Top             =   1470
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label lblPrtYM 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "印單年月"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   600
      Width           =   780
   End
   Begin VB.Label lblPrtSNo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "印單序號"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   1005
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "~"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   2775
      TabIndex        =   26
      Top             =   1035
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "服務類別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4170
      TabIndex        =   25
      Top             =   150
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   150
      Width           =   765
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   525
      Left            =   2520
      TabIndex        =   23
      Top             =   -1470
      Width           =   1245
   End
End
Attribute VB_Name = "frmSO3230A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBillHeadFmt As New ADODB.Recordset
Dim strViewName As String

Private Sub cboBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
    Dim strCD019SQL As String
    If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
    rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
    '#7127 It was changed field source  from CD068A field By Kin 2015/11/05
    strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
                          " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                          "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                                  " And CD068.BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1"
    'If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
    If Not GetRS(rsCD019, strCD019SQL) Then Exit Sub
    If rsCD019.EOF Then Exit Sub
    strCitemCode = rsCD019.GetString(, , , ",")
    strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
    gimCitemCode.Filter = gimCitemCode.Filter
    gimCitemCode.SetQueryCode strCitemCode
        'gimCitemCode.ShowMulti
End Sub

Private Sub chkAll_Click()
    On Error GoTo chkErr
        If chkAll = 1 Then
            chkCancel.Value = 1
            chkRealDate.Value = 1
            chkUCCode.Value = 1
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "chkAll_Click"
End Sub

Private Sub chkCancel_Click()
    On Error Resume Next
        If chkCancel.Value = 1 And chkRealDate.Value = 1 And chkUCCode.Value = 1 Then
            chkAll.Value = 1
        Else
            chkAll.Value = 0
        End If
End Sub

Private Sub chkRealDate_Click()
    On Error Resume Next
        If chkCancel.Value = 1 And chkRealDate.Value = 1 And chkUCCode.Value = 1 Then
            chkAll.Value = 1
        Else
            chkAll.Value = 0
        End If
End Sub

Private Sub chkUCCode_Click()
    On Error Resume Next
        If chkCancel.Value = 1 And chkRealDate.Value = 1 And chkUCCode.Value = 1 Then
            chkAll.Value = 1
        Else
            chkAll.Value = 0
        End If
End Sub

Private Sub cmdExit_Click()
  On Error GoTo chkErr
    Unload Me
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo chkErr
    Screen.MousePointer = vbHourglass
    Select Case True
         Case optDetail.Value
             Call PreviousRpt(GetPrinterName(5), "SO3230A1.Rpt", "收費單統計表(明細表) [SO3230A]")
         Case optSummarics.Value
             Call PreviousRpt(GetPrinterName(5), "SO3230A2.Rpt", "收費單統計表(彙總表) [SO3230A]")
    End Select
    Screen.MousePointer = vbDefault
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo chkErr
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    Screen.MousePointer = vbHourglass
    cmdExit.SetFocus
    Me.Enabled = False
    ReadyGoPrint
    If Not subChoose Then Exit Sub
    Select Case True
           Case optDetail '明細表
                Call subDetail
           Case optSummarics '彙總表
                Call subSummarics
    End Select
    Me.Enabled = True
    Screen.MousePointer = vbDefault
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean
    On Error GoTo chkErr
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        If Not MustExist(gymPrtYM, 1, "印單年月") Then Exit Function
        If chkHavePrint = 0 And chkNoPrint = 0 Then
            MsgBox "已列印或未列印至少需選擇一項!!", vbExclamation, gimsgPrompt
            chkHavePrint.SetFocus
            Exit Function
        End If
        IsDataOk = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Function subChoose() As Boolean
    On Error GoTo chkErr
    Dim strPrint As String
    Dim strCalculate As String
    Dim strChCalculate As String
        strChoose = ""
        strChooseString = ""
        strPrint = ""
        strGroupName = ""
   
        '印單序號
        If mskPrtSNo1.Text <> "" Then Call subAnd("A.PrtSNo >= '" & Trim(mskPrtSNo1.Text) & "'")
        If mskPrtSNo2.Text <> "" Then Call subAnd("A.PrtSNo <= '" & Trim(mskPrtSNo2.Text) & "'")
        '印單年月
        If gymPrtYM.GetValue <> "" Then Call subAnd("A.PrtSNo >= '" & gymPrtYM.GetValue & "' And A.PrtSNo < '" & gymPrtYM.GetValue + 1 & "'")
        
        'GiMulti
        If gimCustStatusCode.GetQryStr <> "" Then Call subAnd("D.CustStatusCode " & gimCustStatusCode.GetQryStr)
        If gimClctEn.GetQryStr <> "" Then Call subAnd("A.ClctEn " & gimClctEn.GetQryStr)
        If gimCMCode.GetQryStr <> "" Then Call subAnd("A.CMCode " & gimCMCode.GetQryStr)
        If gimCitemCode.GetQryStr <> "" Then Call subAnd("A.CitemCode " & gimCitemCode.GetQryStr)
        If gimOldClctEn.GetQryStr <> "" Then Call subAnd("A.OldClctEn " & gimOldClctEn.GetQryStr)
        
        'GiList
        If gilCompCode.GetCodeNo <> "" Then Call subAnd("A.CompCode=" & gilCompCode.GetCodeNo)
        If gilServiceType.GetCodeNo <> "" Then Call subAnd("A.ServiceType='" & gilServiceType.GetCodeNo & "'")
        If Not (chkHavePrint.Value = 1 And chkNoPrint.Value = 1) Then
            If chkHavePrint.Value = 1 Then Call subAnd("Nvl(A.PrtCount,0) > 0")
            If chkNoPrint.Value = 1 Then Call subAnd("Nvl(A.PrtCount,0) = 0")
        End If
        
        
        If chkAll Then
            strCalculate = chkAll.Caption
        Else
            If chkUCCode.Value = 1 Then
                strChCalculate = " Or (A.UCCode Is Not Null ) "
                strCalculate = "," & chkUCCode.Caption
            End If
            If chkRealDate = 1 Then
                strChCalculate = strChCalculate & " Or (A.UCCode Is Null And A.RealDate Is Not Null And A.CancelFlag = 0)  "
                strCalculate = strCalculate & "," & chkRealDate.Caption
            End If
            If chkCancel = 1 Then
                strChCalculate = strChCalculate & " OR A.CancelFlag = 1 "
                strCalculate = strCalculate & "," & chkCancel.Caption
            End If
            If strChCalculate <> "" Then subAnd "(" & Mid(strChCalculate, 4) & ")"
            If strCalculate <> "" Then strCalculate = Mid(strCalculate, 2)
        End If
        Select Case True
            Case optDetail.Value
                strPrint = "明細表;排序方式:" & gimOrder.GetColumnOrderDspStr
                
            Case optSummarics.Value
                strPrint = "彙總表"
        End Select
        If gimOrder.GetQryStr <> "" Then Call SetGiMultiOrder(strGroupName, gimOrder)
        '   End If
        strChooseString = "印單年月:" & subSpace(gymPrtYM.GetValue(True)) & ";" & _
                       "印單序號:" & subSpace(Trim(mskPrtSNo1)) & "~" & subSpace(Trim(mskPrtSNo2)) & ";" & _
                       "公司別　:" & subSpace(gilCompCode.GetDescription) & ";" & _
                       "服務類別:" & subSpace(gilServiceType.GetDescription) & ";" & _
                       "收費項目:" & subSpace(gimCitemCode.GetDispStr) & ";" & _
                       "客戶狀態:" & subSpace(gimCustStatusCode.GetDispStr) & ";" & _
                       "收費方式:" & subSpace(gimCMCode.GetDispStr) & ";" & _
                       "收費人員: " & subSpace(gimClctEn.GetDispStr) & ";" & _
                       "收費狀況:" & strCalculate & ";" & _
                       "包含日結資料:" & IIf(chkIncClose.Value = 1, "是", "否") & ";" & _
                       "已列印資料:" & IIf(chkHavePrint.Value = 1, "是", "否") & ";" & _
                       "未列印資料:" & IIf(chkNoPrint.Value = 1, "是", "否") & ";" & _
                       "列印內容:" & strPrint
                       
        subChoose = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "subChoose")
End Function

Private Sub subDetail()
    On Error GoTo chkErr
    Dim lngAffected As Long
        If Not subInsertMDB(lngAffected) Then Exit Sub
        If lngAffected = 0 Then
            MsgNoRcd
            SendSQL , , True
        Else
            'strSQL = "Select * From So3230A1 So3230A"
            strSQL = "Select * From " & strViewName & " V"
            'Call PrintRpt(GetPrinterName(5), RptName("SO3230", "1"),  , "收費單統計表(明細表) [SO3230A]", strSQL, strChooseString, " 印單序號", True, , , , GiPaperLandscape)
            Call PrintRpt(GetPrinterName(5), "SO3230A1.rpt", , "收費單統計表(明細表) [SO3230A]", strSQL, strChooseString, , True, , , strGroupName)
        End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subDetail")
End Sub

Private Sub subSummarics()
    On Error GoTo chkErr
    Dim lngAffected As Long
        If Not subInsertMDB2(lngAffected) Then Exit Sub
        If lngAffected = 0 Then
            MsgNoRcd
            SendSQL , , True
        Else
            strSQL = "Select * From SO3230A2"
           'Call PrintRpt(GetPrinterName(5), RptName("SO3230", "2"), , "收費單統計表(彙總表) [SO3230A]", strSQL, strChooseString, , True, "Tmp0000.Mdb")
           Call PrintRpt(GetPrinterName(5), "SO3230A2.Rpt", , "收費單統計表(彙總表) [SO3230A]", strSQL, strChooseString, , True, "Tmp0000.Mdb")
        End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subSummarics")
End Sub

Private Function subInsertMDB2(ByRef lngAffected As Long) As Boolean
    On Error GoTo chkErr
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim lngBR As Long, strSQL As String
    Dim strPrtSNo1 As String, lngCountPrtSNo As Long
    Dim lngTotal As Long
    Dim blnCutPrtSNo As Boolean
    Dim sngTotalAmt As Double
        If Not GetRS(rs1, "Select ClctEn,ClctName,CancelFlag,UCCode,PrtSNo PrtSNo1,PrtSNO PrtSNO2 ,0 lngCount ,0 CountPrtSNo,0 TotalAmt From " & GetOwner & "So033 Where 1=0") Then Exit Function
        If Not CreateMDBTable(rs1, "SO3230A2", cnn) Then Exit Function
        If Not subQry2(strSQL, "SO033") Then Exit Function
        If chkIncClose.Value = 1 Then
            If Not subQry2(strSQL, "SO034") Then Exit Function
            If Not subQry2(strSQL, "SO035") Then Exit Function
        End If
        strSQL = "Select * From (" & Mid(strSQL, 11) & ") Order By PrtSNo"
        If Not GetRS(rs1, strSQL) Then Exit Function
        cnn.BeginTrans
        SendSQL rs1.Source, True
        Set rs2 = rs1.Clone
        If Not (rs1.BOF Or rs1.EOF) Then
            rs2.MoveNext
            strPrtSNo1 = rs1("PrtSNo")
            sngTotalAmt = 0
            lngCountPrtSNo = 1
            lngTotal = 1
            Do While Not rs2.EOF
                '斷號
                If ChkCutPrtSNO(rs1, rs2) Then
                    sngTotalAmt = sngTotalAmt + rs1("ShouldAmt")
                    If Not UpdateMDB2(rs1, strPrtSNo1, lngCountPrtSNo, lngTotal, 1, False, sngTotalAmt) Then cnn.RollbackTrans: Exit Function
                    '新增一筆斷號
                    'If Not UpdateMDB2(rs2, strPrtSNo1, lngCountPrtSNo, lngTotal, 1, True, sngTotalAmt) Then cnn.RollbackTrans: Exit Function
                    'If chkAll = 1 Then If Not UpdateMDB2(rs2, rs1("PrtSNo") + 1, Val(rs2("PrtSNo") - rs1("PrtSNo")) - 1, Val(rs2("PrtSNo") - rs1("PrtSNo")) - 1, 0, True) Then Exit Function
                    strPrtSNo1 = rs2("PrtSNo")
                    sngTotalAmt = 0
                    lngTotal = 1
                    lngCountPrtSNo = 1
                '檢查是否有改變
                ElseIf ChkCanExecuteUpdate(rs1, rs2) Then
                    sngTotalAmt = sngTotalAmt + rs1("ShouldAmt")
                    If Not UpdateMDB2(rs1, strPrtSNo1, lngCountPrtSNo, lngTotal, 1, False, sngTotalAmt) Then cnn.RollbackTrans: Exit Function
                    strPrtSNo1 = rs2("PrtSNo")
                    lngTotal = 1
                    sngTotalAmt = 0
                    lngCountPrtSNo = 1
                Else
                    If rs1("PrtSNo") <> rs2("PrtSNo") Then
                        lngCountPrtSNo = lngCountPrtSNo + 1
                    End If
                    sngTotalAmt = sngTotalAmt + rs1("ShouldAmt")
                    lngTotal = lngTotal + 1
                End If
                rs1.MoveNext
                rs2.MoveNext
                DoEvents
            Loop
            sngTotalAmt = sngTotalAmt + rs1("ShouldAmt")
            If Not UpdateMDB2(rs1, strPrtSNo1, lngCountPrtSNo, lngTotal, , False, sngTotalAmt) Then cnn.RollbackTrans: Exit Function
        End If
        lngAffected = rs1.RecordCount
        Call Close3Recordset(rs1)
        cnn.CommitTrans
        subInsertMDB2 = True
        
    Exit Function
chkErr:
    cnn.RollbackTrans
    ErrSub Me.Name, "subInsertMDB2"
End Function

Private Function ChkCanCutSNo() As Boolean
    On Error Resume Next
        
End Function

Private Function subQry2(ByRef strSQL As String, _
    strTable As String) As Boolean
    On Error GoTo chkErr
    Dim strWhere As String
    Dim strForm As String
'        If InStr(strChoose, "B.") > 0 Then
'            strForm = ", So001 A "
'            strWhere = " A.CustId = B.CustId And "
'        End If
        If InStr(strChoose, "D.") > 0 Then
            strForm = ", " & GetOwner & "So002 D "
            strWhere = " A.CustId = D.CustId And A.ServiceType = D.ServiceType And "
        End If
        strSQL = strSQL & " Union All Select " & GetUseIndexStr(strTable, "PrtSNo") & " ClctEn,ClctName,Nvl(CancelFlag,0) CancelFlag,Decode(UCCode,Null,0,1) UCCode,PrtSNo,ShouldAmt From " & GetOwner & strTable & " A " & strForm & " Where " & strWhere & strChoose
        '暫不使用收費方式
'        strSQL = strSQL & "Union All Select ClctEn,ClctName,Nvl(CancelFlag,0) CancelFlag,Decode(UCCode,Null,0,1) UCCode,CMName,PrtSNo From " & strTable & " A Where " & strChoose
        subQry2 = True
    Exit Function
chkErr:
    ErrSub Me.Name, "subQry2"
End Function

Private Function UpdateMDB2(rs1 As ADODB.Recordset, strPrtSNo1 As String, _
    lngCountPrtSNo As Long, lngTotal As Long, _
    Optional lngLast As Long = 0, Optional blnCutSNo As Boolean = False, _
    Optional sngTotalAmt As Double) As Boolean
    On Error GoTo chkErr
    Dim blnMove As Boolean
        cnn.Execute "Insert Into SO3230A2 (ClctEn,ClctName,CancelFlag,UCCode,PrtSNo1,PrtSNo2,lngCount,CountPrtSNo,TotalAmt) Values (" & _
            GetNullString(rs1("ClctEn")) & "," & GetNullString(rs1("ClctName")) & "," & _
            GetNullString(rs1("CancelFlag")) & "," & GetNullString(IIf(blnCutSNo, -1, rs1("UCCode"))) & "," & _
            GetNullString(strPrtSNo1) & "," & GetNullString(IIf(blnCutSNo, rs1("PrtSNo") - 1, rs1("PrtSNo"))) & "," & _
            lngTotal & "," & lngCountPrtSNo & "," & IIf(blnCutSNo, 0, sngTotalAmt) & ")"
        UpdateMDB2 = True
    Exit Function
chkErr:
    ErrSub Me.Name, "UpdateMDB2"
End Function

Private Function ChkCutPrtSNO(rs1 As ADODB.Recordset, _
    rs2 As ADODB.Recordset) As Boolean
    On Error GoTo chkErr
    Dim strPrtSNo As String
        ChkCutPrtSNO = False
        If rs2("PrtSNo") - rs1("PrtSNo") >= 2 Then
            ChkCutPrtSNO = True
            Exit Function
        End If
    Exit Function
chkErr:
    ErrSub Me.Name, "ChkCutPrtSNo"
End Function

Private Function ChkCanExecuteUpdate(rs1 As ADODB.Recordset, _
    rs2 As ADODB.Recordset) As Boolean
    On Error GoTo chkErr
    Dim lngLoop As Long
        ChkCanExecuteUpdate = False
        For lngLoop = 0 To rs1.Fields.Count - 1
            If rs1.Fields(lngLoop) <> rs2.Fields(rs1.Fields(lngLoop).Name) And UCase(rs1.Fields(lngLoop).Name <> "PRTSNO") And UCase(rs1.Fields(lngLoop).Name <> "SHOULDAMT") Then
                ChkCanExecuteUpdate = True
                Exit Function
            End If
        Next
        ChkCanExecuteUpdate = False
    Exit Function
chkErr:
    ErrSub Me.Name, "ChkCanExecuteUpdate"
End Function

Private Function subInsertMDB(ByRef lngAffected As Long) As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMduName As String
        If Not subQry1(strSQL, "SO033") Then Exit Function
        If chkIncClose.Value = 1 Then
            If Not subQry1(strSQL, "SO034") Then Exit Function
            If Not subQry1(strSQL, "SO035") Then Exit Function
        End If
        strSQL = "Select * From (" & Mid(strSQL, 11) & ") Order By PrtSNo"
        
        SendSQL rsTmp.Source, True
        On Error Resume Next
          gcnGi.Execute "Drop View " & strViewName
        On Error GoTo chkErr
        strViewName = GetTmpViewName
        SendSQL strSQL
        gcnGi.Execute "Create View " & strViewName & " as " & strSQL
        
        'If Not GetRS(rsTmp, strSQL) Then Exit Function
'        If Not CreateMDBTable(rsTmp, "SO3230A1", cnn) Then Exit Function
'        cnn.BeginTrans
'        While Not rsTmp.EOF
'            cnn.Execute "Insert Into SO3230A1 (PrtSNO,CustId,CitemName,CustName,Address,Tel1,CustStatusName,MduId,MduName,RealPeriod,ShouldAmt,RealStartDate,RealStopDate,ClctEn,ClctName,TypeName,AddrSort,PrtCount,MediaBillNo,OldClctEn,OldClctName) Values (" & _
'                 GetNullString(rsTmp("PrtSNo")) & "," & GetNullString(rsTmp("CustId")) & "," & _
'                 GetNullString(rsTmp("CitemName")) & "," & GetNullString(rsTmp("CustName")) & "," & _
'                 GetNullString(rsTmp("Address")) & "," & GetNullString(rsTmp("Tel1")) & "," & _
'                 GetNullString(rsTmp("CustStatusName")) & "," & GetNullString(rsTmp("MduId")) & "," & _
'                 GetNullString(rsTmp("MduName")) & "," & GetNullString(rsTmp("RealPeriod")) & "," & _
'                 GetNullString(rsTmp("ShouldAmt")) & "," & GetNullString(rsTmp("RealStartDate"), giDateV, giAccessDb) & "," & _
'                 GetNullString(rsTmp("RealStopDate"), giDateV, giAccessDb) & "," & GetNullString(rsTmp("ClctEn")) & "," & _
'                 GetNullString(rsTmp("ClctName")) & "," & GetNullString(rsTmp("TypeName")) & "," & _
'                 GetNullString(rsTmp("AddrSort")) & "," & GetNullString(rsTmp("PrtCount")) & "," & _
'                 GetNullString(rsTmp("MediaBillNo")) & "," & GetNullString(rsTmp("OldClctEn")) & "," & _
'                 GetNullString(rsTmp("OldClctName")) & ")"
'           rsTmp.MoveNext
'           DoEvents
'        Wend
'        cnn.CommitTrans
        lngAffected = GetRsValue("Select Count(*) From " & strViewName)
        Call Close3Recordset(rsTmp)
        subInsertMDB = True
    Exit Function
chkErr:
'    cnn.RollbackTrans
    Call ErrSub(Me.Name, "subInsertMDB")
End Function

Private Function subQry1(ByRef strSQL As String, _
    strTable As String) As Boolean
    On Error GoTo chkErr
    Dim strWhere As String
    Dim strForm As String
'        If InStr(strChoose, "D.") > 0 Then
            strForm = ", " & GetOwner & "So002 D "
            strWhere = " And A.CustId = D.CustId And A.CompCode = D.CompCode And A.ServiceType = D.ServiceType "
'        End If
        strSQL = strSQL & " Union All Select " & GetUseIndexStr(strTable, "PrtSNo") & " A.PrtSNO,A.CustId,A.CitemName,B.CustName ,C.Address,B.Tel1,D.CustStatusName,A.MduId, " & _
            "E.Name as MduName,A.RealPeriod,A.ShouldAmt,A.RealStartDate,A.RealStopDate,A.ClctEn,A.ClctName,'" & strTable & "' as TypeName,C.AddrSort,Nvl(A.PrtCount,0) PrtCount,A.MediaBillNo,A.OldClctEn,A.OldClctName From " & _
            GetOwner & "So001 B," & GetOwner & "SO014 C," & GetOwner & "SO017 E" & strForm & "," & GetOwner & strTable & " A Where A.CustId =B.CustId And A.CompCode=B.CompCode And A.AddrNo=C.AddrNo And A.CompCode=C.CompCode And A.MduId = E.MduId(+) And A.CompCode = E.CompCode(+) " & strWhere & _
            "And " & strChoose
        subQry1 = True
    Exit Function
chkErr:
    ErrSub Me.Name, "subQry1"
End Function

Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo chkErr
    Call FunctionKey(KeyCode, Shift, Me)
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
    On Error GoTo chkErr
        Call InitData
        Call subGil
        Call subGim
        Call CboAddItem
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        Set cnn = GetTmpMdbCn
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub InitData()
    On Error Resume Next
    Dim intPara23 As Integer, intPara24 As Integer
        intPara24 = GetSystemParaItem("Para24", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043")
        intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043") & "")
        If intPara24 = 1 Then
            lblBillHeadFmt.Visible = True
            cboBillHeadFmt.Visible = True
            '  用媒體入帳才Lock
             'If intPara23 = 2 Then gimCitemCode.Enabled = False
        End If
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

Private Sub subGil()
    On Error GoTo chkErr
        Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
        Call SetgiList(gilServiceType, "CodeNo", "Description", "CD046")
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub subGim()
    On Error GoTo chkErr
        Call SetgiMulti(gimClctEn, "EmpNo", "EmpName", "CM003", "收費員代號", "收費員名稱")
        Call SetgiMulti(gimOldClctEn, "EmpNo", "EmpName", "CM003", "原收費員代號", "原收費員名稱")
        Call SetgiMulti(gimCustStatusCode, "CodeNo", "Description", "CD035", "代號", "名稱")
        Call SetgiMulti(gimCitemCode, "CodeNo", "Description", "CD019", "代號", "名稱")
        Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "代號", "名稱", , True)
        SetgiMultiAddItem gimOrder, "A,B,C,D", "地址,收費人員,客戶編號,印單序號"
        gimCustStatusCode.SetQueryCode "1"
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM Me
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo chkErr
    If Not ChgComp(gilCompCode, "SO3200", "SO3230") Then Exit Sub
    Call subGil
    Call subGim
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.SetCodeNo ""
    gilServiceType.Query_Description
    gilServiceType.ListIndex = 1
  Exit Sub
  
chkErr:
  Call ErrSub(Me.Name, "gilCompCode_Change")

End Sub

Private Sub gilCompCode_Validate(Cancel As Boolean)
  On Error GoTo chkErr
    If gilCompCode.GetCodeNo = "" Then strErrFile = "公司別": GoTo 99
  Exit Sub
99:
  MsgMustBe (strErrFile)
  Cancel = True
  Exit Sub
chkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Sub gilServiceType_Change()
    On Error Resume Next
        Call GiMultiFilter(gimCMCode, gilServiceType.GetCodeNo)

End Sub

Private Sub mskPrtSNO1_Change()
    On Error Resume Next
        mskPrtSNo2 = mskPrtSNo1
End Sub

Private Sub SetGiMultiOrder(ByRef strFormulaName As String, _
        gimOrder As Object, Optional blnUserGroupName As Boolean = False)
    On Error GoTo chkErr
      '排序方式
    Dim varCollect As Variant
    Dim varSplit As Variant
    Dim strOrderName As String
    Dim lngLoop As Long

        varCollect = Split(Replace(gimOrder.GetColumnOrderCode, "'", ""), ",")
        lngLoop = 0
        For Each varSplit In varCollect
            Select Case Trim(varSplit)
                Case "A"
                    strOrderName = "{V.AddrSort}"
                Case "B"
                    strOrderName = "{V.ClctEn}"
                Case "C"
                    strOrderName = "{V.CustId}"
                Case "D"
                    strOrderName = "{V.PrtSNo}"
            End Select
            If lngLoop = 0 And blnUserGroupName Then strFormulaName = strFormulaName & IIf(strFormulaName = "", "", ";") & "GroupName=" & strOrderName
            Call ReturnAndString(strFormulaName, "Sort" & Chr(Asc("A") + lngLoop) & "=" & strOrderName, ";")
            lngLoop = lngLoop + 1
        Next
        Call ReturnAndString(strFormulaName, "Sort" & Chr(Asc("A") + lngLoop) & "={V.CustId}", ";")
        
    Exit Sub
chkErr:
    ErrSub Me.Name, "SetGiMultiOrder"
End Sub

Private Sub optDetail_Click()
    gimOrder.Enabled = True
End Sub

Private Sub optSummarics_Click()
    gimOrder.Enabled = False
End Sub
