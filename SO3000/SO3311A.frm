VERSION 5.00
Object = "{92575699-8EB8-11D3-B585-0080C8F80BC4}#4.0#0"; "GiNumber.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Begin VB.Form frmSO3311A 
   BorderStyle     =   1  '單線固定
   Caption         =   "收費單/劃撥單登錄-參數設定 [SO3311A]"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   4140
   ClientWidth     =   5625
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3311A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CheckBox chkOtpion3 
      Caption         =   "自動新增週期性項目資料"
      Enabled         =   0   'False
      Height          =   255
      Left            =   525
      TabIndex        =   11
      Top             =   5430
      Value           =   1  '核取
      Width           =   4905
   End
   Begin prjGiList.GiList gilDefSTCode 
      Height          =   315
      Left            =   1890
      TabIndex        =   7
      Top             =   3504
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
      F2Corresponding =   -1  'True
   End
   Begin prjNumber.GiNumber gnbDefRealAmt 
      Height          =   315
      Left            =   4230
      TabIndex        =   6
      Top             =   3015
      Width           =   1065
      _ExtentX        =   1879
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
      MaxLength       =   8
      AutoSelect      =   -1  'True
   End
   Begin prjNumber.GiNumber gmbDefRealAmt 
      Height          =   315
      Left            =   1890
      TabIndex        =   5
      Top             =   3020
      Width           =   915
      _ExtentX        =   1614
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
      MaxLength       =   2
      AutoSelect      =   -1  'True
   End
   Begin prjGiList.GiList gilDefPTCode 
      Height          =   315
      Left            =   1890
      TabIndex        =   4
      Top             =   2536
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
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilDefClctEn 
      Height          =   315
      Left            =   1890
      TabIndex        =   3
      Top             =   2052
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
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilDefCMCode 
      Height          =   315
      Left            =   1890
      TabIndex        =   2
      Top             =   1568
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
      F2Corresponding =   -1  'True
   End
   Begin Gi_Date.GiDate gdaDefCMDate 
      Height          =   315
      Left            =   1890
      TabIndex        =   1
      Top             =   1080
      Width           =   1130
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
   Begin Gi_YM.GiYM gymClctYM 
      Height          =   315
      Left            =   1890
      TabIndex        =   0
      Top             =   600
      Width           =   900
      _ExtentX        =   1588
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
   Begin VB.CheckBox chkOption2 
      Caption         =   "自動更新週期性項目之次收日"
      Enabled         =   0   'False
      Height          =   255
      Left            =   525
      TabIndex        =   10
      Top             =   4965
      Value           =   1  '核取
      Width           =   4905
   End
   Begin VB.CheckBox chkOption1 
      Caption         =   "自動更新週期性項目之金額"
      Height          =   255
      Left            =   525
      TabIndex        =   9
      Top             =   4500
      Value           =   1  '核取
      Width           =   4905
   End
   Begin VB.TextBox txtDefNote 
      Height          =   315
      Left            =   1890
      MaxLength       =   50
      TabIndex        =   8
      Top             =   3990
      Width           =   3540
   End
   Begin VB.CommandButton cmdFastEntry 
      Caption         =   "F8. 快速登錄..."
      Height          =   375
      Left            =   1950
      TabIndex        =   14
      Top             =   6420
      Width           =   1545
   End
   Begin VB.TextBox txtManualNo 
      Height          =   315
      Left            =   1530
      MaxLength       =   10
      TabIndex        =   12
      Top             =   5910
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   3870
      TabIndex        =   16
      Top             =   6420
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2. 確定"
      Height          =   375
      Left            =   420
      TabIndex        =   13
      Top             =   6420
      Width           =   1215
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   285
      Left            =   1890
      TabIndex        =   26
      Top             =   90
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth1       =   600
      FldWidth2       =   1500
      F5Corresponding =   -1  'True
   End
   Begin VB.Label lblComp 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   570
      TabIndex        =   27
      Top             =   180
      Width           =   765
   End
   Begin VB.Label lblDefNote 
      AutoSize        =   -1  'True
      Caption         =   "預設備註欄"
      Height          =   195
      Left            =   525
      TabIndex        =   25
      Top             =   4050
      Width           =   975
   End
   Begin VB.Label lblDefRealPeriod 
      AutoSize        =   -1  'True
      Caption         =   "預設實收期數"
      Height          =   195
      Left            =   525
      TabIndex        =   24
      Top             =   3090
      Width           =   1170
   End
   Begin VB.Label lblDefClctEn 
      AutoSize        =   -1  'True
      Caption         =   "預設收費人員"
      Height          =   195
      Left            =   525
      TabIndex        =   23
      Top             =   2130
      Width           =   1170
   End
   Begin VB.Label lblDefCMDate 
      AutoSize        =   -1  'True
      Caption         =   "預設入帳日期"
      Height          =   195
      Left            =   525
      TabIndex        =   22
      Top             =   1170
      Width           =   1170
   End
   Begin VB.Label lblDefCMCode 
      AutoSize        =   -1  'True
      Caption         =   "預設收費方式"
      Height          =   195
      Left            =   525
      TabIndex        =   21
      Top             =   1650
      Width           =   1170
   End
   Begin VB.Label lblClctYM 
      AutoSize        =   -1  'True
      Caption         =   "收費年月"
      Height          =   195
      Left            =   525
      TabIndex        =   20
      Top             =   690
      Width           =   780
   End
   Begin VB.Label lblPTCode 
      AutoSize        =   -1  'True
      Caption         =   "預設付款種類"
      Height          =   195
      Left            =   525
      TabIndex        =   19
      Top             =   2610
      Width           =   1170
   End
   Begin VB.Label lblDefSTCode 
      AutoSize        =   -1  'True
      Caption         =   "預設短收原因"
      Height          =   195
      Left            =   525
      TabIndex        =   18
      Top             =   3570
      Width           =   1170
   End
   Begin VB.Label lblDefRealAmt 
      AutoSize        =   -1  'True
      Caption         =   "預設實收金額"
      Height          =   195
      Left            =   3000
      TabIndex        =   17
      Top             =   3090
      Width           =   1170
   End
   Begin VB.Label lblManualNo 
      AutoSize        =   -1  'True
      Caption         =   "手開單號"
      Height          =   195
      Left            =   495
      TabIndex        =   15
      Top             =   5970
      Visible         =   0   'False
      Width           =   780
   End
End
Attribute VB_Name = "frmSO3311A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsSo3311 As New ADODB.Recordset
Private strTranDate As String
Private intxPara6 As Integer
Private intDayCut As Integer

Private Sub cmdCancel_Click()
    On Error Resume Next
        Unload Me
End Sub

Private Sub cmdFastEntry_Click()
    On Error GoTo chkErr
        If Not IsDataOk Then Exit Sub
        With frmSO3311E
            .uCompCode = gilCompCode.GetCodeNo
            .uYM = gymClctYM.GetValue
            .uTranDate = strTranDate
            .uNote = txtDefNote
            .uStCode = gilDefSTCode.GetCodeNo
            .uSTName = gilDefSTCode.GetDescription
            .uPeriod = gmbDefRealAmt.Value
            If Len(gnbDefRealAmt.Text) = 0 Then
                .uRealAmt = ""
            Else
                .uRealAmt = gnbDefRealAmt.Value
            End If
            .uCMCode = gilDefCMCode.GetCodeNo
            .uCMName = gilDefCMCode.GetDescription
            .uPTCode = gilDefPTCode.GetCodeNo
            .uPTName = gilDefPTCode.GetDescription
            .uRealDate = gdaDefCMDate.GetValue(True)
            .uClctEn = gilDefClctEn.GetCodeNo
            .uClctName = gilDefClctEn.GetDescription
'            .uManualNo = txtManualNo
            .Show vbModal
        End With
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdFastEntry_Click"
End Sub

Private Sub cmdOK_Click()
    On Error GoTo chkErr
        If Not IsDataOk Then Exit Sub
        With frmSO3311B
            .uPara6 = intxPara6
            .uYM = gymClctYM.GetValue
            .uTranDate = strTranDate
            .uNote = txtDefNote
            .uStCode = gilDefSTCode.GetCodeNo
            .uSTName = gilDefSTCode.GetDescription
            .uRealAmt = gnbDefRealAmt.Value2 & ""
            .uPeriod = gmbDefRealAmt.Value
            .uCMCode = gilDefCMCode.GetCodeNo
            .uCMName = gilDefCMCode.GetDescription
            .uPTCode = gilDefPTCode.GetCodeNo
            .uPTName = gilDefPTCode.GetDescription
            .uRealDate = gdaDefCMDate.GetValue(True)
            .uClctEn = gilDefClctEn.GetCodeNo
            .uClctName = gilDefClctEn.GetDescription
            .uCompCode = gilCompCode.GetCodeNo
'            .uManualNo = txtManualNo
            .uOption3 = IIf(chkOtpion3.Value = 1, True, False)
            .uOption2 = IIf(chkOption2.Value = 1, True, False)
            .uOption1 = IIf(chkOption1.Value = 1, True, False)
            .Show vbModal
        End With
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdOk_Click"
End Sub

Private Sub txtDefNote_GotFcous()
    On Error Resume Next
        txtDefNote.SelStart = 0
        txtDefNote.SelLength = Len(txtDefNote)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        'PrcList Me
        Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF2 Then
            If Not ChkCloseDate(gdaDefCMDate.GetValue(True), gilCompCode.GetCodeNo, "") Then
                gdaDefCMDate.SetFocus
                Exit Sub
            End If
            Call cmdOK_Click: Exit Sub
        End If
        If KeyCode = vbKeyF8 Then Call cmdFastEntry_Click: Exit Sub
End Sub

Private Sub Form_Load()
    On Error GoTo chkErr
        If Alfa2 = True Then
            GetGlobal
        End If
        Call subGil
        Call SetDefaultValue
    Exit Sub
chkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub gdaDefCMDate_Validate(Cancel As Boolean)
    On Error GoTo chkErr
        '若大於今日，則警告："此日期超過今天日期"
        '若（今日-本欄位值>xpara6），則警告："此日期已超過系統設定的安全期限"
        '若小於 xTranDate ，則錯誤："xTranDate 之前已做過日結，不可使用之前日期入帳",
        '且Focus停留在本欄位
        '本欄位允許空白，警告仍視為合法值
'        ChkRealDateOK gdaDefCMDate, Cancel, gilCompCode.GetCodeNo, "", strTranDate, intDayCut, intxPara6
'        Exit Sub
        Cancel = Not ChkCloseDate(gdaDefCMDate.GetValue(True), gilCompCode.GetCodeNo, "")
'        If gdaDefCMDate.Text = "" Then Exit Sub
'        If Not IsDate(Format(gdaDefCMDate.GetValue, "####/##/##")) Then
'            MsgBox "日期錯誤！", , "訊息！"
'            Cancel = True
'            Exit Sub
'        End If
'
'        If CDate(gdaDefCMDate.Text) > RightDate Then MsgBox "此日期超過今天日期": Exit Sub
'
'        If DateDiff("d", gdaDefCMDate.Text, RightDate) > intxPara6 Then
'            MsgBox "此日期已超過系統設定的安全期限", vbExclamation, "訊息！": Cancel = True: Exit Sub
'        End If
'        If CDate(gdaDefCMDate.Text) <= IIf(strTranDate <> "", CDate(strTranDate), 0) Then
'            If intDayCut = 1 And CDate(gdaDefCMDate.Text) = IIf(strTranDate <> "", CDate(strTranDate), 0) Then Exit Sub
'
'            MsgBox strTranDate & "之前已做過日結，不可使用之前日期入帳", vbExclamation, "訊息！"
'            Cancel = True
'            Exit Sub
'        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "gdaDefCMDate_Validate"
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3310", "SO3311") Then Exit Sub
        Call subGil
        Call GiListFilter(gilDefClctEn, , gilCompCode.GetCodeNo)
        Call SetDefaultValue
End Sub

Private Sub gilDefSTCode_Validate(Cancel As Boolean)
    On Error GoTo chkErr
        '若本欄有值，則根據代碼值取CD016中該代碼的參考號，若參考號為1，則將自動更新
        '週期項目之金額與自動更新週期性項目之次收日二項設為Not Checked
        chkOption2.Value = 1
        If gilDefSTCode.GetCodeNo <> "" Then
            If Not GetRS(rsSo3311, "Select RefNo from " & GetOwner & "CD016 Where CodeNo=" & gilDefSTCode.GetCodeNo, gcnGi) Then Exit Sub
            If rsSo3311("RefNO").Value = 1 Then
                chkOption1.Value = 0
                chkOption2.Value = 0
                gmbDefRealAmt.Text = ""
                gnbDefRealAmt.Text = ""
            End If
            rsSo3311.Close
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "gilDefSTCode_Validate"
End Sub

Private Sub txtManualNo_GotFocus()
    On Error Resume Next
        txtManualNo.SelStart = 0
        txtManualNo.SelLength = Len(txtManualNo.Text)
End Sub


Private Sub SetDefaultValue()
    On Error GoTo chkErr
        '設預值！
        '1、預設入帳日期：今日
        '2、預設收費日=1（人工）3、預設付款種類=1（現金）
        '4、自動更新週期性項目之金額=not Checked
        '5、自動更新週期性項目之次收日=Checked
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        
        gdaDefCMDate.Text = RightDate
        chkOption1.Value = 1
        chkOption2.Value = 1
        strTranDate = GetRsValue("Select tranDate from " & GetOwner & "So062 Where CompCode= " & gCompCode & " And Type=1 Order By TranDate Desc") & ""
        If strTranDate = "" Then strTranDate = "1990/01/01"
        If Not GetSystemPara(rsSo3311, gCompCode, gServiceType, "SO043", "Para6,Para8") Then Exit Sub
        
        intDayCut = GetSystemParaItem("DayCut", gilCompCode.GetCodeNo, "", "SO041", , True)
        
        chkOption1.Value = IIf(Val(rsSo3311("Para8") & "") = 0, 0, 1)
        intxPara6 = rsSo3311("Para6").Value
        gilDefCMCode.ListIndex = 1
        gilDefPTCode.ListIndex = 1
        Call CloseRecordset(rsSo3311)
    Exit Sub
chkErr:
    ErrSub Me.Name, "SetDefaultValue"
End Sub

Private Sub subGil()
    On Error GoTo chkErr
        '預設收費方式
        SetgiList gilDefCMCode, "CodeNo", "Description", "CD031", , , , , 3, 12, , True
        '預設收費人員
        SetgiList gilDefClctEn, "EmpNO", "EmpName", "CM003", , , , , 10, 20, , True
        '預設付款種類
        SetgiList gilDefPTCode, "CodeNo", "Description", "CD032", , , , , 3, 12, , True
        '預設短收原因
        SetgiList gilDefSTCode, "CodeNo", "Description", "CD016", , , , , 3, 12, , True
        '公司別
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode)
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Function IsDataOk() As Boolean
    On Error GoTo chkErr
        IsDataOk = False
        If Not ChkDTok Then Exit Function
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        If Not MustExist(gdaDefCMDate, 1, "預設入帳日期") Then Exit Function
        If Not MustExist(gilDefCMCode, 2, "預設收費方式") Then Exit Function
        If Not MustExist(gilDefPTCode, 2, "預設付款種類") Then Exit Function
        
        If gilDefSTCode.GetCodeNo <> "" Then
            Call gilDefSTCode_Validate(False)
        End If
        IsDataOk = True
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function
