VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO3410A 
   BorderStyle     =   1  '單線固定
   Caption         =   "收費單未收原因登錄 [SO3410A]"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   4620
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
   Icon            =   "SO3410A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin Gi_Date.GiDate gdaRealDate 
      Height          =   375
      Left            =   1980
      TabIndex        =   3
      Top             =   1410
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
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
   Begin prjGiList.GiList gilUCCode 
      Height          =   315
      Left            =   1980
      TabIndex        =   4
      Top             =   1890
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
   Begin prjGiList.GiList gilClctEn 
      Height          =   315
      Left            =   1980
      TabIndex        =   2
      Top             =   990
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
   Begin prjGiList.GiList gilCMCode 
      Height          =   315
      Left            =   1980
      TabIndex        =   1
      Top             =   540
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
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2.確定"
      Height          =   375
      Left            =   885
      TabIndex        =   5
      Top             =   2460
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   3045
      TabIndex        =   7
      Top             =   2460
      Width           =   1335
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   1980
      TabIndex        =   0
      Top             =   90
      Width           =   2430
      _ExtentX        =   4286
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
      Left            =   660
      TabIndex        =   11
      Top             =   180
      Width           =   765
   End
   Begin VB.Label lblUCCode 
      AutoSize        =   -1  'True
      Caption         =   "預設未收原因"
      Height          =   195
      Left            =   630
      TabIndex        =   10
      Top             =   1920
      Width           =   1170
   End
   Begin VB.Label lblRealDate 
      AutoSize        =   -1  'True
      Caption         =   "預設登錄日期"
      Height          =   195
      Left            =   630
      TabIndex        =   9
      Top             =   1470
      Width           =   1170
   End
   Begin VB.Label lblClctEn 
      AutoSize        =   -1  'True
      Caption         =   "預設收費人員"
      Height          =   195
      Left            =   630
      TabIndex        =   8
      Top             =   1050
      Width           =   1170
   End
   Begin VB.Label lblCMCode 
      AutoSize        =   -1  'True
      Caption         =   "預設收費方式"
      Height          =   195
      Left            =   630
      TabIndex        =   6
      Top             =   630
      Width           =   1170
   End
End
Attribute VB_Name = "frmSO3410A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private intPara6 As Long
'Private ObjopenDb As New giOpenDBObj.OpenDBObj
Public blnCanBeClose As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo chkErr
        If Not ChkDTok Then Exit Sub
        If Not IsDataOK Then Exit Sub
        With frmSO3410B
            .uPara6 = intPara6
            .uClctEn = gilClctEn.GetCodeNo
            .uClctName = gilClctEn.GetDescription
            .uCMCode = gilCMCode.GetCodeNo
            .uCMName = gilCMCode.GetDescription
            .uUCCode = gilUCCode.GetCodeNo
            .uUCname = gilUCCode.GetDescription
            .uDate = gdaRealDate.GetValue
            .uCompCode = gilCompCode.GetCodeNo
            .Show vbModal
        End With
        If blnCanBeClose = True Then
            Unload Me
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdOk_Click"
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF2 Then
            If Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "") Then
                gdaRealDate.SetFocus
                Exit Sub
            End If
            Call cmdOK_Click
        End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
        If Alfa2 Then
            Call GetGlobal
        End If
        blnCanBeClose = False
        Call subGil
        Call SetDefaultValue
End Sub

Private Function IsDataOK() As Boolean
    On Error GoTo chkErr
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        IsDataOK = True
        Exit Function
    Exit Function
chkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub subGil()
    On Error Resume Next
        '預設收費方式
        SetgiList gilCMCode, "CodeNO", "Description", "CD031", , , , , , , , True
        gilCMCode.ListIndex = 1
        
        '預設收費人員
        SetgiList gilClctEn, "EmpNO", "EmpName", "CM003", , , , , , , , True
        
        '預設未收原因
        '#3869 過濾只有RefNo=1 Or RefNo is Null By Kin 2008/06/04   '990507 #5364 未停用且未收原因參考號為空值、0、1的方可允許選擇
        SetgiList gilUCCode, "CodeNo", "Description", "CD013", , , , , , , " Where StopFlag=0 And (RefNo=0 or RefNO=1 or RefNO is Null) ", True
        
        '公司別
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode)
        
End Sub

Private Sub SetDefaultValue()
    On Error GoTo chkErr
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        '收費參數檔（SO043)，取收費日登錄期限<Para6>，供後續使用！
        intPara6 = Val(GetSystemParaItem("Para6", gilCompCode.GetCodeNo, "", "SO043") & "")
        
        '預設日期為今日！
        Call gdaRealDate.SetValue(Format(RightNow, "YYYYMMDD"))
    Exit Sub
chkErr:
    ErrSub Me.Name, "SetDefaultValue"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub gdaRealDate_Validate(Cancel As Boolean)
    On Error Resume Next
        If gdaRealDate.GetValue = "" Then Exit Sub
        Cancel = Not ChkCloseDate(gdaRealDate.GetValue(True), gilCompCode.GetCodeNo, "")
'        If gdaRealDate.Text = "" Then
'                If vbNo = MsgBox("本欄位是否為空值！", vbExclamation, "訊息！") Then
'                    Cancel = True
'                End If
'        End If
'        If Not IsDate(gdaRealDate.Text) Then
'            MsgBox "日期不合法！", vbExclamation, "訊息！"
'            Exit Sub
'        End If
'        If CDate(gdaRealDate.Text) > RightDate Then MsgBox "此日期超過今天日期！", vbExclamation, "訊息！": Exit Sub
'
'        If (RightDate - CDate(gdaRealDate.Text)) > intPara6 Then MsgBox "此日期已超過系統設定的安全期限！", vbExclamation, "訊息！"
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3400", "SO3410") Then Exit Sub
        Call subGil
        Call GiListFilter(gilClctEn, , gilCompCode.GetCodeNo)
End Sub
