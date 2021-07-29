VERSION 5.00
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO1100X 
   BorderStyle     =   1  '單線固定
   Caption         =   "BB點數管理系統 [SO1100X]"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   11790
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO1100X.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdPay 
      Caption         =   "F9.付款"
      Height          =   345
      Left            =   3870
      Picture         =   "SO1100X.frx":0442
      TabIndex        =   31
      Top             =   3150
      Width           =   780
   End
   Begin VB.CommandButton cmdCloseBill 
      Caption         =   "F8.結清"
      Height          =   345
      Left            =   3000
      Picture         =   "SO1100X.frx":05C6
      TabIndex        =   10
      Top             =   3150
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "F3.查詢"
      Height          =   345
      Left            =   7020
      Picture         =   "SO1100X.frx":074A
      TabIndex        =   11
      Top             =   3150
      Width           =   780
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "F7.轉換"
      Height          =   345
      Left            =   2115
      Picture         =   "SO1100X.frx":08CE
      TabIndex        =   9
      Top             =   3150
      Width           =   780
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "F6.贈送紅利"
      Enabled         =   0   'False
      Height          =   345
      Left            =   840
      Picture         =   "SO1100X.frx":0A52
      TabIndex        =   8
      Top             =   3150
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "結束(&X)"
      Height          =   345
      Left            =   8325
      TabIndex        =   12
      Top             =   3150
      Width           =   810
   End
   Begin VB.Frame fraFind 
      Caption         =   "查詢條件"
      Height          =   1050
      Left            =   360
      TabIndex        =   13
      Top             =   1845
      Width           =   11295
      Begin prjGiList.GiList gilSaveUpdEn 
         Height          =   315
         Left            =   1035
         TabIndex        =   5
         Top             =   570
         Width           =   2235
         _ExtentX        =   3942
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
         FldWidth2       =   1300
         F2Corresponding =   -1  'True
      End
      Begin Gi_Time.GiTime gdtSavedate1 
         Height          =   315
         Left            =   1035
         TabIndex        =   1
         Top             =   210
         Width           =   2085
         _ExtentX        =   3678
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
      Begin Gi_Time.GiTime gdtSavedate2 
         Height          =   315
         Left            =   3435
         TabIndex        =   2
         Top             =   210
         Width           =   2085
         _ExtentX        =   3678
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
      Begin Gi_Time.GiTime gdtTransdate1 
         Height          =   315
         Left            =   6465
         TabIndex        =   3
         Top             =   210
         Width           =   2085
         _ExtentX        =   3678
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
      Begin Gi_Time.GiTime gdtTransdate2 
         Height          =   315
         Left            =   8865
         TabIndex        =   4
         Top             =   210
         Width           =   2085
         _ExtentX        =   3678
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
      Begin prjGiList.GiList gilBBTransUpdEn 
         Height          =   315
         Left            =   4335
         TabIndex        =   6
         Top             =   570
         Width           =   2235
         _ExtentX        =   3942
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
         FldWidth2       =   1300
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilTransStoreCode 
         Height          =   315
         Left            =   7095
         TabIndex        =   7
         Top             =   570
         Width           =   2235
         _ExtentX        =   3942
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
         FldWidth2       =   1300
         F2Corresponding =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "商城:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6630
         TabIndex        =   30
         Top             =   630
         Width           =   435
      End
      Begin VB.Label Label18 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "轉換人員:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3450
         TabIndex        =   29
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "轉換時間:"
         Height          =   195
         Left            =   5580
         TabIndex        =   28
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   8595
         TabIndex        =   27
         Top             =   270
         Width           =   195
      End
      Begin VB.Label lblSavedate 
         AutoSize        =   -1  'True
         Caption         =   "儲值時間:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   270
         Width           =   825
      End
      Begin VB.Label lbluntil 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   3165
         TabIndex        =   24
         Top             =   270
         Width           =   195
      End
      Begin VB.Label lblSaveUpdEn 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "儲值人員:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   630
         Width           =   825
      End
   End
   Begin VB.TextBox txtTotalBonus 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   10755
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   21
      Text            =   "0"
      Top             =   3255
      Width           =   870
   End
   Begin VB.TextBox txtTotalSavePoint 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   10755
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   20
      Text            =   "0"
      Top             =   2985
      Width           =   870
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   1740
      Left            =   330
      TabIndex        =   0
      Top             =   45
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   3069
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjGiGridR.GiGridR ggrData2 
      Height          =   2235
      Left            =   330
      TabIndex        =   14
      Top             =   3555
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   3942
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjGiGridR.GiGridR ggrData3 
      Height          =   1920
      Left            =   300
      TabIndex        =   26
      Top             =   5850
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   3387
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblEditMode 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  '單線固定
      Caption         =   "顯示"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   10440
      TabIndex        =   23
      Top             =   4320
      Width           =   675
   End
   Begin VB.Label lblPresent 
      AutoSize        =   -1  'True
      Caption         =   "剩餘紅利點數:"
      Height          =   195
      Left            =   9510
      TabIndex        =   19
      Top             =   3300
      Width           =   1215
   End
   Begin VB.Label lblPrePay 
      AutoSize        =   -1  'True
      Caption         =   "剩餘儲值點數:"
      Height          =   195
      Left            =   9510
      TabIndex        =   18
      Top             =   3030
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "轉換記錄"
      Height          =   1365
      Left            =   90
      TabIndex        =   17
      Top             =   5880
      Width           =   285
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "儲值記錄"
      Height          =   1005
      Left            =   90
      TabIndex        =   16
      Top             =   3555
      Width           =   285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "設備資料"
      Height          =   1005
      Left            =   90
      TabIndex        =   15
      Top             =   90
      Width           =   285
   End
End
Attribute VB_Name = "frmSO1100X"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'981009 #5328 取SO1131E來改
'Private lngCustId As Long
Private lngEditMode As giEditModeEnu   ' //記錄目前在編輯、新增或檢視模式
Public WithEvents rsData As ADODB.Recordset
Attribute rsData.VB_VarHelpID = -1
Public WithEvents rsData2 As ADODB.Recordset
Attribute rsData2.VB_VarHelpID = -1
Public rsData3 As New ADODB.Recordset

Private objParentForm As Form
Private strWhere As String
Private strWhere2 As String

Public Property Let uParentForm(ByVal vData As Form)
    Set objParentForm = vData
End Property

' //自訂屬性：EditMode
' //記錄目前在編輯、新增或檢視模式
' //giEditModeEnu(自訂列舉值，設定於Sys_Lib)
' //取得目前編輯模式
Public Property Get EditMode() As giEditModeEnu
On Error GoTo ChkErr
   EditMode = lngEditMode
   Exit Property
ChkErr:
   Call ErrSub(Me.Name, "Get EditMode")
End Property

' //設定編輯模式
Public Property Let EditMode(ByVal vNewValue As giEditModeEnu)
On Error GoTo ChkErr
    lngEditMode = vNewValue
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Let EditMode")
End Property

' //根據傳來之編輯模式, 設定各物件屬性
Private Sub ChangeMode(ByVal lngMode As giEditModeEnu)

Dim blnFlag As Boolean  ' // 記錄是否在資料瀏覽模式
On Error GoTo ChkErr

    lngEditMode = lngMode
    Select Case lngMode
        Case giEditModeInsert
            blnFlag = False
            lblEditMode.Caption = "新增"
            cmdCancel.Caption = "取消(&X)"
        Case giEditModeEdit
            blnFlag = False
            lblEditMode.Caption = "修改"
            cmdCancel.Caption = "取消(&X)"
            'chkStopFlag.Enabled = GetUserPriv("SO1100V", "(SO1100V3)") And rsData2.RecordCount > 0
        Case giEditModeView
            blnFlag = True
            lblEditMode.Caption = "顯示"
            cmdCancel.Caption = "結束(&X)"
    End Select
     
   cmdAdd.Enabled = GetUserPriv("SO11B0", "(SO11B06)")
   cmdFind.Enabled = GetUserPriv("SO11B0", "(SO11B0A)")
   cmdChange.Enabled = GetUserPriv("SO11B0", "(SO11B07)")
   cmdCloseBill.Enabled = GetUserPriv("SO11B0", "(SO11B08)")
   cmdPay.Enabled = GetUserPriv("SO11B0", "(SO11B09)")
'   ggrData.Enabled = blnFlag
'   ggrData2.Enabled = blnFlag
'   ggrData3.Enabled = blnFlag
    If rsData.RecordCount = 0 Then
        cmdAdd.Enabled = False
        cmdChange.Enabled = False
        cmdFind.Enabled = False
        cmdCloseBill.Enabled = False
        cmdPay.Enabled = False
    End If
   
   'fraFind.Enabled = blnFlag
   Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "ChangeMode")
End Sub



Private Sub btnCloseBill_Click()

End Sub

Private Sub cmdAdd_Click()
  On Error GoTo ChkErr
    Me.MousePointer = vbHourglass
    frmSO1100X1.BBVodAccountId = rsData("FaciId") & ""
    frmSO1100X1.Show 1, Me
    If frmSO1100X1.uSave Then
      Call subChoose
      Call OpenData2
    End If
    Me.MousePointer = vbDefault
 Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdAdd_Click")
End Sub


Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
  Exit Sub
End Sub

Private Sub cmdChange_Click()
  On Error GoTo ChkErr:
    
    Dim rsSO004J As New ADODB.Recordset
    Dim strSQL  As String
    Dim rsSO193 As New Recordset
    Dim rsSO004 As New Recordset
    Me.Enabled = False
    Me.MousePointer = vbHourglass
    strSQL = "Select Nvl(A.TotalSavePoint,0) TotalSavePoint,Nvl(A.TotalBonus,0) TotalBonus,B.*,B.RowId,C.Description SavePlanName  " & _
                        " From " & GetOwner & "SO004J A," & GetOwner & "SO193 B, " & GetOwner & "CD130 C " & _
                        " Where A.bbAccountID = B.bbAccountID And A.bbAccountID = '" & rsData("FaciId") & "'" & _
                        " And B.closeBillno is Null " & _
                        " And Nvl(B.StopFlag,0) = 0 " & _
                        " And Nvl(B.bonusstopdate,To_Date('29991231235959','YYYYMMDDHH24MISS') )>= " & _
                            " To_Date('" & Format(RightNow, "yyyymmddhhmmss") & "','YYYYMMDDHH24MISS')" & _
                        " And B.SavePlanCode = C.CodeNo(+)  Order by Nvl(B.Savedate,To_Date('19110101','yyyymmdd'))"
                        
    If Not GetRS(rsSO193, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    strSQL = "Select * From " & GetOwner & "SO004 Where SeqNo = '" & rsData("SeqNo") & "'"
    If Not GetRS(rsSO004, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Set rsSO193.ActiveConnection = Nothing
    Set frmSO1100X2.uRS193 = rsSO193.Clone
    Set frmSO1100X2.uRS004 = rsSO004.Clone
    frmSO1100X2.Show 1, Me
    If frmSO1100X2.uSave Then
        Call subChoose
        Call OpenData3
        Call OpenData2
        If Not GetRS(rsSO004J, "Select * From " & GetOwner & "SO004J Where bbAccountID = '" & rsData("FaciId") & "'", _
                                gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If rsSO004J.RecordCount > 0 Then
            txtTotalSavePoint.Text = Val(rsSO004J("TotalSavePoint") & "")
            txtTotalBonus.Text = Val(rsSO004J("TotalBonus") & "")
        End If
    End If
    On Error Resume Next
    
    Call CloseRecordset(rsSO004J)
    Call CloseRecordset(rsSO193)
    Call CloseRecordset(rsSO004)
    Me.Enabled = True
    Me.MousePointer = vbDefault
    Exit Sub
ChkErr:
  Me.Enabled = True
  Me.MousePointer = vbDefault
  Call ErrSub(Me.Name, "cmdChange_Click")
End Sub

Private Sub cmdFind_Click()
  On Error GoTo ChkErr
    strWhere = ""
    strWhere2 = ""

    Call subChoose
    Call subGrd2
    Call OpenData2
    Call subGrd3
    Call OpenData3
 Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdFind_Click")
End Sub

Private Sub cmdPay_Click()
  On Error GoTo ChkErr
    Dim clsChargeInterface As Object
    Dim clsbbCashPay As Object
    Dim rsSO004J As New ADODB.Recordset
    Set clsChargeInterface = CreateObject("csbbCashPay.clsInterface")
    With clsChargeInterface
        .uOwnerName = GetOwner
        .uErrPath = ReadGICMIS1("ErrLogPath") & ""
        .ugaryGi = PutGlobal
        Set .uConn = gcnGi
        Set clsbbCashPay = .objAction("csbbCashPay.clsBBCashPay")
        If .uErr > 0 Then GoTo 88
    End With
    
    With clsbbCashPay
        .ubbAccountId = rsData("FaciId") & ""
        .uCustId = gCustId
        .uFaciSeqNo = rsData("SeqNo") & ""
        .uFaciSNo = rsData("FaciSNo") & ""
        .ShowForm
    End With
    Call subChoose
    'Call OpenData3
    Call OpenData2
    If Not GetRS(rsSO004J, "Select * From " & GetOwner & "SO004J Where bbAccountID = '" & rsData("FaciId") & "'", _
                            gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If rsSO004J.RecordCount > 0 Then
        txtTotalSavePoint.Text = Val(rsSO004J("TotalSavePoint") & "")
        txtTotalBonus.Text = Val(rsSO004J("TotalBonus") & "")
    End If

88:
  On Error Resume Next
  Call CloseRecordset(rsSO004J)
  Set clsChargeInterface = Nothing
  Set clsbbCashPay = Nothing
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPay_Click")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift)
 Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub
Private Sub FunctionKey(KeyCode As Integer, Shift As Integer)
    On Error GoTo ChkErr
    
        If KeyCode = vbKeyF6 And cmdAdd.Enabled Then Call cmdAdd_Click
        If KeyCode = vbKeyF7 And cmdChange.Enabled Then Call cmdChange_Click
        If KeyCode = vbKeyF3 And cmdFind.Enabled Then Call cmdFind_Click
        If KeyCode = vbKeyF9 And cmdPay.Enabled Then Call cmdPay_Click
        If KeyCode = vbKeyEscape Then cmdCancel_Click
        'if KeyCode=vbKeyF8 and cmdCloseBill.Enabled then cmdCloseBill

    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub
Private Sub Form_Load()
  On Error GoTo ChkErr
    Call InitData
    Call subGil
    Call subGrd
    Call OpenData
    Call DefaultValue
    Call ChangeMode(giEditModeView)
    gdtSavedate1.ShowSecond
    gdtSavedate2.ShowSecond
    gdtTransdate1.ShowSecond
    gdtTransdate2.ShowSecond
    gdtSavedate1.Text = ""
    gdtSavedate2.Text = ""
    
    'SSTab1.Tab = 0
    
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub InitData()
    On Error Resume Next
        Me.Move objParentForm.Left, objParentForm.Top + objParentForm.Height - Me.Height
End Sub

Private Sub subGil()
    On Error GoTo ChkErr
        SetgiList gilSaveUpdEn, "EmpNo", "EmpName", "CM003", , , , , 10, 20, " Where Nvl(StopFlag,0)=0", True
        SetgiList gilBBTransUpdEn, "EmpNo", "EmpName", "CM003", , , , , 10, 20, " Where Nvl(StopFlag,0)=0", True
        SetgiList gilTransStoreCode, "CodeNo", "Description", "CD129", , , 800, , , , " Where Property = 1 And Nvl(StopFlag,0) = 0 ", True
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub

Private Sub subChoose()
   On Error GoTo ChkErr
   
'   strWhere2 = " And 1 = 1 "
'   strWhere = " And 1 = 1 "
   strWhere = Empty
   strWhere = Empty
    If Len(gilTransStoreCode.GetCodeNo & "") > 0 Then
        strWhere2 = strWhere2 & " And  A.transStoreCode = '" & gilTransStoreCode.GetCodeNo & "' "
    End If
    If Len(gilBBTransUpdEn.GetCodeNo & "") > 0 Then
        strWhere2 = strWhere2 & " And A.UpdEn = '" & gilBBTransUpdEn.GetDescription & "' "
    End If
    If Len(gdtTransdate1.GetValue & "") > 0 Then
        strWhere2 = strWhere2 & " And " & _
                            " A.transdate >= To_Date('" & gdtTransdate1.GetValue & "','YYYYMMDDHH24MISS')"
    End If
    If Len(gdtTransdate2.GetValue & "") > 0 Then
        strWhere2 = strWhere2 & " And " & _
                            " A.transdate <= To_Date('" & gdtTransdate2.GetValue & "','YYYYMMDDHH24MISS')"
    End If
    If Len(gdtSavedate1.GetValue & "") > 0 Then
        strWhere = strWhere & " And  B.Savedate >= To_Date('" & gdtSavedate1.GetValue & "','YYYYMMDDHH24MISS')"
    End If
    If Len(gdtSavedate2.GetValue & "") > 0 Then
        strWhere = strWhere & " And  B.Savedate <= To_Date('" & gdtSavedate2.GetValue & "','YYYYMMDDHH24MISS')"
    End If
     If Len(gilSaveUpdEn.GetCodeNo & "") > 0 Then
        strWhere = strWhere & " And  B.UpdEn = '" & gilSaveUpdEn.GetDescription & "' "
    End If
    If Len(strWhere) > 0 Then
        strWhere = " And 1 = 1 " & strWhere
    End If
     If Len(strWhere2) > 0 Then
        strWhere2 = " And 1 = 1 " & strWhere2
    End If
   Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub OpenData()
    On Error GoTo ChkErr
    Dim strSQL As String
        Set rsData = New ADODB.Recordset
        
        strSQL = "Select A.*, B.Description InitPlace,Nvl(C.RefNo,0) RefNo " & _
                "From " & GetOwner & "SO004 A," & GetOwner & "CD056 B," & GetOwner & "CD022 C " & _
                " Where A.InitPlaceNo=B.CodeNo(+) And A.FaciCode=C.CodeNo  " & _
                " And A.CustId=" & gCustId & " And  A.CompCode=" & gCompCode & _
                " And Nvl(C.RefNo,0)=3 Order By Nvl(A.InstDate,To_Date('19110101','yyyymmdd'))"
        
        If Not GetRS(rsData, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        
        Set rsData.ActiveConnection = Nothing
        Set ggrData.Recordset = rsData
        ggrData.Refresh
        If rsData.RecordCount = 0 Then
            cmdAdd.Enabled = False
            cmdChange.Enabled = False
            cmdFind.Enabled = False
            cmdCloseBill.Enabled = False
        End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub

Private Sub OpenData2()
    On Error GoTo ChkErr
    Dim strSQL As String
        Set rsData2 = New ADODB.Recordset
        Dim rsSO004J As New ADODB.Recordset
        If rsData.RecordCount > 0 Then
            strSQL = "Select A.TotalSavePoint,A.TotalBonus,B.*,B.RowId,C.Description SavePlanName  " & _
                        " From " & GetOwner & "SO004J A," & GetOwner & "SO193 B, " & GetOwner & "CD130 C " & _
                        " Where A.bbAccountID = B.bbAccountID And A.bbAccountID = '" & rsData("FaciId") & "'" & _
                        " And B.SavePlanCode = C.CodeNo(+) " & strWhere & " Order by Nvl(B.Savedate,To_Date('19110101','yyyymmdd'))"
        Else
            strSQL = "Select A.TotalSavePoint,A.TotalBonus,B.* ,C.Description SavePlanName " & _
                        " From " & GetOwner & "SO004J A ," & GetOwner & "SO193 B, " & GetOwner & "CD130 C " & _
                        " Where 1 = 0 "
        End If
        
        If Not GetRS(rsData2, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If rsData.RecordCount > 0 And rsData2.RecordCount = 0 And Len(strWhere) > 0 Then MsgBox "查無點數記錄!!", vbInformation, "訊息"
        Set rsData2.ActiveConnection = Nothing
        Set ggrData2.Recordset = rsData2
        ggrData2.Refresh
        If rsData.RecordCount > 0 Then
            If Not GetRS(rsSO004J, "Select * From " & GetOwner & "SO004J Where bbAccountID = '" & rsData("FaciId") & "'", _
                                gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
            If rsSO004J.RecordCount > 0 Then
                txtTotalSavePoint.Text = Val(rsSO004J("TotalSavePoint") & "")
                txtTotalBonus.Text = Val(rsSO004J("TotalBonus") & "")
            End If
        End If
        On Error Resume Next
        Call CloseRecordset(rsSO004J)
        strWhere = Empty
    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData2"
End Sub

Private Sub OpenData3()
    On Error GoTo ChkErr
    Dim strSQL As String
        Set rsData3 = New ADODB.Recordset
            If rsData2.RecordCount > 0 Then
                
                strSQL = "Select A.*,B.Description transStoreName From " & GetOwner & "SO033BBTRANS A," & GetOwner & "CD129 B " & _
                                " Where A.transStoreCode = B.CodeNo " & _
                                " And A.bbAccountID = '" & rsData2("bbAccountID") & "' " & strWhere2 & " Order By SeqNo,transdate"
                                
            Else
                strSQL = "Select A.*,B.Description transStoreName From " & GetOwner & "SO033BBTRANS A," & GetOwner & "CD129 B " & _
                            " Where 1 = 0 "
            End If
         
        
        If Not GetRS(rsData3, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        
        Set rsData3.ActiveConnection = Nothing
        Set ggrData3.Recordset = rsData3
        ggrData3.Refresh
        strWhere2 = Empty
    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData3"
End Sub

Private Sub DefaultValue()
    On Error Resume Next
        If rsData.RecordCount = 0 Then Exit Sub
        rsData.MoveFirst

End Sub

Private Sub subGrd()
    On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
        mFlds.Add "PRDate", , , , , "機上盒狀態   "
        mFlds.Add "FaciSNo", , , , , "機上盒序號               "
        mFlds.Add "FaciCode", , , , , "代碼", vbRightJustify
        mFlds.Add "FaciName", , , , , "項目名稱     "
        mFlds.Add "InstDate", giControlTypeDate, , , , "裝機日期"
        mFlds.Add "PRDate", giControlTypeDate, , , , "拆機日期"
        mFlds.Add "InitPlace", , , , , "裝置點   "
        mFlds.Add "ModelName", , , , , "型號名稱"
        mFlds.Add "BuyName", , , , , "買賣方式"
        mFlds.Add "SmartCardNo", , , , , "智慧卡序號        "
        mFlds.Add "SeqNo", , , , , "設備唯一值          "
        ggrData.AllFields = mFlds
        ggrData.UseCellForeColor = True
        ggrData.SetHead
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subGrd"
End Sub

Private Sub subGrd2()
    On Error GoTo ChkErr
    Dim mFlds2 As New GiGridFlds
        mFlds2.Add "StopFlag", , , , , "停用", vbCenter
        mFlds2.Add "SeqNo", , , , , "儲值流水號             ", vbRightJustify
        mFlds2.Add "Savedate", giControlTypeTime, , , , "儲值日期            "
        mFlds2.Add "Savepoint", , , , , "儲值點數"
        mFlds2.Add "SavePlanName", , , , , "儲值方案            "
        mFlds2.Add "bonus", , , , , "紅利點數"
        mFlds2.Add "SaveBillno", , , , , "收費單號          "
        mFlds2.Add "CitemName", , , , , "收費項目                "
        mFlds2.Add "bonusstopdate", giControlTypeDate, , , , "紅利到期日       "
        mFlds2.Add "UsedSavepoint", , , , , "儲值扣抵點數"
        mFlds2.Add "Usedbonus", , , , , "紅利扣抵點數"
        mFlds2.Add "UpdEn", , , , , "異動人員"
        ggrData2.AllFields = mFlds2
        ggrData2.UseCellForeColor = True
        ggrData2.SetHead
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subGrd2"
End Sub

Private Sub subGrd3()
    On Error GoTo ChkErr
    Dim mFlds3 As New GiGridFlds
        mFlds3.Add "transdate", giControlTypeTime, , , , "轉換日期            "
        mFlds3.Add "transSavepoint", , , , , "轉換儲值點數"
        mFlds3.Add "transbonus", , , , , "轉換紅利點數"
        mFlds3.Add "transStoreName", , , , , "轉換對象       "
        mFlds3.Add "Rate", , , , , "匯率  "
        mFlds3.Add "Mallpoint", , , , , "商城點數"
        mFlds3.Add "Mallbonus", , , , , "商城贈點"
        mFlds3.Add "UpdEn", , , , , "異動人員  "
        ggrData3.AllFields = mFlds3
        ggrData3.UseCellForeColor = True
        ggrData3.SetHead
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subGrd3"
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        Call CloseRecordset(rsData)
        Call CloseRecordset(rsData2)
        Call CloseRecordset(rsData3)
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        ReleaseCOM frmSO1100X
End Sub
Private Sub gdtSavedate1_GotFocus()
  On Error Resume Next
    If gdtSavedate1.GetValue <> "" Then Exit Sub
    gdtSavedate1.SetValue Date
    gdtSavedate1.ShowSecond
End Sub



Private Sub gdtSavedate1_Validate(Cancel As Boolean)
On Error GoTo ChkErr
    If gdtSavedate1.GetValue = "" Then
        If gdtSavedate2.GetValue <> "" Then gdtSavedate2.Text = ""
    End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdtSavedate1_Validate")
End Sub

Private Sub gdtSavedate2_GotFocus()
    On Error Resume Next
    If gdtSavedate2.GetValue <> "" Then Exit Sub
    gdtSavedate2.SetValue GetDayLastTime(gdtSavedate1.GetValue(True))
    gdtSavedate2.ShowSecond
End Sub

Private Sub gdtSavedate2_Validate(Cancel As Boolean)
On Error Resume Next
     If gdtSavedate1.GetValue = "" Or gdtSavedate2.GetValue = "" Then Exit Sub
     If Not IsDate(gdtSavedate2.Text) Then Exit Sub
     If DateDiff("d", gdtSavedate1.GetValue(True), gdtSavedate2.GetValue(True)) < 0 Then
        MsgBox "儲值時間截止日不可小於儲值時間起始日!", vbExclamation, "警告"
        gdtSavedate2.SetValue GetDayLastTime(gdtSavedate1.GetValue(True))
        Cancel = True
     End If
End Sub


Private Sub gdtTransdate1_GotFocus()
   On Error Resume Next
    If gdtTransdate1.GetValue <> "" Then Exit Sub
    gdtTransdate1.SetValue Date
    gdtTransdate1.ShowSecond
End Sub

Private Sub gdtTransdate1_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If gdtTransdate1.GetValue = "" Then
        If gdtTransdate2.GetValue <> "" Then gdtTransdate2.Text = ""
    End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gdtTransdate1_Validate")
End Sub


Private Sub gdtTransdate2_GotFocus()
  On Error Resume Next
    If gdtTransdate2.GetValue <> "" Then Exit Sub
    gdtTransdate2.SetValue GetDayLastTime(gdtTransdate1.GetValue(True))
    gdtTransdate2.ShowSecond
End Sub

Private Sub gdtTransdate2_Validate(Cancel As Boolean)
On Error Resume Next
     If gdtTransdate1.GetValue = "" Or gdtTransdate2.GetValue = "" Then Exit Sub
     If Not IsDate(gdtTransdate2.Text) Then Exit Sub
     If DateDiff("d", gdtTransdate1.GetValue(True), gdtTransdate2.GetValue(True)) < 0 Then
        MsgBox "轉換時間截止日不可小於轉換時間起始日!", vbExclamation, "警告"
        gdtTransdate2.SetValue GetDayLastTime(gdtTransdate1.GetValue(True))
        Cancel = True
     End If

End Sub

'Private Sub ggrData_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
'        If giFld.FieldName = "SelectStatus" Then
'            If Value <> "" Then Value = vbRed
'        End If
'End Sub

'Private Sub ggrData_DblClick()
'    On Error Resume Next
'        If lngEditMode = giEditModeView Then Exit Sub
'        If Not IsDataOk Then Exit Sub
'        strFaciSNo = rsData("FaciSNo") & ""
'        strFaciSeqNo = rsData("SeqNo") & ""
'        strSmartCardNo = rsData("SmartCardNo") & ""
'        Unload Me
'End Sub

'Private Function IsDataOk() As Boolean
'    On Error GoTo ChkErr
'        If rsData.RecordCount > 0 Then
'            If rsData("PRDate") & "" <> "" Then
'                If blnPRCanClick Then
'                    If MsgBox("該設備已拆除,請確定是否選取??", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then Exit Function
'                Else
'                    MsgBox "拆除的設備不能選擇!!", vbExclamation, gimsgPrompt: Exit Function
'                End If
'            End If
'            If rsData("FaciCode") & "" <> "" Then
'                If GetRsValue("Select Nvl(RefNo,0) From " & GetOwner & "CD022 Where CodeNo = " & rsData("FaciCode")) = 4 Then
'                    MsgBox "設備參考號4不能選擇!!", vbExclamation, gimsgPrompt: Exit Function
'                End If
'            End If
'        End If
'        IsDataOk = True
'    Exit Function
'ChkErr:
'    ErrSub Me.Name, "IsDataOk"
'End Function

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    If giFld.FieldName = "PRDate" And Trim(giFld.HeadName) = "機上盒狀態" Then
            Value = GetFaciStatus(ggrData.Recordset, Value)

    End If
'    If giFld.FieldName = "VODStatus" Then
'        Select Case Value
'            Case 1
'                Value = "啟用"
'            Case 0
'                Value = "未啟用"
'            Case 2
'                Value = "暫停"
'            Case Else
'                Value = ""
'        End Select
'    End If
End Sub

'981118 #5327 測試後調整,將停用欄位移至叩抵點數旁,若為停用則顯示紅色
Private Sub ggrData2_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
        If giFld.FieldName = "StopFlag" Then
            If Value = 1 Or Value = "是" Then Value = vbRed
        End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "ggrData2_ColorCellData"
End Sub

Private Sub ggrData2_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
'    If giFld.FieldName = "PRDate" And Trim(giFld.HeadName) = "設備狀態" Then
'            Value = GetFaciStatus(ggrData2.Recordset, Value)
'    End If
    'If giFld.FieldName = "CloseFlag" Then Value = IIf(Value = 0, "否", "是")
    If giFld.FieldName = "StopFlag" Then
        Value = IIf(Value = 0, "否", "是")
    End If
'    '981110 #5328「異動原因」後,增加BILLNO(單號編號)、ITEM(項次)、MINUSTYPE扣點種類(1 購買、2 贈送、3 預借)
'    If giFld.FieldName = "MINUSTYPE" Then
'        '2013.07.22 #6547 調整顯示 取消#5328的限制
'        'If ggrData2.Recordset.Fields("CreditSeqNo") & "" <> "" Then Value = 0
'
'        Select Case Val(Value)
'            Case 1
'                Value = "購買"
'            Case 2
'                Value = "贈送"
'            Case 3
'                Value = "預借"
'            Case Else
'                Value = ""
'        End Select
'    End If
'
'    '981113 #5328 penny提只取後8碼
'    If giFld.FieldName = "CreditSeqNo" Then
'        If Value <> "" Then Value = Right(Value, 8)
'    End If
End Sub

Private Sub ggrData3_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
'  On Error Resume Next
'    If giFld.FieldName = "CloseFlag" Then Value = IIf(Value = 0, "否", "是")
'    If giFld.FieldName = "StopFlag" Then Value = IIf(Value = 0, "否", "是")
End Sub





Private Sub rsData_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  On Error GoTo ChkErr
    strWhere = ""
    Call subGrd2
    Call OpenData2
    Call subGrd3
    Call OpenData3
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "rsData_MoveComplete")
End Sub



'Private Sub gdtRecTime1_GotFocus()
'  On Error Resume Next
'    If gdtRecTime1.GetValue <> "" Then Exit Sub
'    gdtRecTime1.SetValue Date
'    gdtRecTime1.ShowSecond
'End Sub
'
'Private Sub gdtRecTime1_Validate(Cancel As Boolean)
'On Error GoTo ChkErr
'    If gdtRecTime1.GetValue = "" Then
'        If gdtRecTime2.GetValue <> "" Then gdtRecTime2.Text = ""
'    End If
'Exit Sub
'ChkErr:
'    Call ErrSub(Me.Name, "gdtRecTime1_Validate")
'End Sub
'
'Private Sub gdtRecTime2_GotFocus()
'  On Error Resume Next
'    If gdtRecTime2.GetValue <> "" Then Exit Sub
'    gdtRecTime2.SetValue GetDayLastTime(gdtRecTime1.GetValue(True))
'    gdtRecTime2.ShowSecond
'End Sub
'
'Private Sub gdtRecTime2_Validate(Cancel As Boolean)
'   On Error Resume Next
'     If gdtRecTime1.GetValue = "" Or gdtRecTime2.GetValue = "" Then Exit Sub
'     If Not IsDate(gdtRecTime2.Text) Then Exit Sub
'     If DateDiff("d", gdtRecTime1.GetValue(True), gdtRecTime2.GetValue(True)) < 0 Then
'        MsgBox "紀錄時間截止日不可小於紀錄時間起始日!", vbExclamation, "警告"
'        gdtRecTime2.SetValue GetDayLastTime(gdtRecTime1.GetValue(True))
'        Cancel = True
'     End If
'End Sub


