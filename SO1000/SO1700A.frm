VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO1700A 
   BorderStyle     =   1  '單線固定
   Caption         =   "服務申告管理  [SO1700A]"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO1700A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin prjGiGridR.GiGridR ggrdata 
      Height          =   4575
      Left            =   90
      TabIndex        =   12
      Top             =   2460
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   8070
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
   Begin VB.CommandButton cmdView 
      Caption         =   "F12. 單筆..."
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   7185
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "F11.修改..."
      Height          =   375
      Left            =   2265
      TabIndex        =   9
      Top             =   7185
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "F4. 清除"
      Height          =   375
      Left            =   1185
      TabIndex        =   8
      Top             =   7185
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   10665
      TabIndex        =   19
      Top             =   7185
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "F3. 尋找"
      Height          =   375
      Left            =   105
      TabIndex        =   7
      Top             =   7185
      Width           =   1095
   End
   Begin VB.Frame fraData 
      Height          =   2115
      Left            =   90
      TabIndex        =   18
      Top             =   60
      Width           =   11655
      Begin VB.OptionButton optContentCode 
         Caption         =   "&6. 申告內容細項"
         Height          =   345
         Left            =   5850
         TabIndex        =   17
         Top             =   1125
         Width           =   1695
      End
      Begin VB.OptionButton optServDescCode 
         Caption         =   "&5. 申告內容"
         Height          =   345
         Left            =   5850
         TabIndex        =   16
         Top             =   735
         Width           =   1335
      End
      Begin MSMask.MaskEdBox mskCustID 
         Height          =   375
         Left            =   2190
         TabIndex        =   1
         Top             =   720
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         PromptChar      =   "_"
      End
      Begin Gi_Date.GiDate gdaAcceptTime 
         Height          =   405
         Left            =   2190
         TabIndex        =   0
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   714
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
      Begin prjGiList.GiList gilServiceCode 
         Height          =   315
         Left            =   7680
         TabIndex        =   3
         Top             =   330
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
         F3Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin VB.OptionButton optAll 
         Caption         =   "&A. 所有"
         Height          =   345
         Left            =   450
         TabIndex        =   6
         Top             =   1530
         Width           =   945
      End
      Begin VB.OptionButton optServiceCode 
         Caption         =   "&4. 來電分類"
         Height          =   345
         Left            =   5850
         TabIndex        =   15
         Top             =   300
         Width           =   1635
      End
      Begin VB.OptionButton optTel1 
         Caption         =   "&3. 電話(1)"
         Height          =   345
         Left            =   450
         TabIndex        =   14
         Top             =   1125
         Width           =   1305
      End
      Begin VB.OptionButton optCustId 
         Caption         =   "&2. 客戶編號"
         Height          =   345
         Left            =   450
         TabIndex        =   13
         Top             =   735
         Width           =   1455
      End
      Begin VB.OptionButton optAcceptTime 
         Caption         =   "&1. 受理日期"
         Height          =   345
         Left            =   450
         TabIndex        =   11
         Top             =   300
         Value           =   -1  'True
         Width           =   1515
      End
      Begin VB.TextBox txtTel1 
         Height          =   375
         Left            =   2190
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1125
         Width           =   2325
      End
      Begin prjGiList.GiList gilServDescCode 
         Height          =   315
         Left            =   7680
         TabIndex        =   4
         Top             =   735
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
         F3Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin prjGiList.GiList gilContentCode 
         Height          =   315
         Left            =   7680
         TabIndex        =   5
         Top             =   1110
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
         F3Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
   End
End
Attribute VB_Name = "frmSO1700A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
' 記錄目前在編輯、新增或檢視模式
Private lngEditMode As giEditModeEnu
Private strParam0 As String
Private rsSo1700 As ADODB.Recordset
Attribute rsSo1700.VB_VarHelpID = -1
Private strPermission As String
'是否使用數變
Private Const blnYesNo = False
'接收上層數變
Private strStatus As String
'權限依據數變
Private blnUserPrivUpdate   As Boolean
Private blnUserPrivView   As Boolean

Public Sub FunctionKey(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Select Case KeyCode
           '----------------------------------------------------
           Case vbKeyF3  '   F3: 尋找, 相當於按下cmdfind
                        If Not ChkGiList(KeyCode) Then Exit Sub
                        If cmdFind.Enabled = True Then cmdFind.Value = True: Exit Sub
           Case vbKeyF4  '   F4: 清除, 相當於按下cmdClear
                        If cmdClear.Enabled = True Then cmdClear.Value = True: Exit Sub
           Case vbKeyF11  '   F11: 修改, 相當於按下cmdEdit
                        If cmdEdit.Enabled = True Then cmdEdit.Value = True: Exit Sub
           Case vbKeyF12  '   F12: 單筆, 相當於按下cmdView
                        If cmdView.Enabled = True Then cmdView.Value = True: Exit Sub
           '----------------------------------------------------
           Case vbKeyEscape
                        cmdCancel.Value = True: Exit Sub
           '----------------------------------------------------
           Case vbKeyUp
                        KeyCode = 0
                        SendKeys "+{Tab}"
           Case vbKeyDown
                        KeyCode = 0
                        SendKeys "{Tab}"
           '----------------------------------------------------
    End Select
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub

Private Function IsDataOK()
   On Error GoTo ChkErr
   ' 檢查必要欄位是否有值, 若為必要欄位且無值, 則顯示"欄位必需有值",
   ' 且focus移到該欄位上
   IsDataOK = False
   Select Case True
    Case optAcceptTime.Value
        If gdaAcceptTime.GetValue = "" Then
           gdaAcceptTime.SetFocus
           GoTo 66
        Else
           IsDataOK = True
        End If
    Case optCustId.Value
        If mskCustID.Text = "" Then
           mskCustID.SetFocus
           GoTo 66
        Else
           IsDataOK = True
        End If
    Case optTel1.Value
        If txtTel1.Text = "" Then
           txtTel1.SetFocus
           GoTo 66
        Else
           IsDataOK = True
        End If
    Case optServiceCode.Value
        If gilServiceCode.GetCodeNo = "" Then
            gilServiceCode.SetFocus
            GoTo 66
        Else
            IsDataOK = True
        End If
    Case optServDescCode.Value
        If gilServDescCode.GetCodeNo = "" Then
            gilServDescCode.SetFocus
            GoTo 66
        Else
            IsDataOK = True
        End If
    Case optContentCode.Value
        If gilContentCode.GetCodeNo = "" Then
            gilContentCode.SetFocus
            GoTo 66
        Else
            IsDataOK = True
        End If
    Case optAll
        IsDataOK = True
   End Select
   
  Exit Function
66:
    MsgBox " 此為必要欄位,須有值 !! ", vbExclamation, "訊息.."
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub SubGiList()
On Error GoTo ChkErr
    Dim mFlds As New prjGiGridR.GiGridFlds
   '設定各個GiList的屬性
    SetLst gilServiceCode, "CodeNo", "Description", 3, 12, , , "CD008"
    SetLst gilServDescCode, "CodeNo", "Description", 3, 12, , , "CD008A"
    
    '#3495 增加申告內容細項 By Kin 2008/02/19
    SetLst gilContentCode, "CodeNo", "Description", 3, 12, , , "CD008D"
    gilServiceCode.SetParent Me
   
    With mFlds
        .Add "AcceptTime", giControlTypeDate, , , False, "受理時間     ", vbLeftJustify
        .Add "CustID", , , , False, "客戶編號    ", vbLeftJustify
        .Add "CustName", , , , False, "客戶姓名 ", vbLeftJustify
        .Add "ServiceName", , , , False, "       申告原因        ", vbLeftJustify
        .Add "MediaName", , , , , "介紹媒介  "
        .Add "PromName", , , , , "業務活動  "
'        .Add "PromName", , , , , "促銷活動  "
        .Add "ServDescName", , , , False, "      申告內容        ", vbLeftJustify
        
        '#3495 增加申告內容細項 By Kin 2008/02/19
        .Add "ContentName", , , , False, "申告內容細項      ", vbLeftJustify
        .Add "Tel1", , , , False, "電話1              ", vbLeftJustify
        .Add "AcceptName", , , , False, "受理人員", vbLeftJustify
        .Add "HANDLETIME", giControlTypeTime, , , False, "處理時間" & Space(15), vbLeftJustify
        .Add "HandleName", , , , False, "處理人員", vbLeftJustify
        .Add "Updtime", giControlTypeTime, , , False, "異動時間" & Space(15), vbLeftJustify
        .Add "UpdEn", , , , False, "異動人員", vbLeftJustify
    End With
   
    With ggrdata
        .AllFields = mFlds
        .SetHead
         Set .Recordset = rsSo1700
        .Refresh
    End With
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "SubGiLst")
End Sub

Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error Resume Next
   If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
  '取global變數內容,取資料內容
   On Error GoTo ChkErr
   Dim mFlds As New GiGridFlds
   Dim strSQL As String
    'MsgBox GetOwner
   '呼叫系統參數設定
'   If Alfa2 = True Then
'      Call GetGlobal
'   End If
   
    If Crm Then
        Me.Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2)
    Else
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End If
    
   '權限設定
    Call SetCommandPermission
       
   'Select string
   
   'AcceptTime，HandleTime需轉換成中曆時間字串。
   '#3495 增加ContentCode,ContentName欄位 By Kin 2008/02/19
   strParam0 = "select ServiceCode,MediaCode,MediaName,PromCode,PromName,SERVICETYPE,AcceptEn,AcceptTime,CustId,CustName," & _
               "ServiceName,HANDLETIME, AcceptName, HandleName, UpdTime, UpdEn," & _
               "Note,HandleEn,HandleNote,Tel1,ServDescCode,ServDescName,ProcResultNo,ProcResultName, " & _
               "ContentCode,ContentName "
   strSQL = strParam0 & " FROM " & GetOwner & "SO006 WHERE Custid=-1 Order by AcceptTime,HANDLETIME Desc"
   
   Set rsSo1700 = New ADODB.Recordset
   
   If rsSo1700.State = 1 Then
        rsSo1700.Close
   Else
        rsSo1700.CursorLocation = adUseClient
   End If
   rsSo1700.Open strSQL, gcnGi, adOpenKeyset, adLockOptimistic
'   If OpenDB.OpenRecordset(rsSo1700, strSQL, gcnGi, adOpenKeyset, adLockOptimistic, adUseClient, False, False) = giOK Then
   If rsSo1700.EOF Then
      cmdEdit.Enabled = False
      cmdView.Enabled = False
   End If
   Call SubGiList       '設定gilist 屬性
'    Else
'            MsgBox "系統初始化失敗！請重新操作！", vbCritical, "訊息！"
'    End If
    
   Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '若顯示模式:
    '若新增或修改模式:
    '   F1: Help, 若游標恰好在某gilst元件上, 則相當於點選該ComboBox
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift)
  Exit Sub
ChkErr:
       Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub cmdFind_Click()
  On Error GoTo ChkErr
   '先檢查必要欄位, 若不過則不可繼續
    Dim strParam As String
    Dim strSQL As String
   
   ' 檢查資料是否可儲存
    If IsDataOK() = False Then Exit Sub
    Select Case True
        Case optAcceptTime
            strParam = " WHERE  to_char(AcceptTime, 'YYYYMMDD') ='" & gdaAcceptTime.GetValue & "' order by AcceptTime DESC"
        Case optCustId.Value
            strParam = " WHERE CustId=" & mskCustID.Text & " order by AcceptTime DESC"
        Case optTel1.Value
            strParam = " WHERE  Tel1='" & txtTel1.Text & "' order by AcceptTime DESC"
        Case optServiceCode.Value
            strParam = " WHERE ServiceCode='" & gilServiceCode.GetCodeNo & "' order by AcceptTime DESC"
        Case optServDescCode.Value
            strParam = " WHERE ServDescCode='" & gilServDescCode.GetCodeNo & "' order by AcceptTime DESC"
        '#3495 增加申告內容查詢條件 By Kin 2008/02/19
        Case optContentCode.Value
            strParam = " Where ContentCode=" & gilContentCode.GetCodeNo
        Case optAll.Value
            strParam = " Order by AcceptTime Desc ,HandleTime"
    End Select
    'all fields: address,ServName,AreaName,ClctAreaName,ZipCode,CircuitNo,BTName,ClctName,StrtCode,Ord,UpdTime,UpdEn
    strSQL = strParam0 & " FROM " & GetOwner & "So006  " & strParam
    If rsSo1700.State = adStateOpen Then rsSo1700.Close
    rsSo1700.CursorLocation = adUseClient
    rsSo1700.Open strSQL, gcnGi, adOpenKeyset, adLockOptimistic
'   If OpenDB.OpenRecordset(rsSo1700, strSQL, gcnGi, adOpenKeyset, adLockOptimistic, adUseClient, , False) <> giOK Then MsgBox "資料集合初始化失敗！", vbExclamation, "訊息！"
    
    Set ggrdata.Recordset = rsSo1700
    
    If rsSo1700.EOF Or rsSo1700.BOF Then
        MsgBox "無資料！", vbExclamation, "訊息！"
        ggrdata.Blank = True
    Else
        ggrdata.Blank = False
        cmdEdit.Enabled = GetUserPriv("SO1700", "(SO17002)")     'edit
        cmdView.Enabled = GetUserPriv("SO1700", "(SO17004)")      'view
    End If
   
    ggrdata.Refresh
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdFind")
End Sub

Private Sub NewRcd()
'清除螢幕上的欄位
  On Error Resume Next
    gdaAcceptTime.Text = ""
    mskCustID.Text = ""
    txtTel1.Text = ""
    gilServiceCode.Clear
    gilServDescCode.Clear
End Sub

Private Sub cmdClear_Click()
'將各選項之應textbox 清回初值(空白或預設值),並將gird 內容清空
  On Error GoTo ChkErr
    Call NewRcd
    If rsSo1700.State = adStateOpen Then rsSo1700.Close
    rsSo1700.Open "Select * FROM " & GetOwner & "So016 WHERE strtcode=-1 Order by StrtCode,Ord", gcnGi, adOpenForwardOnly, adLockReadOnly
    Set ggrdata.Recordset = rsSo1700
    ggrdata.Refresh
    cmdEdit.Enabled = False
    cmdView.Enabled = False
  Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdClear_Click"
End Sub

Private Sub cmdEdit_Click()
'檢查grid中是否有資料,若無則顯示'無資料',
'若有則呼叫so7400B,進入修改模式
      On Error GoTo ChkErr
      If rsSo1700.EOF Then MsgBox "無資料！", vbExclamation, "訊息！": Exit Sub
      gServiceType = rsSo1700!ServiceType & ""
      Screen.MousePointer = vbHourglass
      DoEvents
     ' EditMode = giEditModeEdit
      frmSO1700B.EditMode = giEditModeEdit
      frmSO1700B.GetRecordSet = rsSo1700
      frmSO1700B.Show vbModal
      Screen.MousePointer = vbDefault
      Call cmdFind_Click
      DoEvents
      Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdEdit_Click")
End Sub

Private Sub cmdView_Click()
'檢查grid中是否有資料,若無則顯示'無資料',
'若有則呼叫so7400B,進入顯示模式
    On Error GoTo ChkErr
      If rsSo1700.EOF Or rsSo1700.BOF Then MsgBox "無資料！", vbExclamation, "訊息！": Exit Sub
      gServiceType = rsSo1700!ServiceType & ""
      Screen.MousePointer = vbHourglass
      DoEvents
      frmSO1700B.EditMode = giEditModeView
      frmSO1700B.GetRecordSet = rsSo1700
      frmSO1700B.Show vbModal
      Screen.MousePointer = vbDefault
      DoEvents
      Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdView_Click")
End Sub

Private Sub SetCommandPermission()
  On Error GoTo ChkErr
'Get command permission
    Call GetPermission(IIf(blnYesNo, strStatus, garyGi(4)), "SO1700", strPermission)
    '取得各個按鈕的權限！
    Call SetPermission(strPermission, _
                        blnUpdate:=blnUserPrivUpdate, UpdKey:="SO17002", _
                        blnView:=blnUserPrivView, ViewKey:="SO17003")
    cmdEdit.Enabled = blnUserPrivUpdate
    cmdView.Enabled = blnUserPrivView
  Exit Sub
ChkErr:
    ErrSub Me.Name, "SetCommandPermission"
End Sub

Private Sub gdaAcceptTime_GotFocus()
  On Error Resume Next
    optAcceptTime.Value = True
End Sub

Private Sub gilContentCode_Change()
  On Error Resume Next
  optContentCode.Value = True
End Sub

'Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
'AcceptTime，HandleTime需轉換成中曆時間字串。
'On Error GoTo ChkErr
'    If fld.Name = "ACCEPTTIME" Or fld.Name = "HANDLETIME" Then
 '       Value = Format(Value, "ee/mm/dd HH:MM AMPM")
 '   End If
    
'Exit Sub
'ChkErr:
 '   Call ErrSub(Me.Name, "ggrdata_ShowCellDate")
'End Sub

Private Sub gilServiceCode_GotFocus()
  On Error Resume Next
    optServiceCode.Value = True
End Sub

Private Sub gilServDescCode_GotFocus()
  On Error Resume Next
    optServDescCode.Value = True
End Sub

Private Sub mskCustId_GotFocus()
  On Error Resume Next
    optCustId.Value = True
    mskCustID.SelStart = 0
    mskCustID.SelLength = Len(mskCustID.Text)
End Sub

Private Sub mskCustID_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    If Not (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8) Then
        KeyAscii = 9
    End If
End Sub

Private Sub optAcceptTime_GotFocus()
  On Error Resume Next
    gdaAcceptTime.SetFocus
End Sub

Private Sub optAll_GotFocus()
  On Error Resume Next
    optAll.SetFocus
End Sub

Private Sub optCustId_GotFocus()
  On Error Resume Next
    mskCustID.SetFocus
End Sub

Private Sub optServiceCode_GotFocus()
  On Error Resume Next
    gilServiceCode.SetFocus
End Sub

Private Sub optServDescCode_GotFocus()
  On Error Resume Next
    gilServDescCode.SetFocus
End Sub

Private Sub optTel1_GotFocus()
  On Error Resume Next
    txtTel1.SetFocus
End Sub

Private Sub txtTel1_GotFocus()
  On Error Resume Next
    txtTel1.SelStart = 0
    txtTel1.SelLength = Len(txtTel1.Text)
    optTel1.Value = True
End Sub

Public Property Let uCustStatus(ByVal vNewValue As Variant)
  On Error GoTo ChkErr
    strStatus = vNewValue
  Exit Property
ChkErr:
    ErrSub Me.Name, "Let uCustStatus"
End Property

