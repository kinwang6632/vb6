VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSO3000 
   BackColor       =   &H00C0FFC0&
   Caption         =   "計費管理 [SO3000]"
   ClientHeight    =   8310
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11880
   Icon            =   "SO3000.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   300
      TabIndex        =   1
      Top             =   240
      Width           =   1245
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '對齊表單下方
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   7965
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "計費管理 [SO3000]"
            TextSave        =   "計費管理 [SO3000]"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgPic 
      Height          =   4665
      Left            =   2010
      Top             =   930
      Width           =   6270
   End
   Begin VB.Menu mnu3100 
      Caption         =   "&1.出單管理"
      Begin VB.Menu mnu3110 
         Caption         =   "&1.本期收費資料產生..."
      End
      Begin VB.Menu mnu3120 
         Caption         =   "&2.本期應收資料查詢列印"
         Begin VB.Menu mnu3121 
            Caption         =   "&1.本期應收明細..."
         End
         Begin VB.Menu mnu3122 
            Caption         =   "&2.本期應收彙總..."
         End
         Begin VB.Menu mnu3123 
            Caption         =   "&3.本期應收區域及道路分析..."
         End
         Begin VB.Menu mnu3124 
            Caption         =   "&4.本期應收期數及金額分析..."
         End
         Begin VB.Menu mnu3125 
            Caption         =   "&5.本期出單數及金額查詢..."
         End
      End
      Begin VB.Menu mnu3130 
         Caption         =   "&3.換單調整..."
      End
      Begin VB.Menu mnu3140 
         Caption         =   "&4.減免出單金額(北視用)"
      End
      Begin VB.Menu mnu3150 
         Caption         =   "&5.合併帳單"
      End
   End
   Begin VB.Menu mnu3200 
      Caption         =   "&2.收費單管理"
      Begin VB.Menu mnu3210 
         Caption         =   "&1.收費單過入..."
      End
      Begin VB.Menu mnu3220 
         Caption         =   "&2.印單序號給號..."
      End
      Begin VB.Menu mnu3230 
         Caption         =   "&3 收費單報表查詢..."
      End
      Begin VB.Menu mnu3240 
         Caption         =   "&4.收費資料查詢／編輯..."
      End
      Begin VB.Menu mnu3250 
         Caption         =   "&5.整批調整"
         Begin VB.Menu mnu3251 
            Caption         =   "&1.印單金額調整..."
         End
         Begin VB.Menu mnu3252 
            Caption         =   "&2.收費單整批作廢..."
         End
         Begin VB.Menu mnu3253 
            Caption         =   "&3.收費單換單調整..."
         End
         Begin VB.Menu mnu3254 
            Caption         =   "&4.收費單收費方式調整..."
         End
      End
      Begin VB.Menu mnu3260 
         Caption         =   "&6.單據列印"
         Begin VB.Menu mnu3261 
            Caption         =   "&1.整批收費單據列印..."
         End
         Begin VB.Menu mnu3262 
            Caption         =   "&2.單筆收費單據列印..."
         End
         Begin VB.Menu mnu3263 
            Caption         =   "&3.收費單據重新列印..."
         End
      End
      Begin VB.Menu mnu3270 
         Caption         =   "&7.轉扣帳磁片資料產生"
         Begin VB.Menu mnu3271 
            Caption         =   "&1.媒體轉帳資料產生..."
         End
         Begin VB.Menu mnu3272 
            Caption         =   "&2.信用卡扣帳資料產生..."
         End
         Begin VB.Menu mnu3273 
            Caption         =   "&3.臨櫃代收資料產生..."
         End
         Begin VB.Menu mnu3274 
            Caption         =   "&4.金融格式轉帳資料產生..."
         End
         Begin VB.Menu mnu3275 
            Caption         =   "&5.ATM入帳折讓資料產生..."
         End
         Begin VB.Menu mnu3277 
            Caption         =   "&7.ACH轉帳扣款資料產生..."
         End
      End
      Begin VB.Menu mnu3280 
         Caption         =   "&8.收費單號調整作業"
         Begin VB.Menu mnu3281 
            Caption         =   "&1.收費單號調整(舊).."
         End
         Begin VB.Menu mnu3282 
            Caption         =   "&2.收費單號調整(新).."
         End
      End
      Begin VB.Menu mnu3290 
         Caption         =   "&9.扣款授權資料作業..."
         Begin VB.Menu mnu3291 
            Caption         =   "&1.亞太銀行扣繳授權作業..."
         End
         Begin VB.Menu mnu3292 
            Caption         =   "&2.ACH扣款授權資料作業..."
         End
         Begin VB.Menu mnu3293 
            Caption         =   "&3.郵局核印作業..."
         End
      End
      Begin VB.Menu mnu32A0 
         Caption         =   "&A.設備結轉作業..."
         Begin VB.Menu mnu32A1 
            Caption         =   "&1.設備結轉作業..."
         End
         Begin VB.Menu mnu32A2 
            Caption         =   "&2.設備結轉列印..."
         End
      End
      Begin VB.Menu mnu32B0 
         Caption         =   "&B.PPV結算作業..."
      End
      Begin VB.Menu mnu32C0 
         Caption         =   "&C.整批退費補償"
      End
      Begin VB.Menu mnu32D0 
         Caption         =   "&D.出單檢核"
      End
      Begin VB.Menu mnu32E0 
         Caption         =   "&E.VOD作業"
         Begin VB.Menu mnu32E1 
            Caption         =   "&1.贈送點數整批匯入"
         End
         Begin VB.Menu mnu32E2 
            Caption         =   "&2.VOD結算作業"
         End
      End
      Begin VB.Menu mnu32F0 
         Caption         =   "&F.收費項目合併作業"
      End
      Begin VB.Menu mnu32G0 
         Caption         =   "&G.一貫出單作業"
      End
      Begin VB.Menu mnu32H0 
         Caption         =   "&H.bb紅利整批贈送作業"
      End
   End
   Begin VB.Menu mnu3300 
      Caption         =   "&3.收款管理"
      Begin VB.Menu mnu3310 
         Caption         =   "&1.收費資料登錄"
         Begin VB.Menu mnu3311 
            Caption         =   "&1.收費單／劃撥單登錄..."
         End
         Begin VB.Menu mnu3312 
            Caption         =   "&2.媒體轉帳資料登錄..."
         End
         Begin VB.Menu mnu3313 
            Caption         =   "&3.信用卡扣帳資料登錄..."
         End
         Begin VB.Menu mnu3314 
            Caption         =   "&4.臨櫃代收資料登錄..."
         End
         Begin VB.Menu mnu3315 
            Caption         =   "&5.金融格式轉帳資料登錄..."
         End
         Begin VB.Menu mnu3316 
            Caption         =   "&6.ATM入帳折讓登錄..."
         End
         Begin VB.Menu mnu3317 
            Caption         =   "&7.ATM轉帳資料登錄..."
         End
         Begin VB.Menu mnu3318 
            Caption         =   "&8.櫃台繳款登錄作業..."
         End
         Begin VB.Menu mnu3319 
            Caption         =   "&9.ACH轉帳扣款登錄作業..."
         End
         Begin VB.Menu mnu331A 
            Caption         =   "&A.派工單所附收費資料登錄..."
         End
      End
      Begin VB.Menu mnu3320 
         Caption         =   "&2.收費報表查詢..."
      End
      Begin VB.Menu mnu3330 
         Caption         =   "&3.多條碼入帳作業"
      End
   End
   Begin VB.Menu mnu3400 
      Caption         =   "&4.未收款管理"
      Begin VB.Menu mnu3410 
         Caption         =   "&1.收費單未收原因登錄..."
      End
      Begin VB.Menu mnu3420 
         Caption         =   "&2.拆機戶整批作業..."
      End
      Begin VB.Menu mnu3430 
         Caption         =   "&3.整批入短收作業"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&0.離開"
   End
End
Attribute VB_Name = "frmSO3000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private GB As Object

Private Sub GetMainMenuPermission()
    On Error GoTo chkErr
    If garyGi(4) = "0" Then Exit Sub
    mnu3100.Enabled = GetUserPriv("SO3000", "SO3100")
    mnu3200.Enabled = GetUserPriv("SO3000", "SO3200")
    mnu3300.Enabled = GetUserPriv("SO3000", "SO3300")
    mnu3400.Enabled = GetUserPriv("SO3000", "SO3400")
    If mnu3100.Enabled Then
        mnu3110.Enabled = GetUserPriv("SO3100", "SO3110")
        mnu3120.Enabled = GetUserPriv("SO3100", "SO3120")
        mnu3130.Enabled = GetUserPriv("SO3100", "SO3130")
        mnu3140.Visible = (Len(GetRsValue("select Table_Name from all_tables where table_name='SO032TMP' And Owner='" & Replace(GetOwner, ".", "") & "'") & "") > 0)
        mnu3150.Visible = (Val(GetSystemParaItem("CombineBill", garyGi(9), "", "SO041", , True) & "") = 1) And GetUserPriv("SO3100", "SO3150")
        'mnu3140.Visible = mnu3140.Enabled
        If mnu3120.Enabled Then
            mnu3121.Enabled = GetUserPriv("SO3120", "SO3121")
            mnu3122.Enabled = GetUserPriv("SO3120", "SO3122")
            mnu3123.Enabled = GetUserPriv("SO3120", "SO3123")
            mnu3124.Enabled = GetUserPriv("SO3120", "SO3124")
            mnu3125.Enabled = GetUserPriv("SO3120", "SO3125")
        End If
    End If
    If mnu3200.Enabled Then
        mnu3210.Enabled = GetUserPriv("SO3200", "SO3210")
        mnu3220.Enabled = GetUserPriv("SO3200", "SO3220")
        mnu3230.Enabled = GetUserPriv("SO3200", "SO3230")
        mnu3240.Enabled = GetUserPriv("SO3200", "SO3240")
        mnu3250.Enabled = GetUserPriv("SO3200", "SO3250")
        mnu3260.Enabled = GetUserPriv("SO3200", "SO3260")
        mnu3270.Enabled = GetUserPriv("SO3200", "SO3270")
        mnu3280.Enabled = GetUserPriv("SO3200", "SO3280")
        mnu3290.Enabled = GetUserPriv("SO3200", "SO3290")
        mnu32A0.Enabled = GetUserPriv("SO3200", "SO32A0")
        mnu32B0.Enabled = GetUserPriv("SO3200", "SO32B0")
        mnu32C0.Enabled = GetUserPriv("SO3200", "SO32C0")
        mnu32D0.Enabled = GetUserPriv("SO3200", "SO32D0")
        mnu32E0.Enabled = GetUserPriv("SO3200", "SO32E0")
        mnu32E2.Enabled = GetUserPriv("SO32E0", "SO32E02")
        mnu32F0.Enabled = GetUserPriv("SO3200", "SO32F0")
        mnu32G0.Enabled = GetUserPriv("SO3200", "SO32G0")
        '#7020 &H.bb紅利整批贈送作業
        mnu32H0.Enabled = GetUserPriv("SO3200", "SO32H0") And (Val(GetRsValue("Select Nvl(StartbbCash,0) From " & GetOwner & "SO041")) = 1)
        '#6087 測試不OK增加權限控管 By Kin 2011/11/04
        If mnu3280.Enabled Then
            mnu3281.Enabled = GetUserPriv("SO3280", "SO3281")
            mnu3282.Enabled = GetUserPriv("SO3280", "SO3282")
        End If
        
        If mnu3250.Enabled Then
            mnu3251.Enabled = GetUserPriv("SO3250", "SO3251")
            mnu3252.Enabled = GetUserPriv("SO3250", "SO3252")
            mnu3253.Enabled = GetUserPriv("SO3250", "SO3253")
            mnu3254.Enabled = GetUserPriv("SO3250", "SO3254")
        End If
        If mnu3260.Enabled Then
            mnu3261.Enabled = GetUserPriv("SO3260", "SO3261")
            mnu3262.Enabled = GetUserPriv("SO3260", "SO3262")
            mnu3263.Enabled = GetUserPriv("SO3260", "SO3263")
        End If
        If mnu3270.Enabled Then
            mnu3271.Enabled = GetUserPriv("SO3270", "SO3271")
            mnu3272.Enabled = GetUserPriv("SO3270", "SO3272")
            mnu3273.Enabled = GetUserPriv("SO3270", "SO3273")
            mnu3274.Enabled = GetUserPriv("SO3270", "SO3274")
            mnu3275.Enabled = GetUserPriv("SO3270", "SO3275")
            mnu3277.Enabled = GetUserPriv("SO3270", "SO3277")
        End If
        If mnu3290.Enabled Then
            mnu3291.Enabled = GetUserPriv("SO3290", "SO3291")
            mnu3292.Enabled = GetUserPriv("SO3290", "SO3292")
        End If
        If mnu32A0.Enabled Then
            mnu32A1.Enabled = GetUserPriv("SO32A0", "SO32A1")
            mnu32A2.Enabled = GetUserPriv("SO32A0", "SO32A2")
        End If
        
    End If
    If mnu3300.Enabled Then
        mnu3310.Enabled = GetUserPriv("SO3300", "SO3310")
        mnu3320.Enabled = GetUserPriv("SO3300", "SO3320")
        If mnu3310.Enabled Then
            mnu3311.Enabled = GetUserPriv("SO3310", "SO3311")
            mnu3312.Enabled = GetUserPriv("SO3310", "SO3312")
            mnu3313.Enabled = GetUserPriv("SO3310", "SO3313")
            mnu3314.Enabled = GetUserPriv("SO3310", "SO3314")
            mnu3315.Enabled = GetUserPriv("SO3310", "SO3315")
            mnu3316.Enabled = GetUserPriv("SO3310", "SO3316")
            mnu3317.Enabled = GetUserPriv("SO3310", "SO3317")
            mnu3318.Enabled = GetUserPriv("SO3310", "SO3318")
            mnu3319.Enabled = GetUserPriv("SO3310", "SO3319")
            mnu331A.Enabled = GetUserPriv("SO3310", "SO331A")
            
        End If
     End If
     If mnu3400.Enabled Then
        mnu3410.Enabled = GetUserPriv("SO3400", "SO3410")
        mnu3420.Enabled = GetUserPriv("SO3400", "SO3420")
        mnu3430.Enabled = GetUserPriv("SO3400", "SO3430")
     End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "GetMainMenuPermission")
End Sub



Private Sub Memo()
    Dim rs As New ADODB.Recordset
    Dim i
    Dim strRS As String
    Dim strType As String
        strRS = InputBox("rs Name:")
        Call GetRS(rs, "Select * From So033 where 1=0")
        For i = 0 To rs.Fields.Count - 1
            Debug.Print rs.Fields(i).Name & ","
        Next
    'giLongV = 0
    'giStringV = 1
    'giDateV = 2
        'custid =131 number
        'billno=200 varchar2
        'realstartdate=135 date
        'servicetype=129 char
        'long =201
        For i = 0 To rs.Fields.Count - 1
            Select Case rs.Fields(i).Type
                Case 131
                    strType = "giLongV"
                Case 200 Or 129 Or 201
                    strType = ""
                Case 135
                    strType = "giDateV"
            End Select
            Debug.Print "getnullstring(" & strRS & "(" & rs.Fields(i).Name & ")" & IIf(strType = "", "", strType & ",") & ")"
        Next
End Sub

Private Sub Command1_Click()
    Dim rs As New ADODB.Recordset
     Get_RS_From_Txt "C:\Documents and Settings\Administrator\Local Settings\Temp", "SO5780.txt", rs

'    Call PrintRpt(, "Report1.rpt", , , "Select * from so3320A1", , , , "Tmp0000.mdb")
'    Dim rs As New ADODB.Recordset
'    If cnn.State = 0 Then cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Gird\V400\ReportFile.mdb;Persist Security Info=False"
'    Call GetAutoCreateRptRS(rs)
'
'    rs.AddNew
'    rs.Fields("TableName") = "SO017"
'    rs.Fields("SQLQuery") = "SO017.MduId = '0003'"
'    rs.Update
'    AutoCreateReport cnn, "SO1300L", rs, cnn
    MsgBox rs.GetString
End Sub



Private Sub Command2_Click()
    Dim o As New frmSO32G0A
        o.Show vbModal
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        If KeyCode = vbKeyEscape Then
            Call mnuExit_Click
        End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    If Not Alfa Then
        On Error Resume Next
        Set GB = CreateObject("GICMIS_BOJ.OBJ")
        Set GB.Form3000 = Me
        GB.FormGICMIS.Hide
        Command1.Visible = Alfa
    Else
        ChkMonitorPixel
    End If
    On Error GoTo chkErr
    Me.Show
    Me.Refresh
    Me.Enabled = False
        Call GetGlobal
        Call GetMainMenuPermission
    LoadImage imgPic
    Me.Caption = Me.Caption & garyGi(13)
    Me.Enabled = True
        'MsgBox garyGi(15)
    Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        Set gcnGi = Nothing
        If GB Is Nothing Then Exit Sub
        Set GB.Form3000 = Nothing
        GB.FormGICMIS.Enabled = True
        GB.FormGICMIS.SetFocus
        GB.FormGICMIS.Show
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        imgPic.Move (Me.Width - imgPic.Width) / 2, (Me.Height - imgPic.Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)
End Sub

Private Sub mnu3110_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3110A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3121_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        With frmSO3121A
                .uForm = "SO3121"
                .Show vbModal
        End With
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3122_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        With frmSO3122A
                .uForm = "SO3122"
                .Show vbModal
        End With
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3123_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        With frmSO3123A
                .uForm = "SO3123"
                .Show vbModal
        End With
        Screen.MousePointer = vbDefault
    'MsgBox mnu3123.Caption
End Sub

Private Sub mnu3124_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        With frmSO3124A
                .uForm = "SO3124"
                .Show vbModal
        End With
        Screen.MousePointer = vbDefault
    'MsgBox mnu3124.Caption
End Sub

Private Sub mnu3125_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        With frmSO3125A
                .uForm = "SO3125"
                .Show vbModal
        End With
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3130_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3130A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3140_Click()
    On Error Resume Next
        frmSO3140A.Show vbModal
End Sub

Private Sub mnu3150_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3150A.Show vbModal
        Screen.MousePointer = vbDefault

End Sub

Private Sub mnu3210_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3210A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3241_Click()
    On Error Resume Next
    '    MsgBox "from3241"
End Sub

Private Sub mnu3220_Click()
    Screen.MousePointer = vbHourglass
    frmSO3220A.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3230_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        With frmSO3230A
'                .uForm = "SO3230A"
                 .Show vbModal
        End With
        Screen.MousePointer = vbDefault
        
End Sub

Private Sub mnu3240_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3240A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3251_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3251A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3252_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3252A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3253_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3253A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3254_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3254A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3261_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3261A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3262_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3262A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3263_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3263A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3271_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3271A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3272_Click()
    On Error Resume Next
    Dim intTBC As Integer
        Screen.MousePointer = vbHourglass
        intTBC = RPxx(gcnGi.Execute("SELECT CreditCardTBCMode FROM " & _
        GetOwner & "SO041 WHERE CompCode = " & garyGi(9)).GetString)
        Select Case intTBC
            Case 0
                frmSO3272A.Show vbModal
            Case 1
                frmSO3272B.Show vbModal
        End Select
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3273_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3273A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3274_Click()
    On Error Resume Next
    Dim obj As Object
        Set obj = CreateObject("csBankCenter4.clsBankCenterShow")
        Call obj.FormOutShow(CStr(PutGlobal), gcnGi)
        Set obj = Nothing
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3277_Click()
    On Error Resume Next
    Dim obj As Object
        Set obj = CreateObject("csACHTranRefer.clsACHTranReferShow")
        Call obj.FormOutShow(CStr(PutGlobal), gcnGi)
        Set obj = Nothing
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3281_Click()
  On Error Resume Next
    Screen.MousePointer = vbHourglass
    frmSO3280A.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3282_Click()
  On Error Resume Next
    Screen.MousePointer = vbHourglass
    frmSO3282A.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3291_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3291A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3292_Click()
    On Error Resume Next
    Dim obj As Object
    If Val(GetRsValue("Select Nvl(ACHCustID,0) From " & GetOwner & _
                            "SO041 Where SysID=" & gCompCode, gcnGi)) = 0 Then
        Set obj = CreateObject("csACHauth.clsACHauthShow")
    Else
        Set obj = CreateObject("csACHauth2.clsACHauthShow")
    End If
    Call obj.FormOutShow(CStr(PutGlobal), gcnGi)
    Set obj = Nothing
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3293_Click()
   On Error Resume Next
    Dim obj As Object
'    If Val(GetRsValue("Select Nvl(ACHCustID,0) From " & GetOwner & _
'                            "SO041 Where SysID=" & gCompCode, gcnGi)) = 0 Then
'        Set obj = CreateObject("csACHauth.clsACHauthShow")
'    Else
'        Set obj = CreateObject("csACHauth2.clsACHauthShow")
'    End If
     If Val(GetRsValue("Select Nvl(StartPost,0) From " & GetOwner & _
                            "SO041 Where SysID=" & gCompCode, gcnGi)) = 0 Then
        MsgBox "參數尚未設定，無法使用", vbInformation, "訊息"
        Exit Sub
     End If
    Set obj = CreateObject("csPosACHauth.clsACHauthShow")
    Call obj.FormOutShow(CStr(PutGlobal), gcnGi)
    Set obj = Nothing
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu32A1_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO32A1A.Show vbModal
        Screen.MousePointer = vbDefault

End Sub

Private Sub mnu32A2_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO32A2A.Show vbModal
        Screen.MousePointer = vbDefault

End Sub

Private Sub mnu32B0_Click()
        On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO1156A.Show vbModal
        Screen.MousePointer = vbDefault
    
End Sub

Private Sub mnu32C0_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO32C0A.Show vbModal
        Screen.MousePointer = vbDefault
        
End Sub

Private Sub mnu32D0_Click()
  On Error Resume Next
    Screen.MousePointer = vbHourglass
    frmSO32D0A.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnu32E1_Click()
  On Error Resume Next
    With frmSO32E01
        .Show vbModal
    End With
End Sub

Private Sub mnu32E2_Click()
  On Error Resume Next
    Screen.MousePointer = vbHourglass
'    With frmSO1193A
'        .uCompCode = gCompCode
'        .uFromBat = True
'        .Show 1
'    End With
    Dim objVod As Object
    Dim strPara35 As String
    Dim strPara36 As String
    Dim rs043 As New ADODB.Recordset
    If Not GetRS(rs043, "SELECT PARA35,PARA36 FROM " & GetOwner & "SO043 WHERE SERVICETYPE = 'D'", gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Sub
    If Not rs043.EOF Then
        strPara35 = rs043("Para35") & ""
        strPara36 = rs043("Para36") & ""
    End If
    Set objVod = CreateObject("csVODBillClose.clsInterface")
    With objVod
        objVod.uCompCode = gCompCode
        objVod.uFromBat = True
        Set objVod.ugcnGi = gcnGi
        objVod.ugaryGi = Join(garyGi, Chr(9))
        objVod.uPara35 = strPara35
        objVod.uPara36 = strPara36
        objVod.Action
    End With
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnu32F0_Click()
  On Error Resume Next
    Screen.MousePointer = vbHourglass
    frmSO32F0A.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnu32G0_Click()
    On Error Resume Next
'        Dim o As New frmSO32G0A
'            o.Show vbModal
    If CallWebPage("SO32B0A") = True Then Me.Enabled = True: Exit Sub
    Dim o As New frmSO32G0A
    o.Show vbModal
End Sub

Private Sub mnu32H0_Click()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    frmSO32H0A.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3311_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3311A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3312_Click()
        On Error Resume Next
        'Exit Sub
        '先不用
        Screen.MousePointer = vbHourglass
        frmSo3312A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3313_Click()
  On Error Resume Next
        Dim intTBC As Integer
        Screen.MousePointer = vbHourglass
        intTBC = RPxx(gcnGi.Execute("SELECT CreditCardTBCMode FROM " & _
        GetOwner & "SO041 WHERE CompCode = " & garyGi(9)).GetString)
        Select Case intTBC
            Case 0
                frmSO3313A.Show vbModal
            Case 1
                If InputBox("請選擇使用格式1.舊格式,2.新格式", "提示", 2) = 1 Then
                    frmSO3313A.Show vbModal
                Else
                    frmSO3313B.Show vbModal
                End If
        End Select
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3314_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSo3314A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3315_Click()
    On Error Resume Next
    Dim obj As Object
        'Screen.MousePointer = vbHourglass
        Set obj = CreateObject("csBankCenter4.clsBankCenterShow")
        Call obj.FormInShow(CStr(PutGlobal), gcnGi)
        Set obj = Nothing
        'frmSO3315A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3317_Click()
On Error Resume Next
  Screen.MousePointer = vbHourglass
  frmSO3317A.Show vbModal
  Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3318_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3318A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3319_Click()
    On Error Resume Next
    Dim obj As Object
        Set obj = CreateObject("csACHTranRefer.clsACHTranReferShow")
        Call obj.FormInShow(CStr(PutGlobal), gcnGi)
        Set obj = Nothing
        Screen.MousePointer = vbDefault

End Sub

Private Sub mnu331A_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO331AA.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3320_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        With frmSO3320A
                .uForm = "SO3320A"
                '.Show , Me
                .Show vbModal
        End With
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3330_Click()
 On Error Resume Next
    
    Dim strData As String, strFilePath As String
    strFilePath = ReadFromINI2("SOEXEPATH") & "\SONEW\CableSoft.SO.Win.EntryForm.exe"
    
    If ChkDirExist(strFilePath) Then
       strData = " CableSoft.SO.Win.Billing.MultiBarcodeIn.dll CableSoft.SO.Win.Billing.MultiBarcodeIn.SO3330A" & " " & gCompCode & " " & garyGi(4) & " " & garyGi(0) & " " & garyGi(1) & ""
       Shell strFilePath & strData, vbNormalFocus
    Else
       MsgBox "多條碼入帳檔案不存在!!", vbExclamation, "訊息": Exit Sub
    End If
End Sub

Private Sub mnu3410_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3410A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3420_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3420A.Show vbModal
        Screen.MousePointer = vbDefault
End Sub

Private Sub mnu3430_Click()
    On Error Resume Next
        Screen.MousePointer = vbHourglass
        frmSO3430A.Show vbModal
        Screen.MousePointer = vbDefault

End Sub

Private Sub mnuExit_Click()
    On Error Resume Next
        Unload Me
End Sub
