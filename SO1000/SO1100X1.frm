VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO1100X1 
   BorderStyle     =   1  '單線固定
   Caption         =   "贈送BB紅利 [SO1100X1]"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5160
   Icon            =   "SO1100X1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&X)"
      Height          =   345
      Left            =   3750
      TabIndex        =   5
      Top             =   2370
      Width           =   810
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2.確定"
      Height          =   345
      Left            =   480
      Picture         =   "SO1100X1.frx":0442
      TabIndex        =   4
      Top             =   2370
      Width           =   780
   End
   Begin VB.Frame fraEdit 
      Height          =   2205
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   5010
      Begin VB.TextBox txtAddBonus 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   0
         Text            =   "0"
         Top             =   300
         Width           =   870
      End
      Begin VB.TextBox txtGiveReason 
         Height          =   660
         Left            =   1200
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   3
         Top             =   1380
         Width           =   3735
      End
      Begin prjGiList.GiList gilBonusKind 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1020
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth1       =   500
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilUpdEn 
         Height          =   315
         Left            =   1350
         TabIndex        =   7
         Top             =   2895
         Visible         =   0   'False
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
      Begin Gi_Date.GiDate gdaBonusStopDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   630
         Width           =   1065
         _ExtentX        =   1879
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
      Begin VB.Label lblbonus 
         AutoSize        =   -1  'True
         Caption         =   "紅利點數:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   390
         TabIndex        =   12
         Top             =   330
         Width           =   765
      End
      Begin VB.Label lblGiveReason 
         AutoSize        =   -1  'True
         Caption         =   "贈送原因:"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   390
         TabIndex        =   11
         Top             =   1590
         Width           =   765
      End
      Begin VB.Label lblBonusStopDate 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "贈點到期日:"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   210
         TabIndex        =   10
         Top             =   690
         Width           =   945
      End
      Begin VB.Label lblUpdEn 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "人員:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   900
         TabIndex        =   9
         Top             =   2925
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblBonusKind 
         AutoSize        =   -1  'True
         Caption         =   "積點種類:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   390
         TabIndex        =   8
         Top             =   1080
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmSO1100X1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsData2 As Recordset
Private fBBVodAccountId As String
Private IsSave As Boolean
Private Sub subGil()
    On Error GoTo ChkErr
        SetgiList gilBonusKind, "CodeNo", "Description", "CD108", , , , , 5, 20, " Where RefNo=2 And Nvl(StopFlag,0)=0", True
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub
Private Sub DefaultValue()
    On Error GoTo ChkErr
        Dim BounsCode  As String
        BounsCode = GetRsValue("Select CodeNo From " & GetOwner & "CD108 Where Nvl(StopFlag,0) = 0 And RefNo = 2") & ""
        gilBonusKind.SetCodeNo BounsCode
        gilBonusKind.Query_Description
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "DefaultValue")
End Sub
Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If Len(txtAddBonus.Text & "") = 0 Or txtAddBonus.Text = "0" Then
        MsgMustBe ("紅利點數")
        txtAddBonus.SetFocus
        Exit Function
    End If
    If Len(gdaBonusStopDate.GetValue & "") > 0 Then
        If Val(gdaBonusStopDate.GetValue) < Val(Format(Now, "yyyymmdd")) Then
            MsgBox "贈點到期日不可小於今天日期！", vbOKOnly, "警告"
            gdaBonusStopDate.SetFocus
            Exit Function
        End If
    End If
    IsDataOk = True
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOK")
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    Dim sqlSO193 As String
    Dim sqlSO004J As String
    'Dim sumSQL As String
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    If Not IsDataOk Then Me.MousePointer = vbDefault: Me.Enabled = True: Exit Sub
    Dim rsSO193 As New ADODB.Recordset
    Dim rsSO004J As New ADODB.Recordset
    Dim aTotalBonus As Double
    gcnGi.BeginTrans
      sqlSO193 = "Select * From " & GetOwner & "SO193 Where bbAccountID = '" & fBBVodAccountId & "'"
      sqlSO004J = "Select * From " & GetOwner & "SO004J Where bbAccountID = '" & fBBVodAccountId & "'"
      If Not GetRS(rsSO193, sqlSO193, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
      If Not GetRS(rsSO004J, sqlSO004J, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
      rsSO193.AddNew
      With rsSO193
        .Fields("bbAccountID") = fBBVodAccountId
        .Fields("SeqNo") = Get_SeqNo_Seq
        .Fields("Savedate") = RightNow
        .Fields("bonus") = txtAddBonus.Text
        .Fields("MinusType") = 2
        If Len(gdaBonusStopDate.Text & "") > 0 Then
            .Fields("bonusstopdate") = gdaBonusStopDate.Text
        End If
        If Len(txtGiveReason.Text & "") > 0 Then
            .Fields("NOTE") = txtGiveReason.Text
        End If
        .Fields("UpdTime") = RightNow
        .Fields("UpdEn") = garyGi(1)
        .Update
      End With
      
'        sumSQL = "Select Sum(bonus)-Sum(Usedbonus) From " & GetOwner & "SO193 " & _
'                        " Where bbAccountID ='" & fBBVodAccountId & "' " & _
'                        " And Nvl(bonusstopdate,to_date('29991231235959','YYYYMMDDHH24MISS')) >= " & _
'                        " To_Date('" & Format(RightNow, "yyyymmddhhmmss") & "','YYYYMMDDHH24MISS')"
'        aTotalBonus = Val(GetRsValue(sumSQL, gcnGi))
        If rsSO004J.RecordCount = 0 Then
            rsSO004J.AddNew
            rsSO004J("bbAccountID") = fBBVodAccountId
        End If
        rsSO004J("TotalBonus") = Val(rsSO004J("TotalBonus") & "") + Val(txtAddBonus.Text & "")
        rsSO004J.Update
    gcnGi.CommitTrans
    On Error Resume Next
    CloseRecordset rsSO004J
    CloseRecordset rsSO193
    Unload Me
    Me.MousePointer = vbDefault
    Me.Enabled = True
    IsSave = True
    Exit Sub
ChkErr:
  gcnGi.RollbackTrans
   Me.MousePointer = vbDefault
   Me.Enabled = True
  Call ErrSub(Me.Name, "cmdOk_Click")
End Sub
Private Function Get_SeqNo_Seq() As String
  On Error GoTo ChkErr
    Get_SeqNo_Seq = Format(Now, "yyyymmdd") & Right("000000000000" & _
                    RPxx(gcnGi.Execute("SELECT " & GetOwner & "S_SO193.NEXTVAL FROM DUAL").GetString & ""), 12)
    
  Exit Function
ChkErr:
    ErrSub Me.Name, "Get_SeqNo_Seq"
End Function
Private Sub FunctionKey(KeyCode As Integer, Shift As Integer)
    On Error GoTo ChkErr
    
        If KeyCode = vbKeyF2 Then Call cmdOK_Click
        If KeyCode = vbKeyEscape Then cmdCancel_Click
'        If KeyCode = vbKeyF3 And cmdFind.Enabled Then Call cmdFind_Click
'        If KeyCode = vbKeyF4 And cmdEdit.Enabled Then Call cmdEdit_Click
'        If KeyCode = vbKeyEscape And cmdCancel.Enabled Then Call cmdCancel_Click
'        If KeyCode = vbKeyX And Shift = 2 Then cmdCancel.Value = True
'
'        If KeyCode = vbKeyUp Then
'            KeyCode = 0
'            SendKeys "+{Tab}"
'        End If
'        If KeyCode = vbKeyDown Then
'            KeyCode = 0
'            SendKeys "{Tab}"
'        End If
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    FunctionKey KeyCode, Shift
End Sub

Private Sub Form_Load()
    IsSave = False
    Me.MousePointer = vbHourglass
    Call subGil
    Call DefaultValue
    Me.MousePointer = vbDefault
End Sub
Public Property Let BBVodAccountId(ByVal vData As String)
  fBBVodAccountId = vData
End Property
Public Property Get uSave() As Boolean
  uSave = IsSave
End Property

