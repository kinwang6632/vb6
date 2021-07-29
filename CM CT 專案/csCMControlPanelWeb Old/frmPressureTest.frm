VERSION 5.00
Begin VB.Form frmPressureTest 
   Caption         =   "機上盒壓力測試"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10320
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPressureTest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   10320
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdCloseForm 
      Caption         =   "關闢測試視窗(&C)"
      Height          =   375
      Left            =   120
      TabIndex        =   56
      Top             =   6060
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8550
      TabIndex        =   55
      Top             =   6060
      Width           =   1605
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "執行(&E)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6690
      TabIndex        =   54
      Top             =   6060
      Width           =   1605
   End
   Begin VB.Frame frmData 
      Caption         =   "執行情形"
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   2895
      Index           =   1
      Left            =   90
      TabIndex        =   27
      Top             =   3030
      Width           =   10065
      Begin VB.CheckBox chkOpen 
         Caption         =   "開機    傳送筆數"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   53
         Top             =   330
         Value           =   1  '核取
         Width           =   1665
      End
      Begin VB.TextBox txtOpenCount 
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   1
         Left            =   2130
         MaxLength       =   3
         TabIndex        =   52
         Top             =   270
         Width           =   915
      End
      Begin VB.CheckBox chkClose 
         Caption         =   "關機    傳送筆數"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   51
         Top             =   750
         Value           =   1  '核取
         Width           =   1665
      End
      Begin VB.TextBox txtCloseCount 
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   1
         Left            =   2130
         MaxLength       =   3
         TabIndex        =   50
         Text            =   "   "
         Top             =   675
         Width           =   915
      End
      Begin VB.CheckBox chkStop 
         Caption         =   "停機    傳送筆數"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   49
         Top             =   1170
         Value           =   1  '核取
         Width           =   1665
      End
      Begin VB.TextBox txtStopCount 
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   1
         Left            =   2130
         MaxLength       =   3
         TabIndex        =   48
         Text            =   "   "
         Top             =   1095
         Width           =   915
      End
      Begin VB.CheckBox chkReOpen 
         Caption         =   "復機    傳送筆數"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   47
         Top             =   1590
         Value           =   1  '核取
         Width           =   1665
      End
      Begin VB.TextBox txtReopenCount 
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   1
         Left            =   2130
         MaxLength       =   3
         TabIndex        =   46
         Text            =   "   "
         Top             =   1500
         Width           =   915
      End
      Begin VB.CheckBox chkDelete 
         Caption         =   "註銷    傳送筆數"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   45
         Top             =   2010
         Value           =   1  '核取
         Width           =   1665
      End
      Begin VB.TextBox txtDeleteCount 
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   1
         Left            =   2130
         MaxLength       =   3
         TabIndex        =   44
         Text            =   "   "
         Top             =   1920
         Width           =   915
      End
      Begin VB.TextBox txtOpenChCount 
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   1
         Left            =   5460
         MaxLength       =   3
         TabIndex        =   43
         Text            =   "   "
         Top             =   270
         Width           =   915
      End
      Begin VB.CheckBox chkOpenCh 
         Caption         =   "開頻道      傳送筆數"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   3480
         TabIndex        =   42
         Top             =   330
         Value           =   1  '核取
         Width           =   1965
      End
      Begin VB.TextBox txtCloseChCount 
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   1
         Left            =   5460
         MaxLength       =   3
         TabIndex        =   41
         Text            =   "   "
         Top             =   675
         Width           =   915
      End
      Begin VB.CheckBox chkCloseCh 
         Caption         =   "關頻道      傳送筆數"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   3480
         TabIndex        =   40
         Top             =   750
         Value           =   1  '核取
         Width           =   1965
      End
      Begin VB.TextBox txtStopChCount 
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   1
         Left            =   5460
         MaxLength       =   3
         TabIndex        =   39
         Text            =   "   "
         Top             =   1095
         Width           =   915
      End
      Begin VB.CheckBox chkStopCh 
         Caption         =   "暫停頻道  傳送筆數"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   3480
         TabIndex        =   38
         Top             =   1170
         Value           =   1  '核取
         Width           =   1965
      End
      Begin VB.TextBox txtReOpenChCount 
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   1
         Left            =   5460
         MaxLength       =   3
         TabIndex        =   37
         Text            =   "   "
         Top             =   1500
         Width           =   915
      End
      Begin VB.CheckBox chkReopenCh 
         Caption         =   "恢復頻道  傳送筆數"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   3480
         TabIndex        =   36
         Top             =   1590
         Value           =   1  '核取
         Width           =   1965
      End
      Begin VB.TextBox txtExtendRangeCount 
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   1
         Left            =   5460
         MaxLength       =   3
         TabIndex        =   35
         Text            =   "   "
         Top             =   1920
         Width           =   915
      End
      Begin VB.CheckBox chkExtendRange 
         Caption         =   "期限延展  傳送筆數"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   3480
         TabIndex        =   34
         Top             =   2010
         Value           =   1  '核取
         Width           =   1965
      End
      Begin VB.TextBox txtChPrivCopyCount 
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   1
         Left            =   5460
         MaxLength       =   3
         TabIndex        =   33
         Text            =   "   "
         Top             =   2340
         Width           =   915
      End
      Begin VB.CheckBox chkChPrivCopy 
         Caption         =   "權限複製  傳送筆數"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   3480
         TabIndex        =   32
         Top             =   2430
         Value           =   1  '核取
         Width           =   1965
      End
      Begin VB.CheckBox chkSendMsg 
         Caption         =   "傳送訊息           筆數"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   6780
         TabIndex        =   31
         Top             =   750
         Value           =   1  '核取
         Width           =   1965
      End
      Begin VB.TextBox txtSendMsg 
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   1
         Left            =   8760
         MaxLength       =   3
         TabIndex        =   30
         Text            =   "   "
         Top             =   675
         Width           =   915
      End
      Begin VB.CheckBox chkClearPinCode 
         Caption         =   "清除親子密碼  筆數"
         ForeColor       =   &H00808080&
         Height          =   195
         Index           =   1
         Left            =   6780
         TabIndex        =   29
         Top             =   330
         Value           =   1  '核取
         Width           =   1965
      End
      Begin VB.TextBox txtClearPinCode 
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   1
         Left            =   8760
         MaxLength       =   3
         TabIndex        =   28
         Text            =   "   "
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.Frame frmData 
      Caption         =   "機上盒開關選項選項"
      ForeColor       =   &H00FF0000&
      Height          =   2895
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   10065
      Begin VB.TextBox txtClearPinCode 
         Height          =   345
         Index           =   0
         Left            =   8760
         MaxLength       =   3
         TabIndex        =   26
         Text            =   "0"
         Top             =   270
         Width           =   915
      End
      Begin VB.CheckBox chkClearPinCode 
         Caption         =   "清除親子密碼  筆數"
         Height          =   195
         Index           =   0
         Left            =   6780
         TabIndex        =   25
         Top             =   330
         Width           =   1965
      End
      Begin VB.TextBox txtSendMsg 
         Height          =   345
         Index           =   0
         Left            =   8760
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "0"
         Top             =   675
         Width           =   915
      End
      Begin VB.CheckBox chkSendMsg 
         Caption         =   "傳送訊息           筆數"
         Height          =   195
         Index           =   0
         Left            =   6780
         TabIndex        =   23
         Top             =   750
         Width           =   1965
      End
      Begin VB.CheckBox chkChPrivCopy 
         Caption         =   "權限複製  傳送筆數"
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   22
         Top             =   2430
         Value           =   1  '核取
         Width           =   1965
      End
      Begin VB.TextBox txtChPrivCopyCount 
         Height          =   345
         Index           =   0
         Left            =   5460
         MaxLength       =   3
         TabIndex        =   21
         Text            =   "10"
         Top             =   2340
         Width           =   915
      End
      Begin VB.CheckBox chkExtendRange 
         Caption         =   "期限延展  傳送筆數"
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   20
         Top             =   2010
         Value           =   1  '核取
         Width           =   1965
      End
      Begin VB.TextBox txtExtendRangeCount 
         Height          =   345
         Index           =   0
         Left            =   5460
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "10"
         Top             =   1920
         Width           =   915
      End
      Begin VB.CheckBox chkReopenCh 
         Caption         =   "恢復頻道  傳送筆數"
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   18
         Top             =   1590
         Value           =   1  '核取
         Width           =   1965
      End
      Begin VB.TextBox txtReOpenChCount 
         Height          =   345
         Index           =   0
         Left            =   5460
         MaxLength       =   3
         TabIndex        =   17
         Text            =   "10"
         Top             =   1500
         Width           =   915
      End
      Begin VB.CheckBox chkStopCh 
         Caption         =   "暫停頻道  傳送筆數"
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   16
         Top             =   1170
         Value           =   1  '核取
         Width           =   1965
      End
      Begin VB.TextBox txtStopChCount 
         Height          =   345
         Index           =   0
         Left            =   5460
         MaxLength       =   3
         TabIndex        =   15
         Text            =   "10"
         Top             =   1095
         Width           =   915
      End
      Begin VB.CheckBox chkCloseCh 
         Caption         =   "關頻道      傳送筆數"
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   14
         Top             =   750
         Value           =   1  '核取
         Width           =   1965
      End
      Begin VB.TextBox txtCloseChCount 
         Height          =   345
         Index           =   0
         Left            =   5460
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "10"
         Top             =   675
         Width           =   915
      End
      Begin VB.CheckBox chkOpenCh 
         Caption         =   "開頻道      傳送筆數"
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   12
         Top             =   330
         Value           =   1  '核取
         Width           =   1965
      End
      Begin VB.TextBox txtOpenChCount 
         Height          =   345
         Index           =   0
         Left            =   5460
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "10"
         Top             =   270
         Width           =   915
      End
      Begin VB.TextBox txtDeleteCount 
         Height          =   345
         Index           =   0
         Left            =   2130
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "10"
         Top             =   1920
         Width           =   915
      End
      Begin VB.CheckBox chkDelete 
         Caption         =   "註銷    傳送筆數"
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   9
         Top             =   2010
         Value           =   1  '核取
         Width           =   1665
      End
      Begin VB.TextBox txtReopenCount 
         Height          =   345
         Index           =   0
         Left            =   2130
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "10"
         Top             =   1500
         Width           =   915
      End
      Begin VB.CheckBox chkReOpen 
         Caption         =   "復機    傳送筆數"
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   7
         Top             =   1590
         Value           =   1  '核取
         Width           =   1665
      End
      Begin VB.TextBox txtStopCount 
         Height          =   345
         Index           =   0
         Left            =   2130
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "10"
         Top             =   1095
         Width           =   915
      End
      Begin VB.CheckBox chkStop 
         Caption         =   "停機    傳送筆數"
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   5
         Top             =   1170
         Value           =   1  '核取
         Width           =   1665
      End
      Begin VB.TextBox txtCloseCount 
         Height          =   345
         Index           =   0
         Left            =   2130
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "10"
         Top             =   675
         Width           =   915
      End
      Begin VB.CheckBox chkClose 
         Caption         =   "關機    傳送筆數"
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   3
         Top             =   750
         Value           =   1  '核取
         Width           =   1665
      End
      Begin VB.TextBox txtOpenCount 
         Height          =   345
         Index           =   0
         Left            =   2130
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "10"
         Top             =   270
         Width           =   915
      End
      Begin VB.CheckBox chkOpen 
         Caption         =   "開機    傳送筆數"
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   1
         Top             =   330
         Value           =   1  '核取
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmPressureTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objForm11() As New frmPressureChild
Dim objForm12() As New frmPressureChild
Dim objForm13() As New frmPressureChild
Dim objForm14() As New frmPressureChild
Dim objForm15() As New frmPressureChild

Dim objForm21() As New frmPressureChild
Dim objForm22() As New frmPressureChild
Dim objForm23() As New frmPressureChild
Dim objForm24() As New frmPressureChild
Dim objForm25() As New frmPressureChild
Dim objForm26() As New frmPressureChild

Dim objForm31() As New frmPressureChild
Dim objForm32() As New frmPressureChild

Dim ColForm As New Collection

Private Sub cmdCloseForm_Click()
    Dim intLoop As Integer, intLoop2 As Integer
        On Error Resume Next
        For intLoop = 0 To UBound(objForm11())
            Unload objForm11(intLoop)
        Next
        For intLoop = 0 To UBound(objForm12())
            Unload objForm12(intLoop)
        Next
        For intLoop = 0 To UBound(objForm13())
            Unload objForm13(intLoop)
        Next
        For intLoop = 0 To UBound(objForm14())
            Unload objForm14(intLoop)
        Next
        For intLoop = 0 To UBound(objForm15())
            Unload objForm15(intLoop)
        Next
        For intLoop = 0 To UBound(objForm21())
            Unload objForm21(intLoop)
        Next
        For intLoop = 0 To UBound(objForm22())
            Unload objForm22(intLoop)
        Next
        For intLoop = 0 To UBound(objForm23())
            Unload objForm23(intLoop)
        Next
        For intLoop = 0 To UBound(objForm24())
            Unload objForm24(intLoop)
        Next
        For intLoop = 0 To UBound(objForm25())
            Unload objForm25(intLoop)
        Next
        For intLoop = 0 To UBound(objForm26())
            Unload objForm26(intLoop)
        Next
        For intLoop = 0 To UBound(objForm31())
            Unload objForm31(intLoop)
        Next
        For intLoop = 0 To UBound(objForm32())
            Unload objForm32(intLoop)
        Next
End Sub

Private Sub cmdExecute_Click()
    Call InitData
    Call AddNewForm
End Sub

Private Sub InitData()
    Dim objControl As Object
    Dim colValue As New Collection
    On Error Resume Next
        For Each objControl In Me
            If TypeOf objControl Is TextBox Then
                If objControl.Index = 0 Then colValue.Add objControl.Text, objControl.Name
            ElseIf TypeOf objControl Is CheckBox Then
                If objControl.Index = 0 Then colValue.Add objControl.Value, objControl.Name
            End If
        Next
        For Each objControl In Me
            If TypeOf objControl Is TextBox Then
                If objControl.Index = 1 Then objControl.Text = 0
            ElseIf TypeOf objControl Is CheckBox Then
                If objControl.Index = 1 Then objControl.Value = colValue(objControl.Name)
            End If
        Next
        Set colValue = Nothing
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub AddNewForm()
    On Error GoTo ChkErr
        If chkOpen(0) Then Call OpenSTB
        If chkClose(0) Then Call CloseSTB
        If chkStop(0) Then Call StopSTB
        If chkReOpen(0) Then Call ReOpenSTB
        If chkDelete(0) Then Call DeleteSTB
        If chkOpenCh(0) Then Call OpenCh
        If chkCloseCh(0) Then Call CloseCh
        If chkStopCh(0) Then Call StopCh
        If chkReopenCh(0) Then Call ReOpenCh
        If chkExtendRange(0) Then Call ExtendRange
        If chkChPrivCopy(0) Then Call ChPrivCopy
        If chkClearPinCode(0) Then Call ClearPinCode
        If chkSendMsg(0) Then Call SendMsg
    Exit Sub
ChkErr:
    ErrSub Me.Name, "AddNewForm"
End Sub

Private Sub OpenSTB()
    Dim intLoop As Integer
        '開機
        intLoop = txtOpenCount(0)
        ReDim Preserve objForm11(intLoop) As New frmPressureChild
        For intLoop = 1 To UBound(objForm11)
            With objForm11(intLoop)
                Set .uControl = txtOpenCount(1)
                .uProcessItem = 11
                .Show , Me
                Unload objForm11(intLoop)
            End With
        Next
End Sub

Private Sub CloseSTB()
    Dim intLoop As Integer
        '關機
        intLoop = txtCloseCount(0)
        ReDim Preserve objForm12(intLoop) As New frmPressureChild
        For intLoop = 1 To UBound(objForm12)
            With objForm12(intLoop)
                Set .uControl = txtCloseCount(1)
                .uProcessItem = 12
                .Show , Me
                Unload objForm12(intLoop)
            End With
        Next
End Sub

Private Sub StopSTB()
    Dim intLoop As Integer
'        '停機
        intLoop = txtStopCount(0)
        ReDim Preserve objForm13(intLoop) As New frmPressureChild
        For intLoop = 1 To UBound(objForm13)
            With objForm13(intLoop)
                Set .uControl = txtStopCount(1)
                .uProcessItem = 13
                .Show , Me
                Unload objForm13(intLoop)
            End With
        Next
End Sub

Private Sub ReOpenSTB()
    Dim intLoop As Integer
'        '復機
        intLoop = txtReopenCount(0)
        ReDim Preserve objForm14(intLoop) As New frmPressureChild
        For intLoop = 1 To UBound(objForm14)
            With objForm14(intLoop)
                Set .uControl = txtReopenCount(1)
                .uProcessItem = 14
                .Show , Me
                Unload objForm14(intLoop)
            End With
        Next
End Sub

Private Sub DeleteSTB()
    Dim intLoop As Integer
'        '註銷
        intLoop = txtDeleteCount(0)
        ReDim Preserve objForm15(intLoop) As New frmPressureChild
        For intLoop = 1 To UBound(objForm15)
            With objForm15(intLoop)
                Set .uControl = txtDeleteCount(1)
                .uProcessItem = 15
                .Show , Me
                Unload objForm15(intLoop)
            End With
        Next
End Sub

Private Sub OpenCh()
    Dim intLoop As Integer
        '開頻道
        intLoop = txtOpenChCount(0)
        ReDim Preserve objForm21(intLoop) As New frmPressureChild
        For intLoop = 1 To UBound(objForm21)
            With objForm21(intLoop)
                Set .uControl = txtOpenChCount(1)
                .uProcessItem = 21
                .Show , Me
                Unload objForm21(intLoop)
            End With
        Next
End Sub

Private Sub CloseCh()
    Dim intLoop As Integer
        '關頻道
        intLoop = txtCloseChCount(0)
        ReDim Preserve objForm22(intLoop) As New frmPressureChild
        For intLoop = 1 To UBound(objForm22)
            With objForm22(intLoop)
                Set .uControl = txtCloseChCount(1)
                .uProcessItem = 22
                .Show , Me
                Unload objForm22(intLoop)
            End With
        Next
End Sub

Private Sub StopCh()
    Dim intLoop As Integer
        '暫停頻道
        intLoop = txtStopChCount(0)
        ReDim Preserve objForm23(intLoop) As New frmPressureChild
        For intLoop = 1 To UBound(objForm23)
            With objForm23(intLoop)
                Set .uControl = txtStopChCount(1)
                .uProcessItem = 23
                .Show , Me
                Unload objForm23(intLoop)
            End With
        Next
End Sub

Private Sub ReOpenCh()
    Dim intLoop As Integer
        '恢復頻道
        intLoop = txtReOpenChCount(0)
        ReDim Preserve objForm24(intLoop) As New frmPressureChild
        For intLoop = 1 To UBound(objForm24)
            With objForm24(intLoop)
                Set .uControl = txtReOpenChCount(1)
                .uProcessItem = 24
                .Show , Me
                Unload objForm24(intLoop)
            End With
        Next
End Sub

Private Sub ExtendRange()
    Dim intLoop As Integer
        '期限延展
        intLoop = txtExtendRangeCount(0)
        ReDim Preserve objForm25(intLoop) As New frmPressureChild
        For intLoop = 1 To UBound(objForm25)
            With objForm25(intLoop)
                Set .uControl = txtExtendRangeCount(1)
                .uProcessItem = 25
                .Show , Me
                Unload objForm25(intLoop)
            End With
        Next
End Sub

Private Sub ChPrivCopy()
    Dim intLoop As Integer
        '收視權限複製
        intLoop = txtChPrivCopyCount(0)
        ReDim Preserve objForm26(intLoop) As New frmPressureChild
        For intLoop = 1 To UBound(objForm26)
            With objForm26(intLoop)
                Set .uControl = txtChPrivCopyCount(1)
                .uProcessItem = 26
                .Show , Me
                Unload objForm26(intLoop)
            End With
        Next
End Sub

Private Sub ClearPinCode()
    Dim intLoop As Integer
        '清除pincode
        intLoop = txtClearPinCode(0)
        ReDim Preserve objForm31(intLoop) As New frmPressureChild
        For intLoop = 1 To UBound(objForm31)
            With objForm31(intLoop)
                Set .uControl = txtClearPinCode(1)
                .uProcessItem = 31
                .Show , Me
                Unload objForm31(intLoop)
            End With
        Next
End Sub

Private Sub SendMsg()
    Dim intLoop As Integer
        '清除pincode
        intLoop = txtSendMsg(0)
        ReDim Preserve objForm32(intLoop) As New frmPressureChild
        For intLoop = 1 To UBound(objForm32)
            With objForm32(intLoop)
                Set .uControl = txtSendMsg(1)
                .uProcessItem = 32
                .Show , Me
                Unload objForm32(intLoop)
            End With
        Next
End Sub
