VERSION 5.00
Begin VB.UserControl IP_TextBox 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1980
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  '實心
   LockControls    =   -1  'True
   ScaleHeight     =   285
   ScaleWidth      =   1980
   Begin VB.TextBox Sepa1 
      Alignment       =   2  '置中對齊
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   900
      TabIndex        =   6
      Text            =   "."
      Top             =   52
      Width           =   120
   End
   Begin VB.TextBox Sepa1 
      Alignment       =   2  '置中對齊
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   1410
      TabIndex        =   5
      Text            =   "."
      Top             =   52
      Width           =   120
   End
   Begin VB.TextBox Sepa1 
      Alignment       =   2  '置中對齊
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   390
      TabIndex        =   4
      Text            =   "."
      Top             =   52
      Width           =   120
   End
   Begin VB.TextBox txtObj 
      Alignment       =   2  '置中對齊
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   30
      Width           =   405
   End
   Begin VB.TextBox txtObj 
      Alignment       =   2  '置中對齊
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   30
      Width           =   405
   End
   Begin VB.TextBox txtObj 
      Alignment       =   2  '置中對齊
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   510
      MaxLength       =   3
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   30
      Width           =   405
   End
   Begin VB.TextBox txtObj 
      Alignment       =   2  '置中對齊
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   0
      MaxLength       =   3
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   30
      Width           =   405
   End
End
Attribute VB_Name = "IP_TextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PR : Hammer
' Start date : 2004/12/24
' Last Modify : 2004/12/24

Option Explicit

Private Sub txtObj_Change(Index As Integer)
  On Error GoTo ChkErr
    If Index = 1 Then
        If Val(txtObj(Index)) > 223 Then
            MsgBox txtObj(Index) & " 輸入值不正確 ! 正確值應該於 1 ~ 223 之間 ! 請確認 ..", vbOKOnly + vbExclamation, "錯誤"
            txtObj(Index) = 223
            txtObj(Index).SetFocus
        End If
    Else
        If Val(txtObj(Index)) > 255 Then
            MsgBox txtObj(Index) & " 輸入值不正確 ! 正確值應該於 1 ~ 255 之間 ! 請確認 ..", vbOKOnly + vbExclamation, "錯誤"
            txtObj(Index) = 255
            txtObj(Index).SetFocus
        End If
    End If
  Exit Sub
ChkErr:
    ErrHandle "IP_TextBox", "Event : txtObj_Change"
End Sub

Private Sub txtObj_GotFocus(Index As Integer)
  On Error Resume Next
    With txtObj(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtObj_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = 8 Then
        If txtObj(Index).SelStart = 0 Then
            If Index > 1 Then
                txtObj(Index - 1).SetFocus
                txtObj(Index - 1).SelStart = Len(txtObj(Index - 1))
                txtObj(Index - 1).SelLength = 0
            End If
        End If
    End If
    If KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
        If txtObj(Index).SelStart = Len(txtObj(Index)) Then
            If Index < 4 Then
                txtObj(Index + 1).SetFocus
                txtObj(Index + 1).SelStart = 0
                txtObj(Index + 1).SelLength = 0
            End If
        End If
    End If
    If KeyCode = vbKeyReturn Then
        If Index < 4 Then
            On Error Resume Next
            txtObj(Index + 1).SetFocus
        End If
    End If
  Exit Sub
ChkErr:
    ErrHandle "IP_TextBox", "Event : txtObj_KeyDown"
End Sub

Private Sub txtObj_KeyPress(Index As Integer, KeyAscii As Integer)
  On Error GoTo ChkErr
    If KeyAscii = 46 Then
        KeyAscii = 0
        If Index < 4 Then
            txtObj(Index + 1).SetFocus
            txtObj(Index + 1).SelStart = 0
            txtObj(Index + 1).SelLength = Len(txtObj(Index + 1))
        End If
    End If
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
    If Index < 4 And Len(txtObj(Index)) >= 2 And txtObj(Index).SelText <> txtObj(Index) Then
        If txtObj(Index).SelStart = Len(txtObj(Index)) Then
            txtObj(Index + 1).SetFocus
            txtObj(Index + 1).SelStart = 0
            txtObj(Index + 1).SelLength = Len(txtObj(Index + 1))
        End If
    End If
  Exit Sub
ChkErr:
    ErrHandle "IP_TextBox", "Event : txtObj_KeyPress"
End Sub

Private Sub txtObj_LostFocus(Index As Integer)
  On Error GoTo ChkErr
    txtObj(Index) = Val(txtObj(Index))
  Exit Sub
ChkErr:
    ErrHandle "IP_TextBox", "Event : txtObj_LostFocus"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  On Error GoTo ChkErr
    With PropBag
        BackStyle = .ReadProperty("BackStyle", 1)
        BorderStyle = .ReadProperty("BorderStyle", 1)
        Appearance = .ReadProperty("Appearance", 1)
        Enabled = .ReadProperty("Enabled", True)
        Appearance = .ReadProperty("Appearance", 1)
        BackStyle = .ReadProperty("BackStyle", 1)
        BorderStyle = .ReadProperty("BorderStyle", 0)
        txtObj(1).Text = .ReadProperty("Text", "0")
    End With
  Exit Sub
ChkErr:
    ErrHandle "IP_TextBox", "Event : UserControl_ReadProperties"
End Sub

Private Sub UserControl_Show()
  On Error GoTo ChkErr
    UserControl.BackColor = vbWhite
  Exit Sub
ChkErr:
    ErrHandle "IP_TextBox", "Event : UserControl_Show"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  On Error GoTo ChkErr
    With PropBag
        Call .WriteProperty("BackStyle", BackStyle, 1)
        Call .WriteProperty("BorderStyle", BorderStyle, 1)
        Call .WriteProperty("Appearance", Appearance, 1)
        Call .WriteProperty("Enabled", Enabled, True)
        Call .WriteProperty("Appearance", Appearance, 1)
        Call .WriteProperty("BackStyle", BackStyle, 1)
        Call .WriteProperty("BorderStyle", BorderStyle, 0)
        Call .WriteProperty("Text", txtObj(1).Text, "0")
    End With
  Exit Sub
ChkErr:
    ErrHandle "IP_TextBox", "Event : UserControl_WriteProperties"
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "傳回或設定一值，決定物件是否能回應使用者產生的事件。"
  On Error GoTo ChkErr
    Enabled = UserControl.Enabled
  Exit Property
ChkErr:
    ErrHandle "IP_TextBox", "Property : Get Enabled"
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  On Error GoTo ChkErr
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
  Exit Property
ChkErr:
    ErrHandle "IP_TextBox", "Property : Let Enabled"
End Property

Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "傳回或設定物件在執行階段時，是否以立體效果繪製。"
  On Error GoTo ChkErr
    Appearance = UserControl.Appearance
  Exit Property
ChkErr:
    ErrHandle "IP_TextBox", "Property : Get Appearance"
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
  On Error GoTo ChkErr
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
  Exit Property
ChkErr:
    ErrHandle "IP_TextBox", "Property : Let Appearance"
End Property

Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "用來決定 Label 本身或 Shape 的背景是否為透明的。"
  On Error GoTo ChkErr
    BackStyle = UserControl.BackStyle
  Exit Property
ChkErr:
    ErrHandle "IP_TextBox", "Property : Get BackStyle"
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
  On Error GoTo ChkErr
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
  Exit Property
ChkErr:
    ErrHandle "IP_TextBox", "Property : Let BackStyle"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "傳回或設定物件的框線樣式。"
  On Error GoTo ChkErr
    BorderStyle = UserControl.BorderStyle
  Exit Property
ChkErr:
    ErrHandle "IP_TextBox", "Property : Get BackStyle"
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
  On Error GoTo ChkErr
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
  Exit Property
ChkErr:
    ErrHandle "IP_TextBox", "Property : Let BackStyle"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "強制物件進行完整的重繪。"
  On Error GoTo ChkErr
    UserControl.Refresh
  Exit Sub
ChkErr:
    ErrHandle "IP_TextBox", "Subroutine : Refresh"
End Sub

Public Property Get Text() As String
Attribute Text.VB_Description = "傳回或設定控制項所包含的文字。"
Attribute Text.VB_ProcData.VB_Invoke_Property = ";文字"
Attribute Text.VB_UserMemId = -517
Attribute Text.VB_MemberFlags = "200"
  On Error GoTo ChkErr
    Text = txtObj(1) & "." & txtObj(2) & "." & txtObj(3) & "." & txtObj(4)
  Exit Property
ChkErr:
    ErrHandle "IP_TextBox", "Property : Get Text"
End Property

Public Property Let Text(ByVal New_Text As String)
  On Error GoTo ChkErr
    Dim v As Variant
    v = Split(New_Text, ".")
    txtObj(1).Text = v(0)
    txtObj(2).Text = v(1)
    txtObj(3).Text = v(2)
    txtObj(4).Text = v(3)
    v = 0
    PropertyChanged "Text"
  Exit Property
ChkErr:
    ErrHandle "IP_TextBox", "Property : Let Text"
End Property

