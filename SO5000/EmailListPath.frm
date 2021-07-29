VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEmailListPath 
   Caption         =   "EmailList 輸出路徑"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   Icon            =   "EmailListPath.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1785
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5400
      TabIndex        =   3
      Top             =   285
      Width           =   495
   End
   Begin VB.TextBox txtFilePath 
      Height          =   360
      Left            =   1275
      TabIndex        =   0
      Top             =   300
      Width           =   3870
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "F2 確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   300
      TabIndex        =   2
      Top             =   1230
      Width           =   1215
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
      Height          =   405
      Left            =   4590
      TabIndex        =   1
      Top             =   1230
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog comdPath 
      Left            =   5760
      Top             =   750
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblEmailList 
      AutoSize        =   -1  'True
      Caption         =   "輸出路徑"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   210
      TabIndex        =   4
      Top             =   360
      Width           =   960
   End
End
Attribute VB_Name = "frmEmailListPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objParentForm As Object

Private Sub cmdCancel_Click()
On Error Resume Next
    objParentForm.uFilePath = ""
    objParentForm.uCancelFlag = True
    Unload Me
End Sub

Private Sub cmdPath_Click()
On Error GoTo ChkErr
   Dim strFilePath As String
      strFilePath = FolderDialog("選擇輸出路徑")
      If strFilePath <> "" Then txtFilePath.Text = strFilePath
   Exit Sub
ChkErr:
   ErrSub Me.Name, "cmdPath_Click"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ChkErr
        If Not ChkDirExist(txtFilePath.Text & "") Then
            txtFilePath.SetFocus
            MsgBox "路徑不存在!", vbExclamation, "警告..."
            Exit Sub
        End If
        cmdSave.Tag = "Save"
        objParentForm.uFilePath = txtFilePath.Text
        objParentForm.uCancelFlag = False
        Unload Me
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdsave_Click"
End Sub

Private Sub Form_Activate()
On Error Resume Next
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
       Case vbKeyF2
          Call cmdSave_Click
       Case vbKeyEscape
          Call cmdCancel_Click
    End Select
    
    
End Sub

Private Sub Form_Load()
   On Error Resume Next

   Dim FSO As New FileSystemObject
   Dim TmpFile As TextStream
   Dim strTmpPath As String
   
   strTmpPath = ReadGICMIS1("ErrLogPath")
   If FSO.FileExists(strTmpPath & "\EmailListPath.log") Then
      Set TmpFile = FSO.OpenTextFile(strTmpPath & "\EmailListPath.log")
      txtFilePath.Text = TmpFile.ReadLine
   Else
      txtFilePath.Text = ""
   End If
End Sub

Public Property Get uParentForm() As Object
    Set uParentForm = objParentForm
End Property

Public Property Set uParentForm(ByVal vData As Object)
    Set objParentForm = vData
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
     If Len(cmdSave.Tag & "") = 0 Then objParentForm.uCancelFlag = True
End Sub
