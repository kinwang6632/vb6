VERSION 5.00
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Begin VB.Form frmSO3150A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�X�ֱb��[SO3150]"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   Icon            =   "SO3150.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5505
   StartUpPosition =   1  '���ݵ�������
   Begin VB.CommandButton cmdSave 
      Caption         =   "F5. �T�w"
      Height          =   435
      Left            =   870
      TabIndex        =   3
      Top             =   1455
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "����(&X)"
      Height          =   435
      Left            =   3270
      TabIndex        =   2
      Top             =   1455
      Width           =   1245
   End
   Begin Gi_YM.GiYM gymCBill 
      Height          =   315
      Left            =   2310
      TabIndex        =   1
      Top             =   615
      Width           =   795
      _ExtentX        =   1402
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
   Begin VB.Label Label1 
      Caption         =   "�X�֦~��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1140
      TabIndex        =   0
      Top             =   675
      Width           =   1020
   End
End
Attribute VB_Name = "frmSO3150A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
End Sub

Private Sub cmdSave_Click()

Dim strCombineDate As String
Dim strMessage As String


On Error GoTo ChkErr

If Len(gymCBill.Text) < 1 Then
    MsgBox "�X�֦~����줣�i�H�L��  !!", vbOKOnly + vbInformation, " �T�� "
    gymCBill.SetFocus
    Exit Sub
End If
   strCombineDate = gymCBill.Text
   
   strMessage = ""

''"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'' �H�U������ �{��  '' sf_CombineBill('200208',:msg);
''"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
Dim Cmd_CombineBill As New ADODB.Command

    Cmd_CombineBill.Parameters.Append Cmd_CombineBill.CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, 10)

    Cmd_CombineBill.Parameters.Append Cmd_CombineBill.CreateParameter("p_CombineYM", adVarChar, adParamInput, 8, strCombineDate)
    Cmd_CombineBill.Parameters.Append Cmd_CombineBill.CreateParameter("p_RetMsg ", adVarChar, adParamOutput, 2000)
    Set Cmd_CombineBill.ActiveConnection = gcnGi
    Cmd_CombineBill.CommandText = "sf_CombineBill"
    Cmd_CombineBill.CommandType = adCmdStoredProc
    Cmd_CombineBill.Execute
    
    
     If Cmd_CombineBill.Parameters(0) < 1 Then
        strMessage = Cmd_CombineBill.Parameters(2)
        MsgBox strMessage, vbExclamation, "�T��"
    End If
 
    Set Cmd_CombineBill = Nothing

''"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Not ChkGiList(KeyCode) Then Exit Sub
    If KeyCode = vbKeyF5 Then cmdSave.Value = True: KeyCode = 0
    If KeyCode = vbKeyEscape Then Unload Me: KeyCode = 0
End Sub

