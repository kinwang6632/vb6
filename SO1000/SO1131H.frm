VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1131H 
   BorderStyle     =   1  '單線固定
   Caption         =   "發票抬頭選擇[SO1131H]"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   Icon            =   "SO1131H.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame fraQuery 
      Height          =   4995
      Left            =   90
      TabIndex        =   6
      Top             =   60
      Width           =   10545
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&X)"
         Height          =   435
         Left            =   9180
         TabIndex        =   10
         Top             =   4380
         Width           =   1005
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "F2.選定"
         Height          =   435
         Left            =   1350
         TabIndex        =   5
         Top             =   4380
         Width           =   1005
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "F3.尋找"
         Height          =   435
         Left            =   210
         TabIndex        =   4
         Top             =   4380
         Width           =   1005
      End
      Begin VB.TextBox txtInvTitle 
         Height          =   315
         Left            =   1170
         MaxLength       =   50
         TabIndex        =   2
         Top             =   660
         Width           =   4785
      End
      Begin VB.TextBox txtInvNo 
         Height          =   315
         Left            =   7050
         MaxLength       =   8
         TabIndex        =   1
         Top             =   210
         Width           =   1545
      End
      Begin VB.TextBox txtChargeTitle 
         Height          =   315
         Left            =   1170
         MaxLength       =   50
         TabIndex        =   0
         Top             =   210
         Width           =   4785
      End
      Begin prjGiGridR.GiGridR ggrData 
         Height          =   3165
         Left            =   150
         TabIndex        =   3
         Top             =   1110
         Width           =   10245
         _ExtentX        =   18071
         _ExtentY        =   5583
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseCellForeColor=   -1  'True
      End
      Begin VB.Label lblInvTitle 
         Caption         =   "發票抬頭"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   300
         TabIndex        =   9
         Top             =   750
         Width           =   795
      End
      Begin VB.Label lblInvNo 
         Caption         =   "統一編號"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   6180
         TabIndex        =   8
         Top             =   270
         Width           =   945
      End
      Begin VB.Label lblChargeTitle 
         Caption         =   "收 件 人"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   390
         TabIndex        =   7
         Top             =   270
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmSO1131H"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#3436 Kin Add 2007/11/27
Option Explicit
Dim lngCustId As Long
Dim rsData As New ADODB.Recordset
Dim objParentForm As Form
Dim intcompcode As String
Dim strAccountNo As String
Dim strInvSeqNo As String
Dim strOldInvSeqNo As String
Private rsDefTmp As New ADODB.Recordset
Private rsReturn As New ADODB.Recordset
Private blnMutilChoice As Boolean
Private strNewInvSeqNo As String
Public Property Get uRS() As ADODB.Recordset
    Set uRS = rsReturn
End Property
Public Property Let uParentForm(ByVal vData As Form)
    Set objParentForm = vData
End Property
Public Property Let uMutilChoice(ByVal vData As Boolean)
    blnMutilChoice = vData
End Property
Public Property Get uNewInvSeqNo() As String
    uNewInvSeqNo = strNewInvSeqNo
End Property
Private Sub cmdCancel_Click()
  On Error Resume Next
    strInvSeqNo = strOldInvSeqNo
    Unload Me
End Sub

Private Sub cmdOK_Click()
  On Error Resume Next
    Set rsReturn = rsDefTmp.Clone
    Unload Me
End Sub

Private Sub cmdQuery_Click()
  On Error GoTo ChkErr
    Dim strSQL As String
    Dim strWhere As String
    Dim i As Integer
    DoEvents
    Screen.MousePointer = vbHourglass
    cmdOK.Enabled = False
    DefineRs
    strWhere = subChoose
    If strWhere = "" Then
        MsgBox "無任何搜尋條件", vbInformation, "訊息"
        Screen.MousePointer = vbDefault
        Set ggrData.Recordset = rsDefTmp
        ggrData.Refresh
        Exit Sub
    End If
    strSQL = "Select * From " & GetOwner & "SO138 A " & IIf(strWhere = "", " Where A.StopFlag=0", " Where " & strWhere & " And A.StopFlag=0") & " Order By A.InvSeqNo"
    If Not GetRS(rsData, strSQL, gcnGi) Then Exit Sub
    If rsData.RecordCount <= 0 Then
        MsgBox "查無資料! 請確認條件", vbInformation, "訊息"
        Screen.MousePointer = vbDefault
        Set ggrData.Recordset = rsDefTmp
        ggrData.Refresh
        Exit Sub
    End If
    If blnMutilChoice Then
        Do While Not rsData.EOF
            rsDefTmp.AddNew
            For i = 0 To rsDefTmp.Fields.Count - 1
                If rsDefTmp.Fields(i).Name = "Choice" Then
                    rsDefTmp.Fields(i) = "0"
                ElseIf rsDefTmp.Fields(i).Name = "InvSeqNo" Then
                    If strInvSeqNo <> "" Then
                        If rsData("invSeqNo") = strInvSeqNo Then rsDefTmp("Choice").Value = "1"
                        rsDefTmp.Fields(i).Value = rsData.Fields(rsDefTmp.Fields(i).Name).Value
                    Else
                        rsDefTmp.Fields(i).Value = rsData.Fields(rsDefTmp.Fields(i).Name).Value
                    End If
                Else
                    rsDefTmp.Fields(i).Value = rsData.Fields(rsDefTmp.Fields(i).Name).Value
                End If
            Next i
            rsDefTmp.Update
            rsData.MoveNext
        Loop
    Else
        Set ggrData.Recordset = rsData
        ggrData.Refresh
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'cmdOK.Enabled = True
    Set ggrData.Recordset = rsDefTmp
    ggrData.Refresh
    Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
    Screen.MousePointer = vbDefault
    ErrSub Me.Name, "cmdQuery"
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
    txtChargeTitle.SetFocus
    cmdOK.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        Select Case KeyCode
            Case vbKeyF2
                If cmdOK.Enabled Then cmdOK.Value = True
            Case vbKeyEscape
                Unload Me
            Case vbKeyF3
                If cmdQuery.Enabled Then cmdQuery.Value = True
        End Select

End Sub

Private Sub Form_Load()
    On Error Resume Next
        If blnMutilChoice Then
            cmdOK.Visible = True
        Else
            cmdOK.Visible = False
        End If
        strNewInvSeqNo = Empty
        Call InitData
        Call subGrd
        Call OpenData
End Sub
Private Sub OpenData()
  On Error GoTo ChkErr
    Dim strQry As String
    Dim strField As String
    DefineRs
    strQry = "Select A.* From " & GetOwner & "SO138 A Where 1=0"
    If Not GetRS(rsData, strQry, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Sub
    
    Set ggrData.Recordset = rsDefTmp
    ggrData.Refresh
  Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub
Private Sub InitData()
    On Error Resume Next
        If objParentForm Is Nothing Then
            Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
        Else
            Me.Move objParentForm.Left + ((objParentForm.Width - Me.Width) / 2), objParentForm.Top + objParentForm.Height - Me.Height
        End If
End Sub
Private Sub subGrd()
    On Error Resume Next
    Dim mFlds As New GiGridFlds
    With mFlds
        If blnMutilChoice Then
            .Add "Choice", , , , , "選取"
        End If
        .Add "InvSeqNo", , , , , "流水編號     "
        .Add "ChargeTitle", , , , , "收件人" & Space(40)
        .Add "InvoiceType", , , , , "發票種類"
        .Add "InvNo", , , , , "發票統一編號"
        .Add "InvTitle", , , , , "發票抬頭" & Space(50)
        .Add "InvAddress", , , , , "發票地址" & Space(50)
        .Add "ChargeAddress", , , , , "收費地址" & Space(50)
    End With
'
    ggrData.AllFields = mFlds
    ggrData.SetHead
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    CloseRecordset rsDefTmp
End Sub

Private Sub ggrData_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error Resume Next
    If UCase(giFld.FieldName) = "CHOICE" Then
        If Value = 1 Then Value = vbRed
    End If
End Sub

Private Sub ggrData_DblClick()
  On Error Resume Next
    Dim varBookMark As Variant
    Static intClick As Integer
    LockWindowUpdate Me.hwnd
    
    '判斷是否可多選SO1100F可多選、SO1100G不可多選
    If blnMutilChoice Then
        If ggrData.Recordset.RecordCount > 0 Then
            If ggrData.Recordset("Choice") = "1" Then
                ggrData.Recordset("Choice") = "0"
                
                '如果取消選擇,則計數要減1,但不能小於0
                If intClick > 0 Then intClick = intClick - 1
            Else
                ggrData.Recordset("Choice") = "1"
                '判斷選到幾個
                intClick = intClick + 1
            End If
            ggrData.Recordset.Update
            If intClick > 0 Then
                cmdOK.Enabled = True
            Else
                cmdOK.Enabled = False
            End If
        End If
    Else
        LockWindowUpdate 0
        strNewInvSeqNo = rsData("InvSeqNo") & ""
        Unload Me
'        If ggrData.Recordset.RecordCount > 0 Then
'            If ggrData.Recordset("Choice") = "1" Then
'                ggrData.Recordset("Choice") = "0"
'                ggrData.Recordset.Update
'                cmdOK.Enabled = False
'            Else
'                varBookMark = ggrData.Recordset.Bookmark
'                 ggrData.Recordset.MoveFirst
'                Do While Not ggrData.Recordset.EOF
'                     ggrData.Recordset("Choice") = "0"
'                     ggrData.Recordset.Update
'                     ggrData.Recordset.MoveNext
'                Loop
'                 ggrData.Recordset.AbsolutePosition = varBookMark
'                 ggrData.Recordset("Choice") = "1"
'                 ggrData.Recordset.Update
'                 cmdOK.Enabled = True
'            End If
'        End If
    End If
    LockWindowUpdate 0
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next

    If giFld.FieldName = "InvoiceType" Then
        Select Case Value
            Case 0
                Value = "免開"
            Case 2
                Value = "二聯式"
            Case 3
                Value = "三聯式"
        End Select
    End If
    If giFld.FieldName = "Choice" Then
        Select Case Value
            Case 1
                Value = "V"
            Case Else
                Value = ""
        End Select
    End If
End Sub
Private Sub DefineRs()
  On Error GoTo ChkErr
    '#3828 原因程式寫死欄位,現在改由直接判斷SO138的欄位產生 By Kin 2008/03/25
    Dim rsDef138 As New ADODB.Recordset
    Dim i As Long
    If Not GetRS(rsDef138, "Select * From SO138 Where 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    
    With rsDefTmp
        If .State = adStateOpen Then
            .Close
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        If blnMutilChoice Then
            .Fields.Append ("Choice"), adBSTR, 1, adFldIsNullable
        End If
    End With
    For i = 0 To rsDef138.Fields.Count - 1
        With rsDefTmp
            .Fields.Append rsDef138.Fields(i).Name, adBSTR, rsDef138.Fields(i).DefinedSize, adFldIsNullable
        End With
    Next i
    rsDefTmp.Open
'    With rsDefTmp
'        If .State = adStateOpen Then
'            .Close
'        End If
'        .CursorLocation = adUseClient
'        .CursorType = adOpenKeyset
'        .LockType = adLockOptimistic
'        If blnMutilChoice Then
'            .Fields.Append ("Choice"), adBSTR, 1, adFldIsNullable
'        End If
'
'
'        .Fields.Append ("InvSeqNo"), adBSTR, 20, adFldIsNullable
'        .Fields.Append ("ChargeTitle"), adBSTR, 50, adFldIsNullable
'        .Fields.Append ("InvoiceType"), adBSTR, 1, adFldIsNullable
'        .Fields.Append ("InvNo"), adBSTR, 8, adFldIsNullable
'        .Fields.Append ("InvTitle"), adBSTR, 50, adFldIsNullable
'        .Fields.Append ("InvAddress"), adBSTR, 90, adFldIsNullable
'        .Fields.Append ("ChargeAddrNo"), adBSTR, 8, adFldIsNullable
'        .Fields.Append ("ChargeAddress"), adBSTR, 90, adFldIsNullable
'        .Fields.Append ("MailAddrNo"), adBSTR, 8, adFldIsNullable
'        .Fields.Append ("MailAddress"), adBSTR, 90, adFldIsNullable
'        .Fields.Append ("StopFlag"), adBSTR, 1, adFldIsNullable
'        .Fields.Append ("StopDate"), adBSTR, 20, adFldIsNullable
'        .Fields.Append ("UpdTime"), adBSTR, 19, adFldIsNullable
'        .Fields.Append ("UpdEn"), adBSTR, 20, adFldIsNullable
'        .Fields.Append ("InvPurposeCode"), adBSTR, 20, adFldIsNullable
'        .Fields.Append ("InvPurposeName"), adBSTR, 20, adFldIsNullable
'        .Open
'    End With
    On Error Resume Next
    Call CloseRecordset(rsDef138)
    Exit Sub
ChkErr:
  ErrSub Me.Name, "DefineRs"
End Sub
Private Function subChoose() As String
    On Error GoTo ChkErr
        strChoose = ""
        If txtChargeTitle.Text <> "" Then subAnd "UPPER(A.ChargeTitle) Like '" & UCase(txtChargeTitle.Text) & "%'"
        If txtInvNo.Text <> "" Then subAnd "UPPER(A.InvNo) Like '" & UCase(txtInvNo.Text) & "%'"
        If txtInvTitle.Text <> "" Then subAnd "UPPER(A.InvTitle) Like '" & UCase(txtInvTitle.Text) & "%'"
        subChoose = strChoose
    Exit Function
ChkErr:
    ErrSub Me.Name, "subChoose"
End Function


Private Sub txtInvNo_KeyPress(KeyAscii As Integer)
  On Error Resume Next
  If KeyAscii = 8 Then Exit Sub
  If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
  KeyAscii = 8
End Sub
