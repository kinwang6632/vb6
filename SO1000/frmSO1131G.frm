VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1131G 
   BorderStyle     =   1  '單線固定
   Caption         =   "發票抬頭選擇[SO1131G]"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   Icon            =   "frmSO1131G.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   150
      TabIndex        =   3
      Top             =   3660
      Width           =   1005
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.選定"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2730
      TabIndex        =   2
      Top             =   3690
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Height          =   3555
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   10365
      Begin prjGiGridR.GiGridR ggrData 
         Height          =   3165
         Left            =   150
         TabIndex        =   1
         Top             =   180
         Width           =   10065
         _ExtentX        =   17754
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
   End
End
Attribute VB_Name = "frmSO1131G"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#3436 Kin Add 2007/11/27
Option Explicit
Dim lngCustId As Long
Dim rsData As New ADODB.Recordset
Attribute rsData.VB_VarHelpID = -1
Dim objParentForm As Form
Dim intcompcode As String
Dim strAccountNo As String
Dim strInvSeqNo As String
Dim strOldInvSeqNo As String
Dim rsDefTmp As New ADODB.Recordset
Private lngBm As Long
Public Property Let uInvSeqNo(ByVal vData As String)
    strInvSeqNo = vData
    strOldInvSeqNo = strInvSeqNo
End Property
Public Property Get uInvSeqNo() As String
    uInvSeqNo = strInvSeqNo
End Property
Public Property Let uCustId(ByVal vData As String)
    lngCustId = vData
End Property
Public Property Let uCompCode(ByVal vData As String)
    intcompcode = vData
End Property
Public Property Let uAccountNo(ByVal vData As String)
    strAccountNo = vData
End Property
Public Property Let uParentForm(ByVal vData As Form)
    Set objParentForm = vData
End Property

Private Sub cmdCancel_Click()
  On Error Resume Next
    strInvSeqNo = strOldInvSeqNo
    Unload Me
End Sub

Private Sub cmdOK_Click()
  On Error Resume Next
    If rsDefTmp.RecordCount = 0 Then strInvSeqNo = ""
    Unload Me
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        Select Case KeyCode
            Case vbKeyF2
                If cmdOK.Enabled Then cmdOK.Value = True
            Case vbKeyEscape
                If cmdCancel.Enabled Then cmdCancel.Value = True
        End Select

End Sub

Private Sub Form_Load()
    On Error Resume Next
        Call InitData
        Call subGrd
        Call OpenData
End Sub
Private Sub OpenData()
  On Error GoTo ChkErr
    Dim strQry As String
    Dim str02AD As String
    Dim strField As String
    
    Dim i As Integer
    'DefineRs
    cmdOK.Enabled = False
    str02AD = "Select B.InvSeqNo From " & GetOwner & "SO002AD B " & _
              " Where B.CustId=" & lngCustId & _
              " And B.AccountNo='" & strAccountNo & "' " & _
              " And B.CompCode=" & intcompcode & _
              " And A.InvSeqNo=B.InvSeqNo"
    strQry = "Select A.* From " & GetOwner & "SO138 A Where Exists(" & str02AD & ") And Stopflag=0"
'    strQry = "Select A.* From " & GetOwner & "SO138 A Where CustId = " & lngCustId & " And Stopflag=0"
    If Not GetRS(rsData, strQry, gcnGi, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Sub
    If strInvSeqNo <> Empty Then
        If rsData.RecordCount > 0 Then
            rsData.Find "InvSeqNo=" & strInvSeqNo
            If Not rsData.EOF Then
                lngBm = rsData.AbsolutePosition
            Else
                lngBm = 0
            End If
        End If
    End If
'    Do While Not rsData.EOF
'        rsDefTmp.AddNew
'        For i = 0 To rsDefTmp.Fields.Count - 1
'            If rsDefTmp.Fields(i).Name = "Choice" Then
'                rsDefTmp.Fields(i) = "0"
'            ElseIf rsDefTmp.Fields(i).Name = "InvSeqNo" Then
'                If strInvSeqNo <> "" Then
'                    If rsData("invSeqNo") = strInvSeqNo Then
'                        rsDefTmp("Choice").Value = "1"
'                        cmdOK.Enabled = True
'                    End If
'                    rsDefTmp.Fields(i).Value = rsData.Fields(rsDefTmp.Fields(i).Name).Value
'                Else
'                    rsDefTmp.Fields(i).Value = rsData.Fields(rsDefTmp.Fields(i).Name).Value
'                End If
'            Else
'                rsDefTmp.Fields(i).Value = rsData.Fields(rsDefTmp.Fields(i).Name).Value
'            End If
'        Next i
'        rsDefTmp.Update
'        rsData.MoveNext
'        DoEvents
'    Loop
'    If rsDefTmp.RecordCount = 0 Then cmdOK.Enabled = False
'    Set ggrData.Recordset = rsDefTmp
    'ggrData.Visible = True
    Set ggrData.Recordset = rsData
    ggrData.Refresh
    If lngBm > 0 Then
        'ggrData.Recordset.AbsolutePosition = lngBm
        If rsData.RecordCount > 0 Then
            rsData.AbsolutePosition = lngBm
        End If
    End If
    'ggrData.Blank = True
    
    
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
        '.Add "Choice", , , , , "選取"
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
    'ggrData.Visible = False
End Sub

'Private Sub ggrData_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
'    On Error Resume Next
'    If UCase(giFld.FieldName) = "CHOICE" Then
'        If Value = 1 Then Value = vbRed
'    End If
'End Sub

Private Sub ggrData_DblClick()
  On Error Resume Next
    If rsData.RecordCount = 0 Then
        strInvSeqNo = ""
    Else
        strInvSeqNo = ggrData.Recordset("InvSeqNo") & ""
    End If
    Unload Me
'    Dim varBookMark As Variant
'    LockWindowUpdate Me.hwnd
'    If ggrData.Recordset.RecordCount > 0 Then
'        If ggrData.Recordset("Choice") = "1" Then
'            ggrData.Recordset("Choice") = "0"
'            ggrData.Recordset.Update
'            strInvSeqNo = Empty
'        Else
'            varBookMark = rsDefTmp.Bookmark
'             rsDefTmp.MoveFirst
'            Do While Not rsDefTmp.EOF
'                 rsDefTmp("Choice") = "0"
'                 rsDefTmp.Update
'                 rsDefTmp.MoveNext
'            Loop
'             rsDefTmp.AbsolutePosition = varBookMark
'             rsDefTmp("Choice") = "1"
'             rsDefTmp.Update
'             strInvSeqNo = rsDefTmp("InvSeqNo") & ""
'        End If
'    End If
'    LockWindowUpdate 0
'    If strInvSeqNo = Empty Then
'        cmdOK.Enabled = False
'    Else
'        cmdOK.Enabled = True
'    End If
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
'    If giFld.FieldName = "Choice" Then
'        Select Case Value
'            Case 1
'                Value = "V"
'            Case Else
'                Value = ""
'        End Select
'    End If
End Sub
Private Sub DefineRs()
  On Error GoTo ChkErr
    With rsDefTmp
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        '.Fields.Append ("Choice"), adBSTR, 1, adFldIsNullable
        .Fields.Append ("InvSeqNo"), adBSTR, 20, adFldIsNullable
        .Fields.Append ("ChargeTitle"), adBSTR, 50, adFldIsNullable
        .Fields.Append ("InvoiceType"), adBSTR, 1, adFldIsNullable
        .Fields.Append ("InvNo"), adBSTR, 8, adFldIsNullable
        .Fields.Append ("InvTitle"), adBSTR, 50, adFldIsNullable
        .Fields.Append ("InvAddress"), adBSTR, 90, adFldIsNullable
        .Fields.Append ("ChargeAddress"), adBSTR, 90, adFldIsNullable
        .Open
    End With
    Exit Sub
ChkErr:
  ErrSub Me.Name, "DefineRs"
End Sub

