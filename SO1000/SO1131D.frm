VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1131D 
   BorderStyle     =   1  '單線固定
   Caption         =   "帳戶管理 [SO1131D]"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   9315
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO1131D.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   3045
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   5371
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
End
Attribute VB_Name = "frmSO1131D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngCustId As Long
Dim rsData As New ADODB.Recordset
Dim objParentForm As Form

Private blnSaveOK As Boolean
Private intBankCode As Integer
Private strBankName As String
Private intCMCode As Integer
Private strCMName As String
Private intPTCode As Integer
Private strPTName As String
Private strAccountNo As String
Private strMasterId As String

Public Property Let uCustId(ByVal vData As String)
    lngCustId = vData
    
End Property

Public Property Let uParentForm(ByVal vData As Form)
    Set objParentForm = vData
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        Select Case KeyCode
            Case vbKeyEscape
                Unload Me
        End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
        Call InitData
        Call subGrd
        Call OpenData
End Sub

Private Sub InitData()
    On Error Resume Next
        Me.Move objParentForm.Left + (objParentForm.Width - Me.Width) / 2, objParentForm.Top + (objParentForm.Height - Me.Height) / 2
End Sub

Private Sub OpenData()
    On Error Resume Next
        Dim strQry As String
        '#6316 將SO106 改為SO002A By Kin 2012/09/13
        If Not GetRS(rsData, "Select * From " & GetOwner & "SO002A Where CustId = " & lngCustId & " And StopFlag =0") Then Exit Sub
        Set ggrData.Recordset = rsData
        ggrData.Refresh
End Sub

Private Sub subGrd()
  '#3436 因為改為多筆所以將收件人、收費方式、收費地址、郵寄地址拿掉 By Kin 2007/11/28
  '#6316 測試不OK,有些欄位不存在 By Kin 2012/10/02
  Dim mFlds As New GiGridFlds
      'mFlds.Add "SnactionDate", , , , , "狀態" & Space(10)
      mFlds.Add "BankName", , , , , "銀行名稱  " & Space(10)
      
      mFlds.Add "AccountNo", , , , , "帳號/卡號           "
      mFlds.Add "CardName", , , , , "信用卡別"
      'mFlds.Add "StopYM", , , "00/0000", , "有效期限"
      mFlds.Add "CardExpDate", , , "00/0000", , "有效期限"
      'mFlds.Add "PTName", , , , , "付款種類"
      'mFlds.Add "CMname", , , , , "收費方式" & Space(10)
      'mFlds.Add "MasterId", , , , , "帳戶序號"
      
      'mFlds.Add "ChargeAddress", , , , , "收費地址" & Space(28)
      'mFlds.Add "MailAddress", , , , , "郵寄地址" & Space(28)
      ggrData.AllFields = mFlds

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        ReleaseCOM frmSO1131D

End Sub

Private Sub ggrData_DblClick()
    On Error Resume Next
        If Not objParentForm Is Nothing Then
            With objParentForm
                .gilBankCode.SetCodeNo rsData("BankCode") & ""
                .gilBankCode.Query_Description
                .gilCMCode.SetCodeNo rsData("CMCode") & ""
                .gilCMCode.Query_Description
                .gilPTCode.SetCodeNo rsData("PTCode") & ""
                .gilPTCode.Query_Description
                .cboAccountNo.Text = rsData("AccountId") & ""
            End With
        End If
        intBankCode = Val(rsData("BankCode") & "")
        strBankName = rsData("BankName") & ""
        'intCMCode = Val(rsData("CMCode") & "")
        'strCMName = rsData("CMName") & ""
        'intPTCode = Val(rsData("PTCode") & "")
        'strPTName = rsData("PTName") & ""
        'strMasterId = rsData("MasterId")
        strAccountNo = rsData("AccountNo") & ""
        blnSaveOK = True
        Unload Me
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error Resume Next
        If Not GetUserPriv("SO1100G", "(SO1100G9)") Then
            If UCase(giFld.FieldName) = "ACCOUNTID" Then
                If Value = Empty Then
                    Value = ""
                Else
                    Value = Left(Value, 5) & "XXX" & Mid(Value, 9, 2) & "XX" & Right(Value, 4)
                End If
            End If
        End If
        If UCase(giFld.FieldName) = "SNACTIONDATE" Then
            If Value & "" = "" Then
                Value = "未核準"
            Else
                Value = "核準"
            End If
        End If
End Sub
Public Property Get uSaveOk() As Boolean
  On Error GoTo ChkErr
    uSaveOk = blnSaveOK
    Exit Property
ChkErr:
  Call ErrSub(Me.Name, "uSaveOk")
End Property
Public Property Get uBankCode() As Integer
  On Error GoTo ChkErr
    uBankCode = intBankCode
    Exit Property
ChkErr:
  Call ErrSub(Me.Name, "uBankCode")
End Property
Public Property Get uBankName() As String
  On Error GoTo ChkErr
    uBankName = strBankName
    Exit Property
ChkErr:
  Call ErrSub(Me.Name, "uBankName")
End Property
Public Property Get uCMCode() As Integer
  On Error GoTo ChkErr
    uCMCode = intCMCode
    Exit Property
ChkErr:
  Call ErrSub(Me.Name, "uCMCode")
End Property
Public Property Get uCMName() As String
  On Error GoTo ChkErr
    uCMName = strCMName
    Exit Property
ChkErr:
  Call ErrSub(Me.Name, "uCMName")
End Property
Public Property Get uPTCode() As Integer
  On Error GoTo ChkErr
    uPTCode = intPTCode
    Exit Property
ChkErr:
  Call ErrSub(Me.Name, "uPTCode")
End Property
Public Property Get uPTName() As String
  On Error GoTo ChkErr
    uPTName = strPTName
    Exit Property
ChkErr:
  Call ErrSub(Me.Name, "uPTName")
End Property
Public Property Get uAccountNo() As String
  On Error GoTo ChkErr
    uAccountNo = strAccountNo
    Exit Property
ChkErr:
  Call ErrSub(Me.Name, "uAccountNo")
End Property
Public Property Get uMasterId() As String
  On Error GoTo ChkErr
    uMasterId = strMasterId
    Exit Property
ChkErr:
  Call ErrSub(Me.Name, "uMasterId")
End Property

