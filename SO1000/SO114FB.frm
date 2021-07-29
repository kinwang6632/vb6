VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO114FB 
   BorderStyle     =   1  '單線固定
   Caption         =   "異動記錄[SO114FB]"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   10455
   Icon            =   "SO114FB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   10455
   StartUpPosition =   1  '所屬視窗中央
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   4125
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   7276
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSO114FB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FInvSeqNo As String
Private rsLog As New ADODB.Recordset
Public Property Let uInvSeqNo(ByVal vData As String)
  FInvSeqNo = vData
End Property

Private Sub Form_Load()
    Call subGrd
    Call OpenData
End Sub
Private Sub subGrd()
    On Error Resume Next
    Dim mFlds As New GiGridFlds
    Dim mFlds2 As New GiGridFlds
    With mFlds
        .Add "FuncType", , , , , "異動種類"
        .Add "UpdTime", , , , , "異動時間" & Space(10)
        .Add "UpdEn", , , , , "異動人員" & Space(10)
        .Add "InvSeqNo", , , , , "流水編號     "
        .Add "ChargeTitle", , , , , "收件人" & Space(40)
        .Add "InvoiceType", , , , , "發票種類"
        .Add "InvNo", , , , , "發票統一編號"
        .Add "InvTitle", , , , , "發票抬頭" & Space(50)
        .Add "InvAddress", , , , , "發票地址" & Space(50)
        .Add "ChargeAddress", , , , , "收費地址" & Space(50)
        .Add "MailAddress", , , , , "郵寄地址" & Space(50)
        .Add "StopFlag", , , , , "停用"
        .Add "StopDate", giControlTypeDate, , , , "停用日期" & Space(5)
        .Add "InvPurposeName", , , , , "發票用途名稱" & Space(10)
        .Add "PreInvoice", , , , , "是否預開發票"
        .Add "InvoiceKind", , , , , "發票開立種類"
        .Add "BillMailKind", , , , , "帳單寄送方式"
        .Add "DenRecName", , , , , "受贈單位名稱"
        .Add "DenRecDate", giControlTypeDate, , , , "受理贈送日期" & Space(10)
        .Add "ApplyInvDate", giControlTypeDate, , , , "申請電子計算機發票日期"
        
    End With
    ggrData.AllFields = mFlds
    ggrData.SetHead
End Sub
Private Sub OpenData()
  On Error GoTo ChkErr
    Dim aSQL As String
    aSQL = "SELECT * FROM " & GetOwner & "SO138LOG WHERE InvSeqNo=" & FInvSeqNo & _
        " ORDER BY UpdTime DESC"
    If Not GetRS(rsLog, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Set ggrData.Recordset = rsLog
    ggrData.Refresh
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "OpenData")
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
    If giFld.FieldName = "StopFlag" Then
        Select Case Value
            Case 0
                Value = "否"
            Case 1
                Value = "是"
            Case Else
                Value = "否"
        End Select
    End If
    If giFld.FieldName = "PreInvoice" Then
        Select Case Value
            Case 0
                Value = "後開制"
            Case 1
                Value = "預開制"
            Case Else
                Value = "後開制"
        End Select
    End If
    If giFld.FieldName = "InvoiceKind" Then
        Select Case Value
            Case 0
                Value = "電子計算機"
            Case 1
                Value = "電子發票"
            Case Else
                Value = "電子計算機"
                
        End Select
    End If
    '0=一般帳單,1=電子帳單
    If giFld.FieldName = "BillMailKind" Then
        Select Case Value
            Case 0
                Value = "一般帳單"
            Case 1
                Value = "電子帳單"
            Case Else
                Value = "一般帳單"
        End Select
    End If
    '1=修改, 2=刪除, 3=新增
    If giFld.FieldName = "FuncType" Then
        Select Case Value
            Case 1
                Value = "修改"
            Case 2
                Value = "刪除"
            Case 3
                Value = "新增"
        End Select
    End If
End Sub
