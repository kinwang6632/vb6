VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1131F2 
   BorderStyle     =   1  '單線固定
   Caption         =   "階段式優惠歷史資料 [SO1131F2]"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11655
   Icon            =   "SO1131F2.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   11655
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   345
      Left            =   10290
      TabIndex        =   1
      Top             =   3870
      Width           =   1215
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   3720
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "請按左鍵兩次,選取客戶資料"
      Top             =   0
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   6562
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseCellForeColor=   -1  'True
   End
End
Attribute VB_Name = "frmSO1131F2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsData As ADODB.Recordset
Dim lngCustId As Long
Dim strCitemCode As String
Dim strContNo As String
Dim strFaciSeqNo As String
Private strFaciSNo As String
Private strServiceType As String
Private strCustid As String
Public Property Let uCustId(ByVal vData As String)
    lngCustId = vData
End Property



Public Property Let uCitemCode(ByVal vData As String)
    strCitemCode = vData
End Property

Public Property Let uContNo(ByVal vData As String)
    strContNo = vData
End Property
'#4165 要儲存FaciSNo By Kin 2008/10/23
Public Property Let uFaciSNo(ByVal vData As String)
    strFaciSNo = vData
End Property
Public Property Let uFaciSeqNo(ByVal vData As String)
    strFaciSeqNo = vData
End Property


Public Property Let uServiceType(ByVal vData As String)
    strServiceType = vData
End Property

Private Sub cmdCancel_Click()
  On Error Resume Next
    CloseRecordset rsData
    Unload Me
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    subGrd
    OpenData
    Exit Sub
ChkErr:
 Call ErrSub(Me.Name, "Form_Load")
End Sub
Private Sub OpenData()
    On Error GoTo ChkErr
    Dim strSQL As String
        Set rsData = New ADODB.Recordset
        '#5072 服務別與客編改成根據傳來的RecordSet為主 By Kin 2009/05/05
        strSQL = "Select RowId,A.* From " & GetOwner & "SO003ALog A Where A.CustId = " & _
            lngCustId & " And A.CompCode = " & gCompCode
            
        If strContNo <> "" Then strSQL = strSQL & " And ContNo = '" & strContNo & "'"
        strSQL = strSQL & " Order By DiscountDate1"
        
        If Not GetRS(rsData, strSQL, , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        'Set rsTmp = rsData.Clone
        Set ggrData.Recordset = rsData
        ggrData.Refresh
        If rsData.RecordCount > 0 Then
            rsData.AbsolutePosition = 1
        End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub
Private Sub subGrd()
    On Error Resume Next
    Dim mFlds As New GiGridFlds
        mFlds.Add "ContStartDate", giControlTypeDate, , , , "原合約起日"
        mFlds.Add "ContStopDate", giControlTypeDate, , , , "原合約迄日"
        mFlds.Add "DiscountDate1", giControlTypeDate, , , , "原優待起始日"
        mFlds.Add "DiscountDate2", giControlTypeDate, , , , "原優待截止日"
        mFlds.Add "DiscountAmt", , , , , "原優待金額"
        mFlds.Add "Period", , , , , "原每次期數"
        mFlds.Add "MonthAmt", , , , , "原單月金額"
        mFlds.Add "DayAmt", , , , , "原單日金額"
        
        mFlds.Add "NContStartDate", giControlTypeDate, , , , "新合約起日"
        mFlds.Add "NContStopDate", giControlTypeDate, , , , "新合約迄日"
        mFlds.Add "NDiscountDate1", giControlTypeDate, , , , "新優待起始日"
        mFlds.Add "NDiscountDate2", giControlTypeDate, , , , "新優待截止日"
        mFlds.Add "NDiscountAmt", , , , , "新優待金額"
        mFlds.Add "NPeriod", , , , , "新每次期數"
        mFlds.Add "NMonthAmt", , , , , "新單月金額"
        mFlds.Add "NDayAmt", , , , , "新單日金額"
        
        mFlds.Add "ModifyDep", , , , , "修改單位      "
        mFlds.Add "ModifyEn", , , , , "修改人員代碼"
        mFlds.Add "ModifyName", , , , , "修改人員名稱     "
        mFlds.Add "ModifyNotes", , , , , "修改原因                                             "
        
'        mFlds.Add "Punish", , , , , "違約處份"
'        mFlds.Add "StopFlag", , , , , "停用"
'        mFlds.Add "StopDate", giControlTypeDate, , , , "  停用日期  "
'        mFlds.Add "UpdTime", , , , , "異動時間" & Space(10)
'        mFlds.Add "UpdEn", , , , , "異動人員"
        ggrData.AllFields = mFlds

End Sub

Private Sub ggrData_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    Select Case UCase(giFld.FieldName)
        Case UCase("NContStartDate")
            Value = vbRed
        Case UCase("NContStopDate")
            Value = vbRed
        Case UCase("NDiscountDate1")
            Value = vbRed
        Case UCase("NDiscountDate2")
            Value = vbRed
        Case UCase("NDiscountAmt")
            Value = vbRed
        Case UCase("NPeriod")
            Value = vbRed
        Case UCase("NMonthAmt")
            Value = vbRed
        Case UCase("NDayAmt"), UCase("ModifyDep"), UCase("ModifyEn"), UCase("ModifyName"), UCase("ModifyNotes")
            Value = vbRed
    End Select
'        If giFld.FieldName = "CancelFlag" Then
'            If Value <> 0 Then Value = vbRed
'        End If
'        If giFld.FieldName = "FaciSNo" Then
'            If ggrData.Recordset("CitemCode") & "" <> "" Then
'                If GetRsValue("Select ByHouse From " & GetOwner & "CD019 Where CodeNo = " & ggrData.Recordset("CitemCode")) = 1 Then
'                    Value = "BY戶"
'                End If
'            End If
'        End If
End Sub

