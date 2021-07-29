VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "CsMulti.ocx"
Begin VB.Form frmBillA 
   BorderStyle     =   1  '單線固定
   Caption         =   "帳單資訊 [frmBillA]"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9390
   Icon            =   "frmBillA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9390
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdDelete 
      Caption         =   "刪除"
      Height          =   435
      Left            =   8130
      TabIndex        =   2
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "F2.查詢"
      Height          =   435
      Left            =   1470
      TabIndex        =   1
      Top             =   180
      Width           =   1095
   End
   Begin VB.Frame fraQueryWhere 
      Caption         =   "查詢條件"
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   750
      Width           =   9135
      Begin CS_Multi.CSmulti gimVendor 
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   270
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   556
         ButtonCaption   =   "廠  商  別"
      End
      Begin Gi_Date.GiDate gdaShouldDate1 
         Height          =   315
         Left            =   1770
         TabIndex        =   4
         Top             =   720
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Date.GiDate gdaShouldDate2 
         Height          =   315
         Left            =   3240
         TabIndex        =   5
         Top             =   720
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "~"
         Height          =   255
         Left            =   3030
         TabIndex        =   10
         Top             =   780
         Width           =   195
      End
      Begin VB.Label Label1 
         Caption         =   "繳費截止日區間"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   9
         Top             =   780
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdImports 
      Caption         =   "F4.匯入資料"
      Height          =   435
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   1095
   End
   Begin VB.Frame fraQueryResult 
      Caption         =   "查詢結果"
      Height          =   4965
      Left            =   90
      TabIndex        =   6
      Top             =   2040
      Width           =   9165
      Begin prjGiGridR.GiGridR ggrData 
         Height          =   4305
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7594
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
Attribute VB_Name = "frmBillA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FWhere As String
Private FSQL As String
Private rsQuery As New ADODB.Recordset
Private rsTmp As New ADODB.Recordset
Private rsImport As New ADODB.Recordset
Private FblnDelFlag As Boolean
Private Sub cmdDelete_Click()
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim blnTrans As Boolean
    If ggrData.Recordset Is Nothing Then
        MsgBox "無任何資料可刪除！", vbInformation, "訊息"
        Exit Sub
    End If
    If ggrData.Recordset.RecordCount = 0 Then
        MsgBox "無任何資料可刪除！", vbInformation, "訊息"
        Exit Sub
    End If
    If MsgBox("是否確定刪除資料?", vbQuestion + vbYesNo, "詢問") = vbNo Then Exit Sub
    
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    If ggrData.Recordset.RecordCount > 0 Then
        ggrData.Recordset.MoveFirst
        gcnGi.BeginTrans
        blnTrans = True
        ggrData.Visible = False
        Do While Not ggrData.Recordset.EOF
            DoEvents
            If ggrData.Recordset("Choice") = "1" Then
                'aSQL = "DELETE SO315 WHERE ROWID='" & ggrData.Recordset("RowId") & "'"
                aSQL = "DELETE SO315 WHERE VenderCode = '" & ggrData.Recordset("VenderCode") & "'" & _
                        " AND VBillNo = '" & ggrData.Recordset("VBillNo") & "'"
                gcnGi.Execute aSQL
            End If
            ggrData.Recordset.MoveNext
        Loop
        gcnGi.CommitTrans
    End If
    FblnDelFlag = True
    ExeQuery
    Me.Enabled = True
    ggrData.Visible = True
    MsgBox "刪除完成！", vbInformation, "訊息"
    Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
  Me.Enabled = True
  ggrData.Visible = True
  If blnTrans Then gcnGi.RollbackTrans
  
  Call ErrSub(Me.Name, "cmdDelete_Click")
End Sub

Private Sub cmdImports_Click()
  On Error Resume Next
    Dim i As Long
    frmBillB.Show 1, Me
    Screen.MousePointer = vbHourglass
    If Not rsImport Is Nothing Then
        If rsImport.State = adStateOpen Then
            If rsImport.RecordCount > 0 Then
                rsImport.MoveFirst
                Call DefineRs(True)
                Do While Not rsImport.EOF
                    rsTmp.AddNew
                    For i = 0 To rsImport.Fields.Count - 1
                        If Len(rsImport(i) & "") <> "" Then
                            rsTmp.Fields(rsImport.Fields(i).Name) = rsImport.Fields(i)
                        End If
                    Next
                    rsTmp.Update
                    rsImport.MoveNext
                Loop
            End If
        End If
    End If
    CloseRecordset rsImport
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdQuery_Click()
  On Error GoTo ChkErr
    Me.Enabled = False
    Me.ggrData.Visible = False
    Screen.MousePointer = vbHourglass
    ExeQuery
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    Me.ggrData.Visible = True
    If rsQuery.RecordCount = 0 Then
        If Not FblnDelFlag Then
            MsgBox "查詢不到任何資料！", vbInformation, "訊息"
        End If
    End If
    FblnDelFlag = False
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdQuery_Click")
End Sub
Private Sub ExeQuery()
  On Error GoTo ChkErr
    FSQL = Empty
    Call subChoose
    FSQL = "SELECT A.RowId,A.* FROM SO315 A "
    If FWhere <> "" Then
        FSQL = FSQL & " WHERE " & FWhere
    End If
    If Not GetRS(rsQuery, FSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    DefineRs
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "ExeQuery")
End Sub
Private Sub subChoose()
  On Error GoTo ChkErr
    FWhere = Empty
    Call subAnd2(FWhere, " RealDate is null")
    If gdaShouldDate1.GetValue & "" <> "" Then
        Call subAnd2(FWhere, " ShouldDate >= To_Date('" & gdaShouldDate1.GetValue & "','yyyymmdd')")
    End If
    If gdaShouldDate2.GetValue & "" <> "" Then
        Call subAnd2(FWhere, " ShouldDate <= To_Date('" & gdaShouldDate2.GetValue & "','yyyymmdd')")
    End If
    If Trim(gimVendor.GetDispStr) <> "" Then
        Call subAnd2(FWhere, " VenderCode " & gimVendor.GetQryStr)
    End If
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
    If KeyCode = 27 Then
        Unload Me
    End If
    If KeyCode = vbKeyF2 Then
        cmdQuery.Value = True
    End If
    If KeyCode = vbKeyF4 Then
        cmdImports.Value = True
    End If
End Sub

Private Sub Form_Load()
  On Error Resume Next
    FblnDelFlag = False
    Call subGrid
    Call subGim
    Call DefineRs(True)
End Sub
Private Sub subGrid()
    On Error Resume Next
    Dim mFlds As New GiGridFlds
    With mFlds
        .Add "Choice", , , , , "選取"
        .Add "VenderCode", , , , , "廠商別" & Space(5)
        .Add "VBillNo", , , , , "帳單號碼" & Space(2)
        .Add "VCustId", , , , , "訂戶號碼" & Space(2)
        .Add "CitemName", , , , , "繳費項目" & Space(5)
        .Add "ShouldAmt", , , , , "繳費金額" & Space(1)
        .Add "Commission", , , , , "手續費" & Space(2)
        .Add "ShouldDate", , , , , "繳費截止日" & Space(2)
        .Add "RealStartDate", , , , , "有效起始日" & Space(2)
        .Add "RealStopDate", , , , , "有效截止日" & Space(2)
        .Add "RealDate", , , , , "繳費日期" & Space(5)
        .Add "AcceptCompCode", , , , , "已受理公司別"
'        .Add "Detail", , , , , "明細" & Space(40)
'        .Add "Remark", , , , , "收費起日" & Space(40)
'        .Add "CustName", , , , , "姓名" & Space(30)
'        .Add "ID", , , , , "身份證號碼" & Space(5)
'        .Add "ZipCode", , , , , "區碼"
'        .Add "Phone", , , , , "電話" & Space(10)
'        .Add "CanCompCode", , , , , "可扣款公司別" & Space(10)
    End With
    ggrData.AllFields = mFlds
    ggrData.SetHead
   
End Sub
Private Function DefineRs(Optional blnLoad As Boolean = False) As Boolean
  On Error GoTo ChkErr
    Dim i As Long
    With rsTmp
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Fields.Append ("Choice"), adBSTR, 1, adFldIsNullable
        For i = 0 To rsQuery.Fields.Count - 1
            .Fields.Append rsQuery.Fields(i).Name, adBSTR, rsQuery.Fields(i).DefinedSize, adFldIsNullable
        Next i
    End With
    rsTmp.Open
    
    If blnLoad Then
        FSQL = "SELECT A.RowId,A.* FROM SO315 A  Where 1=0 "
        If Not GetRS(rsQuery, FSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    End If
    
    Do While Not rsQuery.EOF
        rsTmp.AddNew
        For i = 0 To rsQuery.Fields.Count - 1
            If Len(rsQuery.Fields(i).Value & "") > 0 Then
                rsTmp.Fields(rsQuery.Fields(i).Name).Value = rsQuery.Fields(i).Value & ""
            End If
        Next i
        rsTmp("Choice").Value = "1"
        
        
        rsTmp.Update
        rsQuery.MoveNext
    Loop
    DefineRs = True
    Set ggrData.Recordset = rsTmp
    ggrData.Refresh
    On Error Resume Next
    
    Exit Function
ChkErr:
  ErrSub Me.Name, "OpenData"
End Function

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    CloseRecordset rsQuery
    CloseRecordset rsTmp
End Sub

Private Sub ggrData_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error GoTo ChkErr
    If UCase(giFld.FieldName) = "CHOICE" Then
        If Value = 1 Then Value = vbRed
    End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "ggrData_ColorCellData"
    
End Sub

Private Sub ggrData_DblClick()
  On Error Resume Next
    If Val(ggrData.Recordset("Choice") & "") = 1 Then
        ggrData.Recordset("Choice") = "0"
        
    Else
        ggrData.Recordset("Choice") = "1"
    End If
    ggrData.Recordset.Update
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If UCase(giFld.FieldName) = "CHOICE" Then
        If Value = "1" Then
            Value = "V"
        Else
            Value = ""
        End If
    End If
    Exit Sub
ChkErr:
  ErrSub Me.Name, "ggrData_ShowCellData"
  
End Sub
Public Sub subGim()
    SetgiMulti gimVendor, "Vendercode", "Vender", "SO315A", "代碼", "名稱"
End Sub
Public Property Set uRs(ByRef vData As ADODB.Recordset)
    Set rsImport = vData
End Property
