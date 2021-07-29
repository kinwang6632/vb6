VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO9700B 
   Caption         =   "Form1"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   12405
   Icon            =   "SO9700B.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   12405
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.Frame Frame1 
      Height          =   2325
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   12255
      Begin prjGiGridR.GiGridR ggrData1 
         Height          =   1935
         Left            =   0
         TabIndex        =   2
         Top             =   180
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   3413
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
      Begin prjGiGridR.GiGridR ggrData4 
         Height          =   1995
         Left            =   5460
         TabIndex        =   4
         Top             =   180
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   3519
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
   Begin prjGiGridR.GiGridR ggrData3 
      Height          =   1725
      Left            =   0
      TabIndex        =   1
      Top             =   4320
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   3043
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
   Begin prjGiGridR.GiGridR ggrData2 
      Height          =   1815
      Left            =   30
      TabIndex        =   3
      Top             =   2430
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   3201
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
Attribute VB_Name = "frmSO9700B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsSource As Recordset
Private rsDefService As Recordset
Private rs004 As Recordset
Private rs007 As Recordset
Private rs033 As Recordset
Private fCustId As String
Private fOrderNo As String
Private fWhere As String
Private Sub Open004Data()
  On Error GoTo ChkErr
    Dim aSQL As String
    Set rs004 = New Recordset
    aSQL = "SELECT A.* FROM " & GetOwner & "SO004 A," & GetOwner & "SO007 B " & _
        " WHERE A.SNO=B.SNO " & fWhere
    If Not GetRS(rs004, aSQL, gcnGi, adUseClient, adOpenDynamic, adLockOptimistic) Then Exit Sub
    Set ggrData4.Recordset = rs004
    ggrData4.Refresh
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Open004Data")
End Sub
Private Sub Open033Data()
  On Error GoTo ChkErr
  
  
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Open033Data")
End Sub
Private Sub Open007Data()
  On Error GoTo ChkErr
    Dim aSQL As String
    Set rs007 = New Recordset
    aSQL = "SELECT B.* FROM " & GetOwner & "SO007 B WHERE 1=1 " & fWhere
    If Not GetRS(rs007, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Set ggrData2.Recordset = rs007
    ggrData2.Refresh
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Open007Data")
End Sub
Private Sub GetWhere()
  On Error GoTo ChkErr
  If Len(fOrderNo & "") > 0 Then
        fWhere = " AND B.ORDERNO = '" & fOrderNo & "'"
    Else
        fWhere = " AND B.CUSTID = " & fCustId
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "GetWhere")
End Sub
Private Sub OpenServiceData()
  On Error GoTo ChkErr
    Dim rsCD046 As New Recordset
    Dim rs033 As New Recordset
    Dim aSQL As String
    
    Dim i As Integer
    If Not GetRS(rsCD046, "SELECT * FROM " & GetOwner & "CD046", gcnGi, adUseClient, adOpenDynamic, adLockOptimistic) Then Exit Sub
    
    For i = 0 To 1
        'X2 贈品,X1本票
        If aSQL <> "" Then aSQL = aSQL & " UNION ALL "
        aSQL = aSQL & "SELECT SUM(A.SHOULDAMT) SHOULDAMT,A.SERVICETYPE " & _
            " FROM " & GetOwner & "SO0" & 33 + i & " A," & GetOwner & "SO007 B," & _
            GetOwner & "CD019 C " & _
            " WHERE A.BILLNO = B.SNO AND A.CITEMCODE = C.CODENO AND NVL(C.REFNO,0) <> 18 " & _
            " AND NVL(A.PTCODE,0) <> 6 " & fWhere & _
            " GROUP BY A.SERVICETYPE " & _
            " UNION ALL " & _
            "SELECT SUM(A.SHOULDAMT) SHOULDAMT,'X2' SERVICETYPE " & _
            " FROM " & GetOwner & "SO0" & 33 + i & " A," & GetOwner & "SO007 B," & _
            GetOwner & "CD019 C " & _
            " WHERE A.BILLNO = B.SNO AND A.CITEMCODE = C.CODENO AND NVL(C.REFNO,0) = 18 " & _
            " AND NVL(A.PTCODE,0) <> 6 " & fWhere & _
            " GROUP BY A.CUSTID " & _
            " UNION ALL " & _
            "SELECT SUM(A.SHOULDAMT) SHOULDAMT,'X1' SERVICETYPE " & _
            " FROM " & GetOwner & "SO0" & 33 + i & " A," & GetOwner & "SO007 B," & _
            GetOwner & "CD019 C " & _
            " WHERE A.BILLNO = B.SNO AND A.CITEMCODE = C.CODENO AND NVL(C.REFNO,0) <> 18 " & _
            " AND NVL(A.PTCODE,0) = 6 " & fWhere & _
            " GROUP BY A.CUSTID "
    Next i
    If Not GetRS(rs033, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    rsCD046.MoveFirst
    Dim aTotal As Long
    Dim aOrdTotal As Long
    Do While Not rsCD046.EOF
        rsDefService.AddNew
        rsDefService("Custid") = fCustId
        rsDefService("OrderNo") = fOrderNo
        rsDefService("ServiceType") = rsCD046("Description")
        If rs033.RecordCount > 0 Then
            rs033.MoveFirst
            rs033.Find "Servicetype='" & rsCD046("CodeNo") & "'", , , 1
            If Not rs033.EOF Then
                rsDefService("AMT") = rs033("ShouldAmt")
                aTotal = aTotal + rs033("ShouldAmt")
                aOrdTotal = aTotal
            Else
                rsDefService("AMT") = "0"
            End If
        End If
        rsDefService.Update
        rsCD046.MoveNext
    Loop
    For i = 0 To 1
        rsDefService.AddNew
        rsDefService("Custid") = fCustId
        rsDefService("OrderNo") = fOrderNo
        If i = 0 Then
            rsDefService("ServiceType") = "本票"
        Else
            rsDefService("ServiceType") = "贈品"
        End If
        If rs033.RecordCount > 0 Then
            rs033.MoveFirst
            rs033.Find "Servicetype='" & "X" & 1 + i & "'", , , 1
            If Not rs033.EOF Then
                rsDefService("AMT") = rs033("ShouldAmt")
                aTotal = aTotal + rs033("ShouldAmt")
            Else
                rsDefService("AMT") = "0"
            End If
        End If
        rsDefService.Update
    Next i
    For i = 0 To 1
        rsDefService.AddNew
        rsDefService("Custid") = fCustId
        rsDefService("OrderNo") = fOrderNo
        If i = 0 Then
            rsDefService("ServiceType") = "總計金額"
            rsDefService("AMT") = aTotal
        Else
            rsDefService("ServiceType") = "訂單總金額"
            rsDefService("AMT") = aOrdTotal
        End If
        rsDefService.Update
    Next i
          
  Set ggrData1.Recordset = rsDefService
  ggrData1.Refresh
  On Error Resume Next
  CloseRecordset rsCD046
  CloseRecordset rs033
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "OpenServiceData")
End Sub
Public Property Let uCustId(ByVal vData As String)
    fCustId = vData
End Property
Public Property Let uOrderNo(ByVal vData As String)
    fOrderNo = vData
End Property
Public Property Set uRsSource(ByVal vData As Recordset)
    Set rsSource = vData
End Property

Private Sub Form_Load()
  On Error GoTo ChkErr
    GetWhere
    DefineRs
    subGrd
    OpenServiceData
    Open004Data
    Open007Data
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub
Private Sub DefineRs()
  On Error GoTo ChkErr
  
    Set rsDefService = New Recordset
    With rsDefService
        If .State = adStateOpen Then .Close
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Fields.Append ("CustId"), adBSTR, 10, adFldIsNullable
        .Fields.Append ("OrderNo"), adBSTR, 20, adFldIsNullable
        .Fields.Append ("ServiceType"), adBSTR, 30, adFldIsNullable
        .Fields.Append ("AMT"), adBSTR, 20, adFldIsNullable
        .Open
    End With
    Exit Sub
ChkErr:
  ErrSub Me.Name, "DefineRs"
End Sub
Private Sub subGrd()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    Dim mFlds2 As New GiGridFlds
    Dim mFlds3 As New GiGridFlds
    Dim mFlds4 As New GiGridFlds
    With mFlds
        .Add "CustId", , , , , "客編" & Space(5)
        .Add "ORDERNO", , , , , "訂單單號" & Space(10)
        .Add "ServiceType", , , , , "產品別" & Space(10)
        .Add "AMT", , , , , "小計    "
        
    End With
    With ggrData1
        .UseCellForeColor = True
        .AllFields = mFlds
        .SetHead
    End With
    
    With mFlds2
        .Add "InstName", , , , , "派工類別" & Space(15)
        .Add "SNo", , , , , "單號" & Space(15)
        .Add "AcceptTime", , , , , "受理日期" & Space(20)
        .Add "ResvTime", , , , , "預約日期" & Space(20)
        .Add "BPName", , , , , "優惠組合" & Space(20)
        .Add "PromName", , , , , "促銷方案" & Space(20)
        .Add "MediaName", , , , , "介紹媒介" & Space(20)
        .Add "IntroName", , , , , "介紹人" & Space(10)
        .Add "BulletinName", , , , , "消息來源" & Space(10)
        .Add "Note", , , , , "備註" & Space(100)
    End With
    With ggrData2
        .UseCellForeColor = True
        .AllFields = mFlds2
        .SetHead
    End With
    
    With mFlds3
        .Add "CitemName", , , , , "產品名稱" & Space(15)
        .Add "RealPeriod", , , , , "期數" & Space(5)
        .Add "ShouldAmt", , , , , "金額" & Space(15)
        .Add "RealStartDate", , , , , "起始日" & Space(20)
        .Add "RealStopDate", , , , , "截止日" & Space(20)
        .Add "ContStartDate", , , , , "合約起日" & Space(20)
        .Add "ContStopDate", , , , , "合約迄日" & Space(20)
    End With
    With ggrData3
        .UseCellForeColor = True
        .AllFields = mFlds3
        .SetHead
    End With
    
     With mFlds4
        .Add "DeclarantName", , , , , "申裝人姓名" & Space(15)
        .Add "FaciName", , , , , "設備名稱" & Space(25)
        .Add "ContTel", , , , , "申請人電話" & Space(15)
        .Add "PTName", , , , , "保證金種類" & Space(20)
    End With
    With ggrData4
        .UseCellForeColor = True
        .AllFields = mFlds4
        .SetHead
    End With
    
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGrd")
End Sub
