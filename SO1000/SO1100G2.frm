VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1100G2 
   BorderStyle     =   1  '單線固定
   Caption         =   "指定設備[SO1100G2]"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   10860
   Icon            =   "SO1100G2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   10860
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消&X"
      Height          =   405
      Left            =   6510
      TabIndex        =   7
      Top             =   5880
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定.F2"
      Height          =   405
      Left            =   3060
      TabIndex        =   6
      Top             =   5880
      Width           =   1125
   End
   Begin VB.Frame Frame3 
      Caption         =   "收費資料"
      Height          =   3315
      Left            =   5550
      TabIndex        =   4
      Top             =   2460
      Width           =   5445
      Begin prjGiGridR.GiGridR ggrData3 
         Height          =   2835
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "請按左鍵兩次,選取發票資料"
         Top             =   210
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   5001
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
   Begin VB.Frame Frame2 
      Caption         =   "週期性項目"
      Height          =   3345
      Left            =   90
      TabIndex        =   2
      Top             =   2430
      Width           =   5445
      Begin prjGiGridR.GiGridR ggrData2 
         Height          =   2835
         Left            =   150
         TabIndex        =   3
         ToolTipText     =   "請按左鍵兩次,選取發票資料"
         Top             =   240
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   5001
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
   Begin VB.Frame Frame1 
      Caption         =   "設備資訊"
      Height          =   2265
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   10695
      Begin prjGiGridR.GiGridR ggrData 
         Height          =   1905
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "請按左鍵兩次,選取發票資料"
         Top             =   240
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   3360
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
Attribute VB_Name = "frmSO1100G2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsDefTmp  As ADODB.Recordset
Private rsDefTmp2 As ADODB.Recordset
Private rsDefTmp3 As ADODB.Recordset
Private strSEQs As String
Private strMasterId As String
Private strPRField As String
Private blnUCCodeIsNull As Boolean
Private str004s As String
Private str003s As String
Private str033s As String
Private strCitemCodes As String
Private blnAddCATV As Boolean
Private str033RowId As String
Private blnStop As Boolean
Private strChoice033 As String
Private strSourceAcc As String
Private blnSelectEnabled As Boolean
Private Sub cmdCancel_Click()
  On Error Resume Next
        
    Unload Me
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    If blnStop Then
        Exit Sub
        Unload Me
    End If
    ggrData3.Blank = True
    If rsDefTmp.RecordCount > 0 Then
        str004s = Empty
        rsDefTmp.MoveFirst
        Do While Not rsDefTmp.EOF
            If rsDefTmp("SeqNo") & "" <> "" And Val(rsDefTmp("Choice") & "") = 1 Then
                If str004s = Empty Then
                    str004s = "'" & rsDefTmp("SeqNo") & "'"
                Else
                    str004s = str004s & ",'" & rsDefTmp("SeqNo") & "'"
                End If
                
            End If
            rsDefTmp.MoveNext
        Loop
    Else
        str004s = Empty
    End If
    If rsDefTmp2.RecordCount > 0 Then
        str003s = Empty
        rsDefTmp2.MoveFirst
        Do While Not rsDefTmp2.EOF
            If str003s = Empty Then
                str003s = "'" & rsDefTmp2("SeqNo") & "'"
            Else
                str003s = str003s & ",'" & rsDefTmp2("SeqNo") & "'"
            End If
            rsDefTmp2.MoveNext
        Loop
    Else
        str003s = Empty
    End If
    If rsDefTmp3.RecordCount > 0 Then
        str033s = Empty
        str033RowId = Empty
        rsDefTmp3.MoveFirst
        Do While Not rsDefTmp3.EOF
            If Val(rsDefTmp3("Choice") & "") = 1 Then
                If str033s = Empty Then
                    str033s = "'" & rsDefTmp3("PK") & "'"
                    str033RowId = "'" & rsDefTmp3("RowId") & "'"
                Else
                    str033s = str033s & ",'" & rsDefTmp3("PK") & "'"
                    str033RowId = str033RowId & ",'" & rsDefTmp3("RowId") & "'"
                End If
            End If
            rsDefTmp3.MoveNext
        Loop
    Else
        str033s = Empty
        str033RowId = Empty
    End If
    ggrData3.Blank = False
    Unload Me
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdOK_Click")
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ChkErr
    If KeyCode = vbKeyF2 Then cmdOK.Value = True
    If KeyCode = vbKeyEscape Then cmdCancel.Value = True
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Set rsDefTmp = New ADODB.Recordset
    Set rsDefTmp2 = New ADODB.Recordset
    Set rsDefTmp3 = New ADODB.Recordset
    blnAddCATV = AddCATV
    Call OpenData(False)
    Call Open003Data
    Call Open033Data(strSEQs, blnAddCATV, True)
    Call subGrd
    Call subGrd2
    Call subGrd3
    If blnAddCATV Then AddCATVData
    Call Filter003(rsDefTmp, rsDefTmp2)
    Call Filter033(rsDefTmp, rsDefTmp3)
    If (blnStop) Or (Not blnSelectEnabled) Then
        Frame1.Enabled = False
        Frame2.Enabled = False
        Frame3.Enabled = False
        cmdOK.Enabled = False
    End If
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub DefineRs()
  On Error GoTo ChkErr
    With rsDefTmp
        If .State = adStateOpen Then
            .Close
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Fields.Append "Choice", adBSTR, 1, adFldIsNullable
        .Fields.Append "ServiceType", adBSTR, 10, adFldIsNullable
        .Fields.Append "FaciStatus", adBSTR, 30, adFldIsNullable
        .Fields.Append "FaciName", adBSTR, 50, adFldIsNullable
        .Fields.Append "FaciSNO", adBSTR, 50, adFldIsNullable
        .Fields.Append "DECLARANTNAME", adBSTR, 80, adFldIsNullable
        .Fields.Append "ACCOUNTNO", adBSTR, 20, adFldIsNullable
        .Fields.Append "SEQNO", adBSTR, 10, adFldIsNullable
    End With
    rsDefTmp.Open
    Exit Sub
ChkErr:
  ErrSub Me.Name, "DefineRs"
End Sub
Private Sub subGrd()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    With mFlds
        .Add "Choice", , , , , "選取"
        .Add "ServiceType", , , , , "服務別"
        .Add "FaciStatus", , , , , "設備狀態" & Space(5)
        .Add "FaciName", , , , , "品名" & Space(25)
        .Add "FaciSNO", , , , , "設備序號" & Space(10)
        .Add "DECLARANTNAME", , , , , "申請人姓名" & Space(15)
        .Add "ACCOUNTNO", , , , , "帳號" & Space(25)
    End With
    ggrData.AllFields = mFlds
    ggrData.SetHead
    Set ggrData.Recordset = rsDefTmp
    ggrData.Refresh
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGrd")
End Sub
Private Sub subGrd2()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    With mFlds
        '.Add "Choice", , , , , "選取"
        .Add "CITEMCODE", , , , , "代碼" & Space(2)
        .Add "CITEMNAME", , , , , "收費項目" & Space(10)
        .Add "PERIOD", , , , , "期數"
        .Add "AMOUNT", , , , , "金額" & Space(10)
        .Add "ACCOUNTNO", , , , , "扣帳帳號" & Space(10)
        .Add "CMNAME", , , , , "收費方式" & Space(10)
        .Add "STARTDATE", giControlTypeDate, , , , "起始日" & Space(10)
        .Add "STOPDATE", giControlTypeDate, , , , "截止日" & Space(10)
        .Add "FaciSNo", , , , , "物品序號" & Space(15)
        .Add "DeclarantName", , , , , "申請人姓名" & Space(10)
        .Add "FaciName", , , , , "項目名稱" & Space(10)
    End With
    ggrData2.AllFields = mFlds
    ggrData2.SetHead
    Set ggrData2.Recordset = rsDefTmp2
    ggrData2.Refresh
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGrd2")
End Sub
Private Sub subGrd3()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    With mFlds
        .Add "Choice", , , , , "選取"
        .Add "BillNo", , , , , "單據編號" & Space(18)
        .Add "CITEMNAME", , , , , "收費項目" & Space(10)
        .Add "REALPERIOD", , , , , "期數"
        .Add "SHOULDAMT", , , , , "金額" & Space(10)
        .Add "ACCOUNTNO", , , , , "扣帳帳號" & Space(10)
        .Add "CMNAME", , , , , "收費方式" & Space(10)
        .Add "REALSTARTDATE", giControlTypeDate, , , , "起始日" & Space(10)
        .Add "REALSTOPDATE", giControlTypeDate, , , , "截止日" & Space(10)
        .Add "SEQNO", , , , , "序號" & Space(15)
        .Add "FaciName", , , , , "項目名稱" & Space(10)
    End With
    ggrData3.AllFields = mFlds
    ggrData3.SetHead
    Set ggrData3.Recordset = rsDefTmp3
    ggrData3.Refresh
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGrd3")
End Sub


Private Sub OpenData(ByVal blnLoad As Boolean)
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim rs004 As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim aWhere As String
    Dim aWhere2 As String
    If GetSystemParaItem("FaciAuthority", gCompCode, gServiceType, "SO042") = 2 Then
        strPRField = "GetDate"
    Else
        strPRField = "PRDate"
    End If
    aWhere = " AND A." & strPRField & " IS NULL "
    If blnLoad Then
        aWhere = " AND 1=0"
    End If
    aWhere2 = Get004SEQNo
    If aWhere2 <> "" Then
        aWhere2 = " AND A.SEQNO In (" & aWhere2 & ")"
    End If
    
    aSQL = "SELECT Null Choice,Null FaciStatus, A.* " & _
            " FROM " & GetOwner & "SO004 A," & GetOwner & "CD022 B " & _
        " WHERE A.FACICODE=B.CODENO AND A.CUSTID=" & gCustId & _
        " AND B.REFNO IN(2,3,5,6,7,8,10) " & aWhere & aWhere2

    If Not GetRS(rs004, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
   
    Call DefineRsTmp(rs004, rsDefTmp)
    str004s = strSEQs
    '停用沒有MasterId,所以要根據收費項目把SEQ找出來,並選擇設備 By Kin 2010/04/15
    If blnStop And strSEQs = "" Then
        strSEQs = Empty
        If str003s <> "" Then
            aSQL = "SELECT SEQNO FROM " & GetOwner & "SO004 " & _
                " WHERE CUSTID=" & gCustId & " AND COMPCODE = " & gCompCode & _
                " AND SEQNO IN(  SELECT FACISEQNO FROM " & GetOwner & "SO003 " & _
                " WHERE CUSTID=" & gCustId & " AND COMPCODE=" & gCompCode & _
                " AND SEQNO IN(" & str003s & "))"
            If Not GetRS(rsTmp, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
            Do While Not rsTmp.EOF
                If strSEQs = "" Then
                    strSEQs = "'" & rsTmp("SEQNO") & "'"
                Else
                    strSEQs = strSEQs & ",'" & rsTmp("SEQNO") & "'"
                End If
            
                rsTmp.MoveNext
            Loop
        End If
    End If
        
    If Not rs004.EOF Then
        rs004.MoveFirst
        Do While Not rs004.EOF
            rsDefTmp.AddNew
            For i = 0 To rs004.Fields.Count - 1
                If UCase(rsDefTmp.Fields(i).Name) = UCase(rs004.Fields(i).Name) Then
                    rsDefTmp(rs004.Fields(i).Name) = rs004.Fields(rs004.Fields(i).Name).Value
                    rsDefTmp.Update
                End If
            Next i
            '如果有包含傳過來的SO004.SEQ要選擇起來 By Kin 2010/04/14
            
            If InStr(1, strSEQs, "'" & rsDefTmp("SEQNO") & "'") > 0 Then
                rsDefTmp("Choice") = "1"
                rsDefTmp.Update
            End If
            
            rs004.MoveNext
        Loop
    End If
    On Error Resume Next
    Call CloseRecordset(rs004)
    Call CloseRecordset(rsTmp)
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "OpenData")
End Sub
Private Sub DefineRsTmp(ByRef rsSource As ADODB.Recordset, ByRef rsDesc As ADODB.Recordset)
  On Error GoTo ChkErr
    Dim i As Long
     With rsDesc
        If .State = adStateOpen Then
            .Close
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        For i = 0 To rsSource.Fields.Count - 1
            Select Case UCase(rsSource.Fields(i).Name)
                Case "CHOICE"
                    .Fields.Append "Choice", adBSTR, 1, adFldIsNullable
                Case "FACISTATUS"
                    .Fields.Append "FaciStatus", adBSTR, 30, adFldIsNullable
                Case Else
                    .Fields.Append rsSource.Fields(i).Name, adBSTR, rsSource.Fields(i).DefinedSize, adFldIsNullable
            End Select
        Next i
    End With
    rsDesc.Open
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "DefineRsTmp")
End Sub
Private Sub Open003Data()
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim rs003 As New ADODB.Recordset
    Dim i As Long
    Dim aWhere As String
    If strCitemCodes <> "" Then
        aWhere = " AND CITEMCODE IN(" & strCitemCodes & ")"
    End If
    
    aSQL = "SELECT Null Choice,A.CITEMCODE,A.CITEMNAME,A.PERIOD,A.AMOUNT," & _
            " A.ACCOUNTNO,A.CMNAME,A.STARTDATE,A.ServiceType," & _
            " A.STOPDATE,A.SEQNO,A.FaciSNo,B.DeclarantName,B.FaciName,A.FaciSeqNo,CLCTDATE FROM " & _
            GetOwner & "SO003 A," & GetOwner & "SO004 B," & GetOwner & "CD022 C  " & _
            "WHERE A.CUSTID = " & gCustId & " AND A.COMPCODE=" & gCompCode & _
            " AND A.CustId=B.CustId(+) And A.CompCode=B.CompCode(+)" & _
            " And A.FaciSeqNo=B.SeqNo(+)" & _
            " AND B." & strPRField & " IS NULL " & _
            " AND C.REFNO IN(2,3,5,6,7,8,10) " & _
            " AND B.FACICODE=C.CODENO " & _
            " AND A.SERVICETYPE<>'C'" & aWhere
            
            
            
    aSQL = aSQL & " Union SELECT Null Choice,A.CITEMCODE,A.CITEMNAME,A.PERIOD,A.AMOUNT," & _
            " A.ACCOUNTNO,A.CMNAME,A.STARTDATE,A.SERVICETYPE," & _
            " A.STOPDATE,A.SEQNO,A.FaciSNo,B.DeclarantName,B.FaciName,A.FaciSeqNo,CLCTDATE FROM " & _
            GetOwner & "SO003 A," & GetOwner & "SO004 B " & _
            "WHERE A.CUSTID = " & gCustId & " AND A.COMPCODE=" & gCompCode & _
            " AND A.CustId=B.CustId(+) And A.CompCode=B.CompCode(+)" & _
            " And A.FaciSeqNo=B.SeqNo(+)" & _
            " And A.ServiceType='C'" & aWhere

    aSQL = "SELECT  DISTINCT * FROM (" & aSQL & ") ORDER BY CLCTDATE,CITEMCODE"
    '將所有SO003的資料都挑出來,後面在根據挑選的設備Filter SO003
    If Not GetRS(rs003, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Call DefineRsTmp(rs003, rsDefTmp2)
    If Not rs003.EOF Then
        rs003.MoveFirst
        Do While Not rs003.EOF
            rsDefTmp2.AddNew
            For i = 0 To rs003.Fields.Count - 1
                If UCase(rsDefTmp2.Fields(i).Name) = UCase(rs003.Fields(i).Name) Then
                    rsDefTmp2(rs003.Fields(i).Name) = rs003.Fields(i).Value
                    rsDefTmp2.Update
                End If
            Next i
'            If ChkExists003(rsDefTmp, rsDefTmp2("FaciSeqNo") & "") Then
'                rsDefTmp2("Choice") = "1"
'                rsDefTmp2.Update
'            End If
            rs003.MoveNext
        Loop
    End If
    On Error Resume Next
    Call CloseRecordset(rs003)
    
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Open003Data")
End Sub
Private Sub Open033Data(ByVal aFaciSeqNo As String, ByVal blnCATV As Boolean, Optional ByVal blnLoad As Boolean = False)
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim strOrder As String
    Dim i As Long
    Dim rs033 As New ADODB.Recordset
    Dim aCATVSQL As String
    
    strOrder = " ORDER BY A.REALDATE DESC,A.BILLNO DESC,A.SEQNO DESC "
    aSQL = "SELECT Null Choice, A.CITEMCODE,A.CITEMNAME,A.REALPERIOD,A.SHOULDAMT,A.ACCOUNTNO,A.CMNAME,A.REALSTARTDATE," & _
                    "A.REALSTOPDATE,A.SEQNO,A.BILLNO || A.ITEM PK,A.ROWID,A.FaciSNo," & _
                    "B.DeclarantName,B.FaciName,A.SERVICETYPE,A.RealDate,A.BillNo FROM " & _
                    GetOwner & "SO033 A," & GetOwner & "SO004 B, " & GetOwner & "CD022 C " & _
                    " Where 1=0 "
    If Not GetRS(rs033, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    
    
    Call DefineRsTmp(rs033, rsDefTmp3)
    '根據挑選的設備將SO033找出來
    If aFaciSeqNo <> Empty Then
        aSQL = "SELECT DISTINCT Null Choice, CITEMCODE,CITEMNAME,REALPERIOD,SHOULDAMT,ACCOUNTNO,CMNAME,REALSTARTDATE," & _
                                "REALSTOPDATE,SEQNO,PK,RowID,FaciSNo,DeclarantName," & _
                                "FaciName,SERVICETYPE,RealDate,BillNo  FROM " & _
                                 "(SELECT A.CITEMCODE,A.CITEMNAME,A.REALPERIOD,A.SHOULDAMT,A.ACCOUNTNO,A.CMNAME,A.REALSTARTDATE," & _
                                    "A.REALSTOPDATE,A.SEQNO,A.BILLNO || A.ITEM PK,A.ROWID,A.FaciSNo," & _
                                    "B.DeclarantName,B.FaciName,A.SERVICETYPE,A.RealDate,A.BillNo FROM " & _
                                 GetOwner & "SO033 A," & GetOwner & "SO004 B, " & GetOwner & "CD022 C " & _
                                 " WHERE A.CUSTID = " & gCustId & _
                                 " AND A.COMPCODE=" & gCompCode & _
                                 " AND A.CANCELFLAG <> 1  AND A.CHEVEN <> 1 " & _
                                 " AND A.CustId=B.CustId(+) And A.CompCode=B.CompCode(+)" & _
                                 " And A.FaciSeqNo=B.SeqNo(+) " & _
                                 " AND A.FaciSeqNo In(" & aFaciSeqNo & ")" & _
                                 " AND B." & strPRField & " IS NULL " & _
                                 " AND C.RefNo IN(2,3,5,6,7,8,10) " & _
                                 " AND B.FaciCode=C.CodeNo "
         aSQL = aSQL & " AND A.UCCode Is Not Null " & strOrder & ")" & _
                IIf(blnUCCodeIsNull, " Union " & aSQL & " And " & _
                " A.UCCode Is Null " & strOrder & ")", "")
                
        If Not GetRS(rs033, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        
        If Not rs033.EOF Then
            rs033.MoveFirst
            Do While Not rs033.EOF
                rsDefTmp3.AddNew
                For i = 0 To rs033.Fields.Count - 1
                    If UCase(rsDefTmp3.Fields(i).Name) = UCase(rs033.Fields(i).Name) Then
                        rsDefTmp3(rs033.Fields(i).Name) = rs033.Fields(i).Value
                        rsDefTmp3.Update
                    End If
                    
                Next i
                If Exitsts033(rsDefTmp3("PK") & "") Then
                    rsDefTmp3("Choice") = "1"
                    rsDefTmp3.Update
                    If InStr(1, strChoice033, "'" & rsDefTmp3("PK") & "'") <= 0 Then
                        If strChoice033 = "" Then
                            strChoice033 = "'" & rsDefTmp3("PK") & "'"
                        Else
                            strChoice033 = strChoice033 & ",'" & rsDefTmp3("PK") & "'"
                        End If
                    End If
                End If
                rs033.MoveNext
            Loop
        End If
    End If
    If blnAddCATV And blnCATV Then
                             
        aCATVSQL = " SELECT Null Choice, A.CITEMCODE,A.CITEMNAME,A.REALPERIOD,A.SHOULDAMT,A.ACCOUNTNO,A.CMNAME,A.REALSTARTDATE," & _
                        "A.REALSTOPDATE,Null SEQNO,A.BILLNO || A.ITEM PK,A.ROWID,Null FaciSNo," & _
                                    "B.CustName DeclarantName, 'CATV' FaciName,A.SERVICETYPE,A.RealDate,A.BillNo FROM " & _
                                 GetOwner & "SO033 A, " & GetOwner & "SO001 B " & _
                                 " WHERE A.CUSTID = " & gCustId & _
                                 " AND A.COMPCODE=" & gCompCode & _
                                 " AND A.CANCELFLAG <> 1  AND A.CHEVEN <> 1 " & _
                                 " AND A.SERVICETYPE='C'" & _
                                 " AND A.CustId=B.CustId " & _
                                 IIf(blnUCCodeIsNull, "", " AND A.UCCode Is Not Null ") & _
                                " ORDER BY A.REALDATE DESC,A.BILLNO DESC "
                                 
        If Not GetRS(rs033, aCATVSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If Not rs033.EOF Then
            rs033.MoveFirst
            Do While Not rs033.EOF
                rsDefTmp3.AddNew
                For i = 0 To rs033.Fields.Count - 1
                    If UCase(rsDefTmp3.Fields(i).Name) = UCase(rs033.Fields(i).Name) Then
                        rsDefTmp3(rs033.Fields(i).Name) = rs033.Fields(rs033.Fields(i).Name).Value
                        rsDefTmp3.Update
                    End If
                    
                Next i
                
                If Exitsts033(rsDefTmp3("PK") & "") Then
                    rsDefTmp3("Choice") = "1"
                    rsDefTmp3.Update
                    If InStr(1, strChoice033, "'" & rsDefTmp3("PK") & "'") <= 0 Then
                        If strChoice033 = "" Then
                            strChoice033 = "'" & rsDefTmp3("PK") & "'"
                        Else
                            strChoice033 = strChoice033 & ",'" & rsDefTmp3("PK") & "'"
                        End If
                    End If
                End If
                rs033.MoveNext
            Loop
        End If
                             
    End If
    
    
    If rsDefTmp3.RecordCount > 0 Then
        rsDefTmp3.MoveFirst
    Else
        strChoice033 = ""
    End If
    On Error Resume Next
    Call CloseRecordset(rs033)
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Open033Data")
End Sub
Private Function ChkExists003(rsSource As ADODB.Recordset, ByVal aSEQNO As String) As Boolean
  On Error GoTo ChkErr
    Dim bm As Variant
    Dim blnRet As Boolean
    blnRet = False
    If rsSource.RecordCount <= 0 Then Exit Function
    bm = rsSource.Bookmark
    rsSource.Find "SeqNo='" & aSEQNO & "'", , , 1
    If Not rsSource.EOF Then
        If rsSource("Choice") & "" = "1" Then
            blnRet = True
        Else
            blnRet = False
        End If
    End If
    On Error Resume Next
    ChkExists003 = blnRet
    rsSource.Bookmark = bm
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "ChkExists003")
End Function
Private Sub Filter003(rsSource As ADODB.Recordset, rsDesc As ADODB.Recordset)
  On Error GoTo ChkErr
    Dim bm As Variant
    Dim strFilter As String
    rsDesc.Filter = "FaciSeqNo='XXX'"
    If rsSource.EOF Then
        Set ggrData2.Recordset = rsDefTmp2
        ggrData2.Refresh
        Exit Sub
    End If
    bm = rsSource.Bookmark
    rsSource.MoveFirst
    
    Do While Not rsSource.EOF
        If Val(rsSource("Choice") & "") = 1 Then
            If rsSource("ServiceType") <> "C" Then
                If strFilter = Empty Then
                    strFilter = "FaciSeqNo='" & rsSource("SeqNo") & "'"
                Else
                    strFilter = strFilter & " Or FaciSeqNo='" & rsSource("SeqNo") & "'"
                End If
            Else
                If strFilter = Empty Then
                    strFilter = " ServiceType='C'"
                Else
                    strFilter = strFilter & " Or ServiceType='C'"
                End If
            End If
        End If
        rsSource.MoveNext
    Loop
    If strFilter <> "" Then
        rsDefTmp2.Filter = strFilter
    End If
    Set ggrData2.Recordset = rsDefTmp2
    ggrData2.Refresh
    On Error Resume Next
    rsSource.Bookmark = bm
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Filter003")
End Sub

Private Sub Filter033(rsSource As ADODB.Recordset, rsDesc As ADODB.Recordset)
  On Error GoTo ChkErr
    Dim bm As Variant
    Dim i As Long
    Dim aSEQNO As String
    Dim aCATV As Boolean
    If rsSource.EOF Then Exit Sub
    bm = rsSource.Bookmark
    rsSource.MoveFirst
    Do While Not rsSource.EOF
        If Val(rsSource("Choice") & "") = 1 Then
            If UCase(rsSource("ServiceType")) = "C" Then
                aCATV = True
                If rsSource("SEQNO") & "" <> "" Then
                    If aSEQNO = Empty Then
                        aSEQNO = "'" & rsSource("SeqNo") & "'"
                    Else
                        aSEQNO = aSEQNO & ",'" & rsSource("SeqNo") & "'"
                    End If
                End If
                
            Else
                If aSEQNO = Empty Then
                    aSEQNO = "'" & rsSource("SeqNo") & "'"
                Else
                    aSEQNO = aSEQNO & ",'" & rsSource("SeqNo") & "'"
                End If
            End If
        End If
        rsSource.MoveNext
    Loop
    Call Open033Data(aSEQNO, aCATV)
    On Error Resume Next
    Set ggrData3.Recordset = rsDesc
    ggrData3.Refresh
    rsSource.Bookmark = bm
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Filter033")
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
    Call CloseRecordset(rsDefTmp)
    Call CloseRecordset(rsDefTmp2)
    Call CloseRecordset(rsDefTmp3)
End Sub

Private Sub ggrData_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    If UCase(giFld.FieldName) = "CHOICE" Then
        If Value = 1 Then Value = vbRed
    End If
End Sub

Private Sub ggrData_DblClick()
  On Error Resume Next
    If ggrData.Recordset.RecordCount <= 0 Then Exit Sub
    If ggrData.Recordset("AccountNo") & "" <> "" Then
        If ggrData.Recordset("AccountNo") <> strSourceAcc Then
            If MsgBox("原帳號為:" & ggrData.Recordset("AccountNo") & vbCrLf & _
                        "新帳號為:" & strSourceAcc & vbCrLf & _
                        "是否替換?", vbQuestion + vbYesNo, "詢問") = vbNo Then
                Exit Sub
            Else
                ggrData.Recordset("AccountNo") = strSourceAcc
            End If
        End If
    End If
    
    ggrData3.Blank = True
    If Val(ggrData.Recordset("Choice") & "") = 0 Then
        ggrData.Recordset("Choice") = "1"
    Else
        ggrData.Recordset("Choice") = "0"
    End If
    ggrData.Recordset.Update
    Call Filter003(rsDefTmp, rsDefTmp2)
    Call Filter033(rsDefTmp, rsDefTmp3)

    Set ggrData2.Recordset = rsDefTmp2
    ggrData2.Refresh
    Set ggrData3.Recordset = rsDefTmp3
    ggrData3.Refresh
    ggrData3.Blank = False
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    If giFld.FieldName = "FaciStatus" And Trim(giFld.HeadName) = "設備狀態" Then
        If ggrData.Recordset("FACINAME") & "" <> "CATV" Then
            Value = GetFaciStatus(ggrData.Recordset, Value)
        End If
    ElseIf UCase(giFld.FieldName) = "CHOICE" Then
        If Value & "" = "1" Then
            Value = "V"
        Else
            Value = ""
        End If
    End If
End Sub
Property Let uSEQs(ByVal vData As String)
  On Error Resume Next
    strSEQs = vData
End Property
Property Get uSEQs() As String
  On Error Resume Next
    uSEQs = strSEQs
End Property
Property Let uMasterId(ByVal vData As String)
  On Error Resume Next
    strMasterId = vData
End Property
Property Let uUCCodeIsNull(ByVal vData As Boolean)
  On Error Resume Next
    blnUCCodeIsNull = vData
End Property


Private Sub ggrData2_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    If UCase(giFld.FieldName) = "CHOICE" Then
        If Value = 1 Then Value = vbRed
    End If
End Sub

Private Sub ggrData2_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    If UCase(giFld.FieldName) = "CHOICE" Then
        If Value & "" = "1" Then
            Value = "V"
        Else
            Value = ""
        End If
    End If
End Sub

Private Sub ggrData3_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    If UCase(giFld.FieldName) = "CHOICE" Then
        If Value = 1 Then Value = vbRed
    End If
End Sub

Private Sub ggrData3_DblClick()
  On Error GoTo ChkErr
    If ggrData3.Recordset.RecordCount <= 0 Then Exit Sub
    If Val(ggrData3.Recordset("Choice") & "") = 0 Then
        ggrData3.Recordset("Choice") = "1"
        If strChoice033 = "" Then
            strChoice033 = ggrData3.Recordset("PK")
        Else
            strChoice033 = strChoice033 & ",'" & ggrData3.Recordset("PK") & "'"
        End If
    Else
        ggrData3.Recordset("Choice") = "0"
        Dim i As Long
        i = InStr(1, strChoice033, "'" & ggrData3.Recordset("PK") & "'")
        If i = 1 Then
            strChoice033 = Replace(strChoice033, "'" & ggrData3.Recordset("PK") & "'", "")
        Else
            strChoice033 = Replace(strChoice033, ",'" & ggrData3.Recordset("PK") & "'", "")
        End If
        
    End If
    
    ggrData3.Recordset.Update
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "ggrData3_DblClick")
End Sub

Private Sub ggrData3_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    If UCase(giFld.FieldName) = "CHOICE" Then
        If Value & "" = "1" Then
            Value = "V"
        Else
            Value = ""
        End If
    End If
End Sub
'加入一個虛擬設備CATV的資料
Private Sub AddCATVData()
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim rs002 As New ADODB.Recordset
    Dim rs003 As New ADODB.Recordset
    '沒有待收單,不用新增
'    aSQL = "SELECT COUNT(1) FROM " & GetOwner & "SO033 " & _
'        " WHERE UCCODE IS NOT NULL AND CUSTID=" & gCustId & _
'        " AND SERVICETYPE='C' AND COMPCODE=" & gCompCode
'    If Val(GetRsValue(aSQL, gcnGi) & "") <= 0 Then
'        Exit Sub
'    End If
    
    
    aSQL = "SELECT A.*,B.CUSTNAME FROM " & GetOwner & "SO002 A," & GetOwner & "SO001 B" & _
        " WHERE CustStatusCode NOT IN(3,4,5) AND A.CUSTID=" & gCustId & _
        " AND A.SERVICETYPE='C' AND A.CUSTID=B.CUSTID "
    If Not GetRS(rs002, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    If Not rs002.EOF Then
        aSQL = "SELECT AccountNo FROM " & GetOwner & "SO003 " & _
            " WHERE CUSTID=" & gCustId & " AND SERVICETYPE='C' " & _
            " AND COMPCODE = " & gCompCode & " AND FACISEQNO NOT IN( " & _
            " SELECT SEQNO FROM " & GetOwner & "SO004 A WHERE " & _
            "  A.CUSTID=" & gCustId & " AND A.COMPCODE=" & gCompCode & ")" & _
            " AND ACCOUNTNO IS NOT NULL "
        If Not GetRS(rs003, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        ggrData.Recordset.AddNew
        ggrData.Recordset("FaciStatus") = rs002("CustStatusName") & ""
        ggrData.Recordset("FaciName") = "CATV"
        ggrData.Recordset("DECLARANTNAME") = rs002("Custname") & ""
        ggrData.Recordset("ServiceType") = "C"
        If Not rs003.EOF Then
            ggrData.Recordset("AccountNo") = rs003("AccountNo") & ""
        End If
        If ExistsCATV() Then ggrData.Recordset("Choice") = "1"
        ggrData.Recordset.Update
    End If
    On Error Resume Next
    Call CloseRecordset(rs002)
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "AddCATVData")
End Sub
Private Function ExistsCATV() As Boolean
  On Error GoTo ChkErr
    Dim aSQL As String
    
    If str003s = Empty Then Exit Function
    If str003s <> Empty Then
        aSQL = "SELECT COUNT(1) FROM " & GetOwner & "SO003 " & _
            " WHERE SERVICETYPE='C' AND SEQNO IN (" & str003s & ") " & _
            " AND COMPCODE=" & gCompCode & " AND CUSTID=" & gCustId
        If Val(GetRsValue(aSQL, gcnGi) & "") <= 0 Then Exit Function
    End If
    ExistsCATV = True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "ExistsCATV")
End Function
Private Function Exitsts033(ByVal aSerachString) As Boolean
  On Error GoTo ChkErr
    If InStr(strChoice033, "'" & aSerachString & "'") > 0 Then
        Exitsts033 = True
    Else
        Exitsts033 = False
    End If
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "Select033")
End Function

Private Function AddCATV() As Boolean
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim rs002 As New ADODB.Recordset
    AddCATV = False
    'ACH有選擇,要以CD068的收費項目過濾
    If strCitemCodes <> Empty Then
        aSQL = "SELECT COUNT(1) FROM " & GetOwner & "SO003 " & _
            " WHERE SERVICETYPE='C' AND CitemCode IN (" & strCitemCodes & ") " & _
            " AND COMPCODE=" & gCompCode & " AND CUSTID=" & gCustId
        If Val(GetRsValue(aSQL, gcnGi) & "") <= 0 Then Exit Function
    End If
    
    
    aSQL = "SELECT COUNT(1) FROM " & GetOwner & "SO004 A, " & GetOwner & "CD022 B " & _
        " WHERE A.CUSTID=" & gCustId & " AND A.SERVICETYPE='C'" & _
        " AND A.FACICODE=B.CODENO AND B.REFNO=10 "
    If Val(GetRsValue(aSQL, gcnGi) & "") > 0 Then Exit Function
    aSQL = "SELECT COUNT(1) FROM " & GetOwner & "SO002 A " & _
        " WHERE CustStatusCode NOT IN(3,4,5) AND A.CUSTID=" & gCustId & _
        " AND A.SERVICETYPE='C' "
    If Not GetRS(rs002, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Val(GetRsValue(aSQL, gcnGi) & "") = 0 Then Exit Function
    AddCATV = True
    On Error Resume Next
    Call CloseRecordset(rs002)
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "AddCATV")
End Function
'根據收費項目過濾可選的設備
Private Function Get004SEQNo() As String
  On Error GoTo ChkErr
    Dim strRet As String
    Dim aSQL As String
    Dim rsTmp As New ADODB.Recordset
    If strCitemCodes = "" Then
        Call CloseRecordset(rsTmp)
        Exit Function
    End If
    aSQL = "SELECT FACISEQNO FROM " & GetOwner & "SO003 " & _
        " WHERE CITEMCODE IN(" & strCitemCodes & ")" & _
        " AND CUSTID=" & gCustId & _
        " AND COMPCODE=" & gCompCode
    If Not GetRS(rsTmp, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    Do While Not rsTmp.EOF
        If strRet = "" Then
            strRet = "'" & rsTmp("FaciSeqNo") & "'"
        Else
            strRet = strRet & ",'" & rsTmp("FaciSeqNo") & "'"
        End If
        rsTmp.MoveNext
    Loop
    Get004SEQNo = strRet
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "Get004SEQNo")
End Function
'取出SO033.SEQ的CitemCode(因為SO106存的是SO003.SEQNO)
Private Function GetCitemCode(ByVal a033SEQ As String) As String
  On Error GoTo ChkErr
    Dim aSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strRet As String
    strRet = Empty
    aSQL = "SELECT CitemCode From " & GetOwner & "SO003 " & _
        " Where CustId=" & gCustId & _
        " AND SEQNO IN (" & a033SEQ & ")"
    If Not GetRS(rsTmp, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    strRet = rsTmp.GetString(, , , ",")
    If Right(strRet, 1) = "," Then
        strRet = Mid(strRet, 1, Len(strRet) - 1)
    End If
    GetCitemCode = strRet
    On Error Resume Next
    Call CloseRecordset(rsTmp)
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "GetCitemCode")
End Function
Property Let u033s(ByVal vData As String)
  On Error Resume Next
    str033s = vData
    strChoice033 = vData
End Property
Property Get u033s() As String
  On Error Resume Next
    u033s = str033s
End Property
Property Let u003s(ByVal vData As String)
  On Error Resume Next
     str003s = vData
    
End Property
Property Get u003s() As String
  On Error Resume Next
    u003s = str003s
    
End Property
Property Let u004s(ByVal vData As String)
  On Error Resume Next
    str004s = vData
End Property
Property Get u004s() As String
  On Error Resume Next
    u004s = str004s
End Property
Property Let uCitemCodes(ByVal vData As String)
  On Error Resume Next
    strCitemCodes = vData
End Property
Property Get u033RowId() As String
  On Error Resume Next
    u033RowId = str033RowId
End Property
Property Let u033RowId(ByVal vData As String)
  On Error Resume Next
    str033RowId = vData
End Property
Property Let uStop(ByVal vData As Boolean)
  On Error Resume Next
    blnStop = vData
End Property
Property Let uSourceAcc(ByVal vData As String)
  On Error Resume Next
    strSourceAcc = vData
End Property
Property Let uSelectEnabled(ByVal vData As Boolean)
    blnSelectEnabled = vData
End Property
