VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmBrower 
   BorderStyle     =   1  '單線固定
   Caption         =   "瀏覽畫面"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   11115
   Icon            =   "frmBrower.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdCancel 
      Caption         =   "F3.退  單"
      Height          =   405
      Left            =   1770
      TabIndex        =   3
      Top             =   5790
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "離開(&X)"
      Height          =   405
      Left            =   9930
      TabIndex        =   2
      Top             =   5790
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.重送授權碼"
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   5790
      Visible         =   0   'False
      Width           =   1275
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   2055
      Left            =   60
      TabIndex        =   0
      Top             =   1680
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   3625
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
   Begin prjGiGridR.GiGridR ggrData2 
      Height          =   1545
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   2725
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
   Begin prjGiGridR.GiGridR ggrData3 
      Height          =   1905
      Left            =   60
      TabIndex        =   5
      Top             =   3780
      Width           =   10935
      _ExtentX        =   19288
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
Attribute VB_Name = "frmBrower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsSourceData As ADODB.Recordset
Private WithEvents rsDefTmp As ADODB.Recordset
Attribute rsDefTmp.VB_VarHelpID = -1
Private WithEvents rsSO004 As ADODB.Recordset
Attribute rsSO004.VB_VarHelpID = -1
Private WithEvents rsSO563 As ADODB.Recordset
Attribute rsSO563.VB_VarHelpID = -1
Private fWebServiceUrl As String
Enum CompareDate
    Greater = 0
    Smaller = 1
End Enum

'是否可以選取
Public Property Let uCanSelect(ByVal vData As Boolean)
    fCanSelect = vData
End Property
'來源資料
Public Property Set uSourceRs(ByRef vData As ADODB.Recordset)
    Set rsSourceData = vData
End Property
Public Property Set uSO004Rs(ByRef vData As ADODB.Recordset)
    Set rsSO004 = vData
End Property
Public Property Set uSO563Rs(ByRef vData As ADODB.Recordset)
    Set rsSO563 = vData
End Property
Private Sub DefineRs()
  On Error GoTo ChkErr
    Dim i As Integer
    
    Set rsDefTmp = New ADODB.Recordset
    If fCanSelect Or True Then
        With rsDefTmp
            If .State = adStateOpen Then
                .Close
            End If
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            '***********************************************************
            
                .Fields.Append ("Choice"), adBSTR, 1, adFldIsNullable
            
            '***********************************************************
        End With
        
        For i = 0 To rsSourceData.Fields.Count - 1
            With rsDefTmp
                .Fields.Append rsSourceData.Fields(i).Name, adBSTR, rsSourceData.Fields(i).DefinedSize, adFldIsNullable
            End With
        Next i
        '------------------------------------------------------------------
        '裝置地點
        'rsDefTmp.Fields.Append "InitPlaceNo", adBSTR, 10, adFldIsNullable
        'rsDefTmp.Fields.Append "InitPlaceName", adBSTR, 50, adFldIsNullable
        '------------------------------------------------------------------
        '訂購狀態
        rsDefTmp.Fields.Append "OrderStatus", adBSTR, 100, adFldIsNullable
        
        '授權狀態
        rsDefTmp.Fields.Append "AuthStatus", adBSTR, 100, adFldIsNullable
        
        '繳費狀態
        rsDefTmp.Fields.Append "BillStatus", adBSTR, 100, adFldIsNullable
        
        '訂購期間
        rsDefTmp.Fields.Append "AuthPeriodDate", adBSTR, 100, adFldIsNullable
        
        '授權更新日期
        rsDefTmp.Fields.Append "UpdAuthDate", adBSTR, 30, adFldIsNullable
        
        'InsertToOracle
        
        rsDefTmp.Open
        
        rsSourceData.MoveFirst
        Do While Not rsSourceData.EOF
            rsDefTmp.AddNew
            For i = 0 To rsDefTmp.Fields.Count - 1
                Select Case UCase(rsDefTmp.Fields(i).Name)
                    Case "CHOICE"
                        rsDefTmp(i).Value = "0"
                    Case "INITPLACENO"
                        
                    Case "INITPLACENAME"
                        rsDefTmp.Fields(i).Value = GetInitPlaceName(rsSourceData("STBSEQNO") & "", rsSourceData("CUSTID") & "")
                    Case "ORDERSTATUS"
                        rsDefTmp.Fields(i).Value = GetOrderStatus(rsSourceData)
                    Case UCase("AuthStatus")
                        rsDefTmp.Fields(i).Value = GetAuthStatus(rsSourceData)
                    Case UCase("BillStatus")
                        rsDefTmp.Fields(i).Value = GetBillStatus(rsSourceData)
                    Case UCase("AuthPeriodDate")
                        rsDefTmp.Fields(i).Value = rsSourceData("AuthStartTime") & " ~ " & rsSourceData("AuthStopTime")
                    Case UCase("UpdAuthDate")
                        rsDefTmp.Fields(i).Value = GetUpdAuthDate(rsSourceData("CMDSEQNO") & "")
                    Case Else
                        rsDefTmp.Fields(rsDefTmp.Fields(i).Name) = _
                        rsSourceData.Fields(rsDefTmp.Fields(i).Name).Value & ""
                End Select
            
'                If UCase(rsDefTmp.Fields(i).Name) = "CHOICE" Then
'                    rsDefTmp(i).Value = "0"
'                Else
'                    rsDefTmp.Fields(rsDefTmp.Fields(i).Name) = _
'                        rsSourceData.Fields(rsDefTmp.Fields(i).Name).Value & ""
'                End If
            Next
            rsDefTmp.Update
            rsSourceData.MoveNext
        Loop
    Else
        Set rsDefTmp = rsSourceData
    End If
    
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "DefineRs")
End Sub

Private Sub subGrd()
    On Error Resume Next
    Dim mFlds As New GiGridFlds
    Dim mFlds2 As New GiGridFlds
    Dim mFlds3 As New GiGridFlds
    With mFlds
        If fCanSelect Then
            .Add "Choice", , , , , "選取"
        End If
        '.Add "InitPlaceName", , , , , "裝置地點" & Space(10)
        .Add "OrderStatus", , , , , "訂購狀態"
        .Add "BillStatus", , , , , "繳費狀態"
        .Add "OrderDate", giControlTypeDate, , , , "訂購日期" & Space(15)
        .Add "Title", , , , , "產品名稱" & Space(30)
        .Add "AuthPeriodDate", , , , , "授權期間" & Space(60)
        .Add "Amount", , , , , "金額" & Space(10)
        .Add "OrderWayCode", , , , , "訂購管道" & Space(5)
        .Add "OrderName", , , , , "訂購人員" & Space(10)
        .Add "ProductType", , , , , "產品種類" & Space(5)
        .Add "RENTALDURATION", , , , , "收視期間"
        .Add "AuthStatus", , , , , "授權狀態"
        .Add "UpdAuthDate", , , , , "授權更新日期" & Space(10)
    End With
    With mFlds2
        .Add "InitPlaceNo", , , , , "裝置地點" & Space(20)
        .Add "PRDate", , , , , "設備" & Space(20)
        .Add "FaciSNo", , , , , "STB" & Space(20), vbLeftJustify
        .Add "STBSNO", , , , , "HDD S/N" & Space(20), vbLeftJustify
        .Add "DVRSizeCode", , , , , "HDD 容量" & Space(10), vbLeftJustify
    End With
    
    With mFlds3
        .Add "ShortTitle", , , , , "簡短片名" & Space(10)
        .Add "Title", , , , , "影片名稱" & Space(25)
        .Add "Actors", , , , , "主演(演員)" & Space(10)
        .Add "Synopsis", , , , , "節目簡介" & Space(30)
        .Add "Directors", , , , , "導演" & Space(20)
        .Add "Studio", , , , , "出版製片廠" & Space(10)
        .Add "AuthStartTime", , , , , "可看起日" & Space(10)
        .Add "id", , , , , "影片代號" & Space(5)
    End With
    
    ggrData.AllFields = mFlds
    ggrData.SetHead
    ggrData2.AllFields = mFlds2
    ggrData2.SetHead
    ggrData3.AllFields = mFlds3
    ggrData3.SetHead
    
    ggrData3.Visible = False
    Set ggrData3.Recordset = rsSO563
    ggrData3.Visible = True
    
    'ggrData.Visible = False
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ChkErr
    Dim aMsg As String
    Dim aRs As ADODB.Recordset
    Screen.MousePointer = vbHourglass
    If Not CanCancelOK(aMsg, aRs) Then
        MsgBox aMsg, vbInformation, "訊息"
        GoTo 88
    End If
    Set frmSO11A1B.uRs = aRs
    frmSO11A1B.Show 1, Me
    
88:
    On Error Resume Next
    CloseRecordset aRs
    Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdCancel_Click")
End Sub
Private Function CanRetryOK(ByRef aMsg As String, ByRef aRs As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    Dim blnRet As Boolean
    Dim aQry As String
    blnRet = False
    aMsg = Empty
    Set aRs = rsDefTmp.Clone
    aRs.Filter = "Choice = 1"
    If aRs.EOF Or aRs.BOF Then
        aMsg = "無選定任何資料！"
        CanRetryOK = False
        Exit Function
    End If
    aRs.MoveFirst
    
    Do While Not aRs.EOF
'        If Len(aRs("AuthStopTime") & "") = 0 Then
'            aMsg = aMsg & "產品代號：[ " & aRs("ProductId") & " ] 無法重送！" & vbCrLf
'            GoTo 88
'        Else
'            If Not IsDate(aRs("AuthStopTime")) Then
'                aMsg = aMsg & "產品代號：[ " & aRs("ProductId") & " ] 無法重送！" & vbCrLf
'                GoTo 88
'            Else
'                If CDate(aRs("AuthStopTime")) > RightNow Then
'                    aMsg = aMsg & "產品代號：[ " & aRs("ProductId") & " ] 無法重送！" & vbCrLf
'                    GoTo 88
'                End If
'            End If
'        End If
        
        If Val(aRs("CancelFlag") & "") <> 0 Then
            aMsg = aMsg & "產品代號：[ " & aRs("ProductId") & " ] 為取消狀態，無法重送！" & vbCrLf
            GoTo 88
        End If

        '#6033 測試不OK,原本是AUTHSTOPTIME<RightNow現在改成>=RightNow By Kin 2011/07/22
        aQry = "SELECT COUNT(1) FROM SO033PVOD2 WHERE AUTHSTOPTIME >= TO_DATE('" & _
                    Format(RightNow, "YYYYMMDDHHMMSS") & "','YYYYMMDDHH24MISS') " & _
                    " AND SEQNO = " & aRs("SEQNO")
        If Val(GetRsValue(aQry, cnCom) & "") = 0 Then
            aMsg = aMsg & "產品代號：[ " & aRs("ProductId") & " ] 無法重送！" & vbCrLf
            GoTo 88
        End If
        
        aQry = " SELECT COUNT(1) CNT FROM " & GetOwner & "SO004 WHERE SEQNO = '" & aRs("STBSEQNO") & "' " & _
                " AND CUSTID = " & aRs("CUSTID") & " AND COMPCODE = " & gCompCode & _
                " AND CMOPENDATE IS NOT NULL AND CMOPENDATE > NVL(CMCloseDate,To_Date('10000101','yyyymmdd')) " & _
                " AND Prdate is null and GetDate is null " & _
                " UNION ALL " & _
                " SELECT COUNT(1) CNT FROM " & GetOwner & "SO004 WHERE STBSNO = '" & aRs("STBSEQNO") & "' " & _
                " AND CUSTID = " & aRs("CUSTID") & " AND COMPCODE = " & gCompCode & _
                " AND CMOPENDATE IS NOT NULL AND CMOPENDATE > NVL(CMCloseDate,To_Date('10000101','yyyymmdd')) " & _
                " AND Prdate is null and GetDate is null "
        aQry = "SELECT Sum(CNT) Cnt FROM ( " & aQry & ")"
        If Val(GetRsValue(aQry, gcnGi) & "") <> 2 Then
            aMsg = aMsg & "產品代號：[ " & aRs("ProductId") & " ] 無法重送！" & vbCrLf
        End If
        
88:
        aRs.MoveNext
    Loop

    blnRet = aMsg = Empty

  CanRetryOK = blnRet
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "CanRetryOK")
End Function
Private Function CanCancelOK(ByRef aMsg As String, ByRef aRs As ADODB.Recordset) As Boolean
  On Error GoTo ChkErr
    Dim rsClone As ADODB.Recordset
    Dim blnRet As Boolean
    Dim aQry As String
    blnRet = False
    aMsg = Empty
    Set rsClone = rsDefTmp.Clone
    rsClone.Filter = "Choice = 1"
    If rsClone.EOF Or rsClone.BOF Then
        'MsgBox "無選定任何資料！", vbInformation, "訊息"
        aMsg = "無選定任何資料！"
        GoTo 88
    End If
    rsClone.MoveFirst
    Do While Not rsClone.EOF
        If (Len(rsClone("ReturnCode") & "") > 0) Or (Val(rsClone("CancelFlag") & "") <> 0) Then
            aMsg = aMsg & "產品代號：[ " & rsClone("ProductId") & " ] 無法取消！" & vbCrLf
        Else
            aQry = "SELECT COUNT(1) FROM SO033PVOD2 WHERE AUTHSTOPTIME >= TO_DATE('" & _
                    Format(RightNow, "YYYYMMDDHHMMSS") & "','YYYYMMDDHH24MISS') " & _
                    " AND SEQNO = " & rsClone("SEQNO")
            
            If Val(GetRsValue(aQry, cnCom) & "") = 0 Then
                aMsg = aMsg & "產品代號：[ " & rsClone("ProductId") & " ] 無法取消！" & vbCrLf
            End If
            
'            If Len(rsClone("AuthStopTime") & "") = 0 Then
'                aMsg = aMsg & "產品代號：[ " & rsClone("ProductId") & " ] 無法取消！" & vbCrLf
'            Else
'                If Not IsDate(rsClone("AuthStopTime")) Then
'                    aMsg = aMsg & "產品代號：[ " & rsClone("ProductId") & " ] 無法取消！" & vbCrLf
'                Else
'                    If CDate(rsClone("AuthStopTime")) < RightNow Then
'                        aMsg = aMsg & "產品代號：[ " & rsClone("ProductId") & " ] 無法取消！" & vbCrLf
'                    End If
'                End If
'            End If
        End If
        rsClone.MoveNext
    Loop
    Set aRs = rsClone
    blnRet = aMsg = Empty
88:
    On Error Resume Next
    'CloseRecordset rsClone
    CanCancelOK = blnRet
  Exit Function
ChkErr:
 Call ErrSub(Me.Name, "CanCancelOK")
End Function
Private Sub cmdExit_Click()
  Unload Me
End Sub



Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    
    Dim blnRet As Boolean
    Dim rsClone As ADODB.Recordset
    Dim lngErrCnt As Long
    Dim lngRightCnt As Long
    Dim strErrMsg As String
    Dim strRetCode As String
    Dim objWrite As New Scripting.FileSystemObject
    Dim objTxt As TextStream
    Dim aPath As String
    Dim aFileName As String
    Dim aMsg As String
    Dim aCasId As Variant
    
    Dim i As Long
    If Len(fWebServiceUrl & "") = 0 Then
        MsgBox "連結網站尚未設定！請先設定連結網站", vbCritical, "訊息"
        GoTo 88
    End If
    aPath = gErrLogPath
    aFileName = "OrdServer_Err.Txt"
    If Right(aPath, 1) <> "\" Then
        aPath = aPath & "\"
    End If
    Set objTxt = objWrite.CreateTextFile(aPath & aFileName, True)
    'Set rsClone = rsDefTmp.Clone
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
'    rsClone.Filter = "Choice = 1"
'    If rsClone.EOF Or rsClone.BOF Then
'        MsgBox "無選定任何資料！", vbInformation, "訊息"
'        GoTo 88
'    End If
    If Not CanRetryOK(aMsg, rsClone) Then
        MsgBox aMsg, vbInformation, "訊息"
        GoTo 88
    End If
    rsClone.MoveFirst
    Do While Not rsClone.EOF
        If Val(rsClone("Choice") & "") = 1 Then
            strErrMsg = Empty
            strRetCode = Empty
            aCasId = Split(GetCasidS(rsClone("SEQNO"), Greater), ",")
            For i = LBound(aCasId) To UBound(aCasId)
                blnRet = GenPrmAuthReSend(fWebServiceUrl, CStr(gCompCode), rsClone("STBSeqNo") & "", _
                        rsSO004("FaciSNO") & "", rsClone("CustId") & "", aCasId(i) & "", _
                        rsClone("OrderNo") & "", strRetCode, strErrMsg, 4)
                If Not blnRet Then
                    aMsg = "授權碼：" & aCasId(i) & " , " & _
                        "STB設備流水號：" & rsClone("STBSeqNo") & " , " & _
                        "流水號：" & rsClone("SeqNo") & " , " & _
                        "Return Code：" & strRetCode & " , " & _
                        "Error ：" & strErrMsg
                    WriteTextLine objTxt, aMsg
                    lngErrCnt = lngErrCnt + 1
                
                    Exit For
                End If
            Next i
            
            If blnRet Then
                lngRightCnt = lngRightCnt + 1
            End If
        End If
        rsClone.MoveNext
        DoEvents
    Loop
    If lngRightCnt > 0 Or lngErrCnt > 0 Then
        MsgBox "重送成功: " & lngRightCnt & " 筆" & vbCrLf & _
                "重送失敗: " & lngErrCnt & " 筆", vbInformation, "訊息"
        If lngErrCnt > 0 Then
            Shell "Notepad " & aPath & aFileName, vbNormalFocus
        End If
    Else
        MsgBox "無任何資料可傳送！", vbInformation, "訊息"
    End If
88:
  On Error Resume Next
  CloseRecordset rsClone
  Set objTxt = Nothing
  Set objWrite = Nothing
  Me.Enabled = True
  Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdOK_Click")
End Sub
Private Function GetCasidS(ByVal aSEQNO As String, ByVal aDateSymbol As CompareDate) As String
  On Error GoTo ChkErr
    Dim rsCasid As New ADODB.Recordset
    Dim aSQL As String
    Dim aRet As String
    Dim aSymbolStr As String
    If aDateSymbol = Greater Then
        aSymbolStr = ">="
    Else
        aSymbolStr = "<="
    End If
    
    aSQL = "SELECT CASID FROM SO033PVOD2 WHERE SEQNO = " & aSEQNO & _
            " AND AUTHSTOPTIME " & aSymbolStr & " TO_DATE('" & _
                    Format(RightNow, "YYYYMMDDHHMMSS") & "','YYYYMMDDHH24MISS') "
    If Not GetRS(rsCasid, aSQL, cnCom, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    If Not rsCasid.EOF Then
        aRet = rsCasid.GetString(, , , ",")
        If Right(aRet, 1) = "," Then
            aRet = Left(aRet, Len(aRet) - 1)
        End If
    End If
    GetCasidS = aRet
    On Error Resume Next
    Call CloseRecordset(rsCasid)
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "GetCasidS")
End Function


Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    If KeyCode = vbKeyF2 Then
        If cmdOK.Visible Then
            cmdOK.Value = True
        End If
    End If
    If KeyCode = vbKeyF3 Then
        If cmdCancel.Visible And cmdCancel.Enabled Then
            cmdCancel.Value = True
        End If
    End If
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGrd
    Call DefineRs
    fWebServiceUrl = GetWebUrl
    cmdOK.Visible = fCanSelect
    
   
    
    
    ggrData.Visible = False
    Set ggrData.Recordset = rsDefTmp
    ggrData.Visible = True
    
    
   
    ggrData2.Visible = False
    Set ggrData2.Recordset = rsSO004
    ggrData2.Visible = True
    
    
    
    
    
    If ggrData.Recordset.RecordCount > 0 Then
        ggrData.Recordset.MoveFirst
    End If
    ggrData.Refresh
    ggrData2.Refresh
    ggrData3.Refresh
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub
Private Function GetWebUrl() As String
  On Error GoTo ChkErr
    Dim aSQL As String
    aSQL = "SELECT PVODOrderURL FROM " & GetOwner & "CD039 WHERE CODENO = " & gCompCode
    GetWebUrl = GetRsValue(aSQL, gcnGi) & ""
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "GetWebUrl")
End Function
Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    CloseRecordset rsDefTmp
    CloseRecordset rsSourceData
End Sub

Private Sub ggrData_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
'    If Len(ggrData.Recordset.Fields("ReturnName") & "") > 0 Then
'        Value = vbRed
'    End If
    If Not fCanSelect Then Exit Sub
    If UCase(giFld.FieldName) = "CHOICE" Then
        If Value = 1 Then Value = vbRed
    End If
    
End Sub

Private Sub ggrData_DblClick()
  On Error GoTo ChkErr
    If rsDefTmp.RecordCount = 0 Then
        Exit Sub
    End If
    If fCanSelect Then
        If Val(rsDefTmp("Choice") & "") = 1 Then
            rsDefTmp("Choice") = "0"
        Else
            rsDefTmp("Choice") = "1"
        End If
        rsDefTmp.Update
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "ggrData_DblClick")
End Sub



Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If fCanSelect Then
        If UCase(giFld.FieldName) = UCase("Choice") Then
            If Value = "1" Then
                Value = "V"
            Else
                Value = ""
            End If
        End If
    End If
    '原本是直接抓COM區CD114的資料，改為抓各SO資料 By Kin 2011/04/08
    If UCase(giFld.FieldName) = UCase("ORDERWAYCODE") Then
        Dim aValue As String
        aValue = GetRsValue("SELECT Description FROM " & GetOwner & "CD114 " & _
                            " WHERE STOPFLAG <> 1 AND CODENO = " & Value, gcnGi) & ""
        Value = aValue
    End If
    
    If UCase(giFld.FieldName) = UCase("BillStatus") Then
        If Val(ggrData.Recordset("ProductType") & "") = 0 Then
            If Val(ggrData.Recordset("Amount") & "") = 0 Then
                Value = "贈送"
            End If
        End If
    End If
    
    '1=single為PPV,2=multiple為PPP,0=subscription為SBP,Free
    If UCase(giFld.FieldName) = UCase("ProductType") Then
        Select Case Val(Value & "")
                
            Case 1
                Value = "PPV"
            Case 2
                Value = "PPP"
            Case Else
                If Val(ggrData.Recordset("OrderKind") & "") = 0 Then
                    Value = "包月"
                Else
                    Value = "Free"
                End If
                If Val(ggrData.Recordset("Amount") & "") = 0 Then
                    Value = "Free"
                End If
                'Value = "Subscription"
        End Select
        
    End If
    If UCase(giFld.FieldName) = UCase("Amount") Then
        If Val(ggrData.Recordset("ProductType") & "") = 0 Then
            Value = "0"
        End If
    End If
    
    If UCase(giFld.FieldName) = UCase("Validity_Flag") Then
        Select Case Val(Value & "")
            Case 1
                Value = "絕對時間"
            Case Else
                Value = "相對時間"
        End Select
        
    End If
    If UCase(giFld.FieldName) = UCase("RENTALDURATION") Then
        Dim aDate As Double
        aDate = ((Value / 60) / 60) / 24
        Value = RoundCS(aDate, 1)
    End If
    If UCase(giFld.FieldName) = UCase("OrderKind") Then
        If Val(Value & "") = 1 Then
            Value = "贈送"
        Else
            Value = "訂購"
        End If
        
    End If
    Exit Sub
'    If UCase(giFld.FieldName) = UCase("UpdAuthDate") Then
'        Value = GetUpdAuthDate(ggrData.Recordset("CMDSeqNo") & "")
'
'    End If
    
    
    
'    If UCase(giFld.FieldName) = UCase("ORDERSTATUS") Then
'        Value = GetOrderStatus(ggrData.Recordset)
'    End If
ChkErr:
  
  Call ErrSub(Me.Name, "ggrData_ShowCellData")
End Sub

Private Sub ggrData2_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    Select Case UCase(giFld.FieldName)
        Case UCase("InitPlaceNo")
            Value = GetInitPlaceName(ggrData2.Recordset("SEQNO") & "", ggrData2.Recordset("CustId") & "")
        Case UCase("PrDate")
            Value = GetFaciStatus(ggrData2.Recordset)
        Case UCase("STBSNO")
            If Len(ggrData2.Recordset("FaciSNO") & "") > 0 Then
                Value = GetDVRSno(ggrData2.Recordset("CustId"), ggrData2.Recordset("FaciSNo") & "", , , ggrData2.Recordset("SEQNO") & "")
            End If
        Case UCase("DVRSizeCode")
            Value = GetDVRHDD(ggrData2.Recordset("CustId"), ggrData2.Recordset("FaciSNo") & "", , ggrData2.Recordset("SEQNO") & "", True) & "(GB)"
    End Select
    
'    If UCase(giFld.FieldName) = "STBSNO" Then
'        If Len(strSTBFaciSNO) > 0 Then
'            Value = GetDVRSno(lngCustId, strSTBFaciSNO, , , strSTBSeqNo)
'            strDvrShowCellTmp = Value
'        End If
'    End If
'    If UCase(giFld.FieldName) = UCase("DVRSizeCode") Then '正式硬碟容量
'        Value = GetDVRHDD(lngCustId, strSTBFaciSNO, , strSTBSeqNo, True) & "(GB)"
'    End If
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "ggrData2_ShowCellData")
End Sub

Private Sub rsDefTmp_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  On Error GoTo ChkErr
    If rsDefTmp.EOF Then Exit Sub
    If Not ggrData3.Recordset Is Nothing Then
        If Len(rsDefTmp("SEQNO") & "") > 0 Then
            ggrData3.Recordset.Filter = "SEQNO = " & rsDefTmp("SEQNO")
            ggrData3.Refresh
        End If
    End If
    
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "rsDefTmp_MoveComplete")
End Sub

Private Sub rsSO004_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  On Error GoTo ChkErr
    If rsSO004.EOF Then Exit Sub
    ggrData.Recordset.Filter = "STBSEQNO = '" & rsSO004("SEQNO") & "'"
    rsDefTmp.Filter = "STBSEQNO = '" & rsSO004("SEQNO") & "'"
    Set ggrData.Recordset = rsDefTmp
    ggrData.Refresh
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "rsSO004_MoveComplete")
End Sub

