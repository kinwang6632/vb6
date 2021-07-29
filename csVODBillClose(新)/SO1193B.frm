VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1193B 
   BorderStyle     =   1  '單線固定
   Caption         =   "設備明細 [SO1193B]"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO1193B.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   10350
      TabIndex        =   2
      Top             =   3150
      Width           =   1365
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2. 確定"
      Height          =   375
      Left            =   150
      TabIndex        =   1
      Top             =   3180
      Width           =   1365
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   3045
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
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
Attribute VB_Name = "frmSO1193B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngCustId As Long
Dim rsData As New ADODB.Recordset
Dim rsInstRS As New ADODB.Recordset
Dim objParentForm As Form
Dim strFaciSNo As String
Dim strFaciSeqNo As String
Dim strSmartCardNo As String
Dim blnInst As Boolean
Dim strRefNo As String
Dim lngEditMode As giEditModeEnu

Dim blnShowClctDate As Boolean
Dim str004Filter As String

Dim strCitem As String
Dim blnCMBaudRate As Boolean
Dim blnPPVInfo As Boolean
Dim blnPRCanClick As Boolean
Dim strServiceType As String
Dim strMVodId As String
Dim rsDefTmp  As ADODB.Recordset
Public Property Let uServiceType(ByVal vData As String)
    strServiceType = vData
End Property

Public Property Let uEditMode(ByVal vData As giEditModeEnu)
    lngEditMode = vData
End Property
Public Property Let uRefNo(ByVal vData As String)
    strRefNo = vData
End Property
Public Property Get uFaciSNo() As String
    uFaciSNo = strFaciSNo
End Property
Public Property Let uMVodId(ByVal vData As String)
    strMVodId = vData
End Property
Public Property Get uMVodId() As String
    uMVodId = strMVodId
End Property
Public Property Get uFaciSeqNo() As String
    uFaciSeqNo = strFaciSeqNo
End Property
Public Property Let uFaciSeqNo(ByVal vData As String)
        strFaciSeqNo = vData
End Property
Public Property Get uSmartCardNo() As String
    uSmartCardNo = strSmartCardNo
End Property

Public Property Let uCustId(ByVal vData As Long)
    lngCustId = vData
End Property

Public Property Let uParentForm(ByVal vData As Form)
    Set objParentForm = vData
End Property

Public Property Let uShowClctDate(ByVal vData As Boolean)
    blnShowClctDate = vData
End Property

Public Property Let uCitemCode(ByVal vData As String)
    strCitem = vData
End Property

Public Property Let uInst(ByVal vData As Boolean)
    blnInst = vData
End Property

Public Property Set uInstRS(ByVal vData As ADODB.Recordset)
    Set rsInstRS = vData
End Property
Public Property Let u004Filter(ByVal vData As String)
    str004Filter = vData
End Property

Public Property Let uPPVInfo(ByVal vData As Boolean)
    blnPPVInfo = vData
End Property

Public Property Let uPRCanClick(ByVal vData As Boolean)
    blnPRCanClick = vData
End Property

Private Sub cmdCancel_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub cmdOk_Click()
  On Error GoTo ChkErr
    Dim strVodAccountId As String
    Dim rs004G As New ADODB.Recordset
    strMVodId = Empty
    strFaciSeqNo = Empty
    ggrData.Visible = False
    ggrData.Recordset.MoveFirst
    Do While Not ggrData.Recordset.EOF
        If Val(ggrData.Recordset("SelectStatus") & "") = 1 Then
            If Len(strFaciSeqNo) > 0 Then
                strFaciSeqNo = strFaciSeqNo & ",'" & ggrData.Recordset("SeqNo") & "'"
            Else
                strFaciSeqNo = "'" & ggrData.Recordset("SeqNo") & "'"
            End If
            If Len(ggrData.Recordset("VodAccountId") & "") > 0 Then
                If Len(strVodAccountId) > 0 Then
                    strVodAccountId = strVodAccountId & "," & ggrData.Recordset("VodAccountId")
                Else
                    strVodAccountId = ggrData.Recordset("VodAccountId")
                End If
            End If
        End If
        ggrData.Recordset.MoveNext
    Loop
    If Len(strVodAccountId) > 0 Then
        If Not GetRS(rs004G, "select distinct MvodId from " & GetOwner & "SO004G where VodAccountId In (" & strVodAccountId & ")", _
                    gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        Do While Not rs004G.EOF
            If Len(rs004G("MvodId") & "") > 0 Then
                If Len(strMVodId) > 0 Then
                    strMVodId = strMVodId & "," & rs004G("MvodId") & ""
                Else
                    strMVodId = rs004G("MvodId") & ""
                End If
            End If
            rs004G.MoveNext
        Loop
    End If
    On Error Resume Next
    Call CloseRecordset(rs004G)
    Unload Me
  Exit Sub
ChkErr:
  ErrSub Me.Name, "cmdOk_Click"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        Select Case KeyCode
            Case vbKeyEscape
                Unload Me
            Case vbKeyF2
                cmdOk.Value = True
        End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
        Set rsDefTmp = New ADODB.Recordset
        If strServiceType = "" Then
            strServiceType = gServiceType
        End If
        If lngCustId = 0 Then
            lngCustId = gCustId
        End If
        Call InitData
        Call subGrd
        Call OpenData
        'Call DefaultValue
End Sub

Private Sub InitData()
    On Error Resume Next
        Me.Move objParentForm.Left, objParentForm.Top + objParentForm.Height - Me.Height
        'strFaciSeqNo = ""
        'strFaciSNo = ""
End Sub
Private Sub SetSelectValue()
  On Error GoTo ChkErr
    Dim strFaciSeqNoCopy As String
    strFaciSeqNoCopy = strFaciSeqNo
    strFaciSeqNoCopy = Replace(strFaciSeqNoCopy, "'", "")
    ggrData.Visible = False
    If Len(strFaciSeqNoCopy) > 0 Then
        If ggrData.Recordset.RecordCount > 0 Then
            ggrData.Recordset.MoveFirst
            Do While Not ggrData.Recordset.EOF
                If InStr(1, strFaciSeqNoCopy, ggrData.Recordset("SeqNo")) > 0 Then
                    ggrData.Recordset("SelectStatus") = "1"
                    ggrData.Recordset.Update
                End If
                ggrData.Recordset.MoveNext
            Loop
            ggrData.Recordset.MoveFirst
        End If
    End If
    ggrData.Visible = True
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "SetSelectValue")
End Sub
Private Sub OpenData()
    On Error GoTo ChkErr
    Dim strWhere As String
    Dim strFaciWhere As String
    Dim strSQL As String
        strFaciWhere = " A.FaciCode In (Select CodeNo From " & GetOwner & "CD022 Where "
        Select Case strServiceType
            Case "I"
                strWhere = " And " & strFaciWhere & " RefNo In (2,5,7,8" & IIf(strRefNo = "", "", "," & strRefNo)
            Case "D"
                strWhere = " And " & strFaciWhere & " RefNo In (3" & IIf(strRefNo = "", "", "," & strRefNo)
            Case "P"
                strWhere = " And " & strFaciWhere & " RefNo In (6" & IIf(strRefNo = "", "", "," & strRefNo)
            Case Else
                If strRefNo <> "" Then strWhere = " And " & strFaciWhere & " RefNo In (" & strRefNo
        End Select
        
        If strWhere <> "" Then strWhere = strWhere & "))"
        strSQL = "Select '' as SelectStatus,A.*, b. Description InitPlace,Nvl(C.RefNo,0) RefNo From " & _
            GetOwner & "SO004 A, " & GetOwner & "CD056 B," & GetOwner & "CD022 C Where A.InitPlaceNo=B.CodeNo(+) And A.FaciCode=C.CodeNo And A.CustId=" & lngCustId & _
            " And A.ServiceType='" & strServiceType & "' and a.CompCode=" & gCompCode & strWhere & " And Nvl(C.RefNo,0)<> 4"
        
        If str004Filter <> "" Then strSQL = strSQL & " And " & str004Filter
        
        If blnShowClctDate Then strSQL = "Select A.*,B.ClctDate,B.CitemName From (" & strSQL & ") A," & GetOwner & "SO003 B Where A.CompCode= B.CompCode(+) And A.ServiceType= B.ServiceType(+) And A.CustId = B.CustId(+) And A.SeqNo = B.FaciSeqNo(+) And B.CitemCode(+) = " & strCitem
        If blnPPVInfo Then strSQL = "Select A.*,B.FuncFlag01 From (" & strSQL & ") A," & GetOwner & "CD043 B Where A.ModelCode= B.CodeNo(+)"
        
        If Not GetRS(rsData, strSQL, , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        
        Set rsData.ActiveConnection = Nothing
        DefineRs
        'If blnInst Then Call InsertInstRs
        Set ggrData.Recordset = rsDefTmp
        ggrData.Recordset.MoveFirst
        SetSelectValue
        ggrData.Refresh
    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub
Private Sub DefineRs()
  On Error GoTo ChkErr
    Dim rsDef138 As New ADODB.Recordset
    Dim i As Integer
    Dim j As Integer
    With rsDefTmp
        If .State = adStateOpen Then
            .Close
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
    End With
    For i = 0 To rsData.Fields.Count - 1
        With rsDefTmp
            .Fields.Append rsData.Fields(i).Name, adBSTR, 300, adFldIsNullable
            
'            If UCase(rsData.Fields(i).Name) = "STOPDATE" Then
'                .Fields.Append "StopDate", adBSTR, 20
'            Else
'                If UCase(rsData.Fields(i).Name) = UCase("SelectStatus") Then
'
'                End If
'                .Fields.Append rsData.Fields(i).Name, adBSTR, rsData.Fields(i).DefinedSize, adFldIsNullable
'            End If
        End With
    Next i
    rsDefTmp.Open
    rsData.MoveFirst
    Do While Not rsData.EOF
        rsDefTmp.AddNew
        For i = 0 To rsData.Fields.Count - 1
            rsDefTmp(rsData.Fields(i).Name) = rsData.Fields(i).Value
        Next
        rsDefTmp.Update
        rsData.MoveNext
    Loop
    
    On Error Resume Next
    Exit Sub
ChkErr:
  ErrSub Me.Name, "DefineRs"
End Sub

'Private Sub DefaultValue()
'    On Error Resume Next
'        If rsData.RecordCount = 0 Then Exit Sub
'        rsData.MoveFirst
'        rsData.Find "SeqNo = '" & strDefFaciSeqNo & "'"
'        If Not rsData.EOF Then
'            strFaciSNo = rsData("FaciSNo") & ""
'            strFaciSeqNo = rsData("SeqNo") & ""
'        End If
'        If rsData.EOF Then rsData.MoveFirst
'
'End Sub

'Private Sub InsertInstRs()
'    On Error GoTo ChkErr
'    Dim rsClone As New ADODB.Recordset
'    Dim strRefFaciCode As String
'    Dim intLoop As Integer
'    Dim intRefNo As Integer
'        Set rsClone = rsInstRS.Clone
'        On Error Resume Next
'        If strRefNo <> "" Then strRefFaciCode = "(" & gcnGi.Execute("Select CodeNo From " & GetOwner & "CD022 Where RefNo In (" & strRefNo & ")").GetString(, , "", "),(")
'        On Error GoTo ChkErr
'        If rsClone.RecordCount > 0 Then rsClone.MoveFirst
'        Do While Not rsClone.EOF
'            With rsData
'                .AddNew
'                If blnInstRsCopy Then
'                    For intLoop = 0 To .Fields.Count - 1
'                        If .Fields(intLoop).Name <> "SELECTSTATUS" And .Fields(intLoop).Name <> "STATUS" And .Fields(intLoop).Name <> "INITPLACE" And .Fields(intLoop).Name <> "REFNO" Then
'                            On Error Resume Next
'                            .Fields(intLoop) = rsClone(.Fields(intLoop).Name)
'                        End If
'                    Next
'                    If rsClone("InitPlaceNo") & "" <> "" Then .Fields("InitPlace") = GetRsValue("Select Description From " & GetOwner & "CD056 Where CodeNo = " & rsClone("InitPlaceNo"))
'                Else
'                    .Fields("SeqNo") = rsClone("FaciSeqNo")
'                    .Fields("FaciSNo") = Null
'                    .Fields("FaciCode") = rsClone("FaciNo")
'                    .Fields("FaciName") = rsClone("FaciName")
'                    .Fields("InstDate") = Null
'                    .Fields("PRDate") = Null
'                    .Fields("InitPlace") = Null
'                    .Fields("ModelName") = Null
'                    '.Fields("Status") = Null
'                End If
'                .Fields("BuyName") = rsClone("BuyName")
'                .Fields("SmartCardNo") = Null
'                If InStr(1, strRefFaciCode, "(" & .Fields("FaciCode") & ")") > 0 Or strRefNo = "" Then
'                    intRefNo = GetRsValue("Select Nvl(RefNo,0) From " & GetOwner & "CD022 Where CodeNo = " & .Fields("FaciCode"))
'                    If intRefNo <> 4 Then .Update
'                End If
'                .CancelUpdate
'            End With
'            rsClone.MoveNext
'        Loop
'
'        Call CloseRecordset(rsClone)
'    Exit Sub
'ChkErr:
'    ErrSub Me.Name, "InsertInstRS"
'End Sub

Private Sub subGrd()
    On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
      '物品序號, 項目名稱, 裝機日期, 拆機日期, 裝置點, 型號名稱, 買賣方式, 智慧卡序號
      
        mFlds.Add "SelectStatus", , , , , "選 "
        mFlds.Add "PRDate", , , , , "狀態   "
        mFlds.Add "FaciSNo", , , , , "物品序號                 "
        mFlds.Add "DeclarantName", , , , , "申請人姓名"
        mFlds.Add "FaciCode", , , , , "代碼"
        mFlds.Add "FaciName", , , , , "項目名稱          "
        
        '#3593 ServiceType根據傳來的參數為準,沒傳時才使用gServiceType By Kin 2007/12/14
        If strServiceType = "I" Or blnCMBaudRate Then
            mFlds.Add "CMBaudNo", , , , , "代碼"
            mFlds.Add "CMBaudRate", , , , , "CM速率"
            mFlds.Add "FixIPCount", , , , , "固定IP"
            mFlds.Add "DynIPCount", , , , , "動態IP"
        End If
        mFlds.Add "InstDate", giControlTypeDate, , , , "  裝機日期  "
        mFlds.Add "PRDate", giControlTypeDate, , , , "  拆機日期  "
        mFlds.Add "InitPlace", , , , , "裝置點   "
        mFlds.Add "ModelName", , , , , "型號名稱"
        'mflds.Add "BuyCode", , , , , "收費方式"
        mFlds.Add "BuyName", , , , , "買賣方式"
        mFlds.Add "SmartCardNo", , , , , "智慧卡序號        "
        If blnPPVInfo Then
            mFlds.Add "ReFaciSNo", , , , , "母機序號"
            mFlds.Add "HintName", , , , , "密碼提示名稱"
            mFlds.Add "ParentPassword", , , , , "問題答案"
            mFlds.Add "DueDate", giControlTypeDate, , , , "租借到期日"
            mFlds.Add "ContractCust", , , , , "綁約戶"
            mFlds.Add "ContractDate", giControlTypeDate, , , , "綁約到期日"
            mFlds.Add "FuncFlag01", , , , , "回傳模組"
        End If
        mFlds.Add "SeqNo", , , , , "設備流水號     "
        If blnShowClctDate Then
            mFlds.Add "CitemName", , , , , "收費項目    "
            mFlds.Add "ClctDate", giControlTypeDate, , , , "   次收日   "
        End If
        '、回傳模組(cd042.FuncFlag01)、
        ggrData.AllFields = mFlds
        ggrData.UseCellForeColor = True
        ggrData.SetHead
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subGrd"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        lngCustId = 0
        strServiceType = ""
        Call CloseRecordset(rsData)
        Call CloseRecordset(rsDefTmp)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        ReleaseCOM frmSO1193B
End Sub

Private Sub ggrData_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
        If giFld.FieldName = "SelectStatus" Then
            If Value <> "" Then Value = vbRed
        End If
End Sub

Private Sub ggrData_DblClick()
    On Error Resume Next
        If lngEditMode = giEditModeView Then Exit Sub
        If Not IsDataOk Then Exit Sub
        ggrData.Visible = False
        If Val(ggrData.Recordset("SelectStatus") & "") = 0 Then
            ggrData.Recordset("SelectStatus") = "1"
        Else
            ggrData.Recordset("SelectStatus") = "0"
        End If
        ggrData.Recordset.Update
        ggrData.Visible = True
End Sub

Private Function IsDataOk() As Boolean
    On Error GoTo ChkErr
        If rsDefTmp.RecordCount > 0 Then
            If rsDefTmp("PRDate") & "" <> "" Then
                If blnPRCanClick Then
                    If MsgBox("該設備已拆除,請確定是否選取??", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then Exit Function
                Else
                    MsgBox "拆除的設備不能選擇!!", vbExclamation, gimsgPrompt: Exit Function
                End If
            End If
            If rsDefTmp("FaciCode") & "" <> "" Then
                If GetRsValue("Select Nvl(RefNo,0) From " & GetOwner & "CD022 Where CodeNo = " & rsDefTmp("FaciCode")) = 4 Then
                    MsgBox "設備參考號4不能選擇!!", vbExclamation, gimsgPrompt: Exit Function
                End If
            End If
        End If
        IsDataOk = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "IsDataOk"
End Function

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error Resume Next
        If giFld.FieldName = "PRDate" And Trim(giFld.HeadName) = "狀態" Then
            Value = GetFaciStatus(ggrData.Recordset, Value)
        ElseIf giFld.FieldName = "SelectStatus" Then
            If Val(ggrData.Recordset("SelectStatus") & "") = 1 Then
                Value = "V"
            Else
                Value = ""
            End If
            'If ggrData.Recordset("SeqNo") & "" = strDefFaciSeqNo Then Value = "V"
        End If
End Sub

