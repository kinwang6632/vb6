VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1120C 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "SO1120C.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所屬視窗中央
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   4035
      Left            =   90
      TabIndex        =   5
      Top             =   570
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   7117
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
   Begin VB.CommandButton cmdFind 
      Caption         =   "F3. 尋找"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      TabIndex        =   1
      Top             =   90
      Width           =   1095
   End
   Begin VB.TextBox txtNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1965
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "確 定 (&Y)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Picture         =   "SO1120C.frx":0442
      TabIndex        =   2
      Top             =   4680
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取 消 (&N)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2850
      TabIndex        =   3
      Top             =   4680
      Width           =   1500
   End
   Begin VB.Label lblNO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmSO1120C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Dim intType As Integer '1.IP 2.CP Number, 3.ReFaciSNo
Dim strCircuitNO As String
Dim objParentForm As Object
Dim strNo As String
Dim strReturnValue As String
Dim lngPortNo As Long
Dim strIPAddress As String
Dim strREFACISNO As String
Dim strSeqNo As String
Dim strMeSeqNo As String
Dim blnCheckSingle As Boolean

Private Sub cmdCancel_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdCancel_Click"
End Sub

Private Sub cmdFind_Click()
    On Error GoTo ChkErr
        Call OpenData(False)
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdFind_Click"
End Sub

Private Sub cmdOk_Click()
    On Error Resume Next
    Dim strTable As String, strFld As String
    Dim lngRcdAft As Long
    If rs.EOF Then Exit Sub
    Me.Enabled = False
    Select Case intType
        Case 1
            strTable = "SO048"
            strFld = "IPAddress"
        Case 2
            strTable = "SO049"
            strFld = "CPNumber"
    End Select
    'gcnGi.Execute "Update " & GetOwner & strTable & " Set UseFlag=0 Where " & strFld & "='" & strNo & "' And CompCode =" & gCompCode
'    If frmSO1623A.strIPaddr <> Empty Then
'        gcnGi.Execute "Update " & GetOwner & strTable & _
'                                " Set UseFlag=0" & _
'                                " Where " & strFld & "='" & frmSO1623A.strIPaddr & "'" & _
'                                " And CompCode =" & gCompCode & _
'                                " And UseFlag=1", lngRcdAft
'    End If
    With frmSO1623A
        If .strIPaddr <> Empty And .blnGetIP Then
            gcnGi.Execute "Update " & GetOwner & "SO048" & _
                                    " Set UseFlag=0" & _
                                    " Where IPAddress = '" & .strIPaddr & "'" & _
                                    " And CompCode =" & garyGi(9)
        End If
    End With
    gcnGi.Execute "Update " & GetOwner & strTable & _
                            " Set UseFlag=1" & _
                            " Where " & strFld & "='" & rs(0).Value & "'" & _
                            " And CompCode =" & gCompCode & _
                            " And UseFlag=0", lngRcdAft
    If lngRcdAft <= 0 Then
        strReturnValue = ""
        MsgBox "該筆 IP Address 在您選定前,已被其他人給取用, 請重新選擇 !", vbInformation, "訊息"
        OpenData False
        Me.Enabled = True
        Exit Sub
    End If
    strReturnValue = rs(0).Value & ""
'        If intType = 3 Then
'            strIPAddress = rs("IPAddress") & ""
'            objParentForm.txtDeclarantName = rs("ContName") & ""
'            objParentForm.txtDomicileAddress2 = rs("DomicileAddress") & ""
'            strSeqNo = rs("SeqNo") & ""
'            'lngPortNo = rs("intCount") + 1
'        End If
        Unload Me
End Sub

Private Sub dgdData_DblClick()
    On Error Resume Next
        Call cmdOk_Click
End Sub

Private Sub Form_Activate()
    If intType = 3 And blnCheckSingle Then
        Me.Visible = False
        If rs.RecordCount = 1 Then
            cmdOk.Value = True
        End If
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Select Case KeyCode
           Case vbKeyEscape
                Unload Me
           Case vbKeyF3
                cmdFind.Value = True
    End Select
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
    On Error GoTo ChkErr
        Call InitData
        Call subGrd
        Call OpenData(True)
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Sub InitData()
    On Error Resume Next
    Dim lngHeight As Long
        Select Case intType
            Case 1
                Me.Caption = "IP 位址"
                lblNO.Caption = "網路編號"
                txtNo = strNo
            Case 2
                Me.Caption = "CP電話號碼"
                lblNO.Caption = "CP電話號碼"
            Case 3
                Me.Caption = "EMTA序號"
                lblNO.Visible = False
                txtNo.Visible = False
                cmdFind.Visible = False
                lngHeight = cmdFind.Height + 60
                ggrData.Top = ggrData.Top - lngHeight
                cmdOk.Top = cmdOk.Top - lngHeight
                cmdCancel.Top = cmdCancel.Top - lngHeight
                Me.Height = Me.Height - lngHeight
        End Select
        strReturnValue = ""
End Sub

Private Sub OpenData(blnLoad As Boolean)
    On Error GoTo ChkErr
        
        Dim strSQL As String
        Dim strWhere As String
        Dim intCPMax As Integer
        Dim strGateway As String
    
        Set rs = New ADODB.Recordset
        
        Select Case intType
            Case 1
                strWhere = "CircuitNo like '" & txtNo & "%' And "
                strSQL = "Select IPAddress,UseFlag,CircuitNo From " & GetOwner & "SO048" & _
                                " Where " & strWhere & " CompCode=" & gCompCode & " And UseFlag=0" & _
                                " And Upper(IPNature)='CM' " & _
                                " Order By IPAddress"
            Case 2
'                '加過濾GetWay 93/03/24
'                If strREFACISNO <> "" Then strGateway = GetRsValue("Select Gateway From " & GetOwner & "SO004 Where CustId=" & gCustId & " And ReFaciSNo ='" & strREFACISNO & "'") & ""
'                If Not blnLoad And txtNo <> "" Then strWhere = " CPNumber like '" & txtNo & "%' And "
'                If strGateway <> "" Then strWhere = strWhere & " Gateway = '" & strGateway & "' And "
'
'                strSQL = "Select CPNumber,UseFlag From " & GetOwner & "SO049 Where " & strWhere & " CompCode=" & gCompCode & " and UseFlag=0 and CPStatus is Null Order By CPNumber"
'            Case 3
'                intCPMax = GetRsValue("Select CPMax From " & GetOwner & "CD022 Where CodeNo =" & objParentForm.gilFaciCode.GetCodeNo)
'                strSQL = "Select FaciSNo,ReFaciSNo,CustId From " & GetOwner & "SO004 Where CustId = " & gCustId & " And FaciCode In (Select CodeNo From " & GetOwner & "CD022 Where RefNo = 6) And PRDate is null And SeqNo <> '" & strMeSeqNo & "'"
'
'                strSQL = "Select A.FaciSNo,B.Description,A.IPAddress,Count(C.ReFaciSNo) intCount,A.ContName,A.ContName2,A.DomicileAddress,A.SeqNo From " & GetOwner & "SO004 A,(" & _
'                strSQL & ") C," & GetOwner & "CD056 B Where A.FaciSNo=C.ReFaciSNo(+) And A.CustId=C.CustId(+) And A.InitPlaceNo=B.CodeNo(+) And A.CustId = " & _
'                gCustId & " And A.FaciCode In (Select CodeNo From " & GetOwner & "CD022 Where Refno= 5)  And A.PRDate Is Null Group by A.FaciSNo,B.Description,A.IPAddress,A.ContName,A.ContName2,A.DomicileAddress,A.SeqNo having Count(C.ReFaciSNo)<" & intCPMax
        End Select
        
        If Not GetRS(rs, strSQL) Then Exit Sub
        If rs.RecordCount > 0 Then cmdOk.Enabled = True
        Set ggrData.Recordset = rs
        ggrData.Refresh
    
    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub

Private Sub subGrd()
    On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
        With ggrData
            Select Case intType
                Case 1
                    mFlds.Add "CircuitNo", , , , , "網路編號"
                    mFlds.Add "IPAddress", , , , , "IP 位址" & Space(10)
                    mFlds.Add "UseFlag", , , , , "使用狀態"
                Case 2
                    mFlds.Add "CPNumber", , , , , "CP 電話號碼" & Space(10)
                    mFlds.Add "UseFlag", , , , , "使用狀態"
                Case 3
                    mFlds.Add "FaciSNo", , , , , Me.Caption & Space(10)
                    mFlds.Add "ContName2", , , , , "聯絡人" & Space(10)
                    mFlds.Add "Description", , , , , "裝置點" & Space(10)
                    mFlds.Add "IPAddress", , , , , "IP 位址"
                    mFlds.Add "intCount", , , , , "使用號碼數"
            End Select
            .AllFields = mFlds
            .SetHead
        End With
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subGrd"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        intType = 0
        strCircuitNO = ""
        objParentForm.uReturnValue = strReturnValue
        objParentForm.uPortNo = Val(lngPortNo)
        objParentForm.uIPAddress = strIPAddress
        'strReturnValue = ""
        strIPAddress = ""
        lngPortNo = 0
        Call CloseRecordset(rs)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ChkErr
        ReleaseCOM frmSO1120C
    Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Unload"
End Sub

Public Property Let uType(ByVal vData As Integer)
    intType = vData
End Property

Public Property Set uParentForm(ByVal vData As Object)
    Set objParentForm = vData
End Property

Public Property Let uOldNo(ByVal vData As String)
    strNo = vData
End Property

Public Property Let uReFaciSNo(ByVal vData As String)
    strREFACISNO = vData
End Property

Public Property Let uCheckSingle(ByVal vData As Boolean)
    blnCheckSingle = vData
End Property

Public Property Get uSeqNo() As String
    uSeqNo = strSeqNo
End Property

Public Property Let uMeSeqNo(vData As String)
    strMeSeqNo = vData
End Property

Public Property Get uReturnValue() As String
    uReturnValue = strReturnValue
End Property

Private Sub ggrData_DblClick()
    Call cmdOk_Click
End Sub
