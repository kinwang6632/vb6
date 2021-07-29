VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1131E 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�]�Ʃ��� [SO1131E]"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO1131E.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   3045
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   5371
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSO1131E"
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
Dim strDefFaciSeqNo As String
Dim blnShowClctDate As Boolean
Dim str004Filter As String
Dim blnInstRsCopy As Boolean
Dim strCitem As String
Dim blnCMBaudRate As Boolean
Dim blnPPVInfo As Boolean
Dim blnPRCanClick As Boolean
Dim strServiceType As String
'2016/09/02 Jacky 7285 Jacy
Dim strFilter As String

'2016/09/02 Jacky 7285 Jacy
Public Property Let uFilter(ByVal vData As String)
    strFilter = vData
End Property
Public Property Let uServiceType(ByVal vData As String)
    strServiceType = vData
End Property
Public Property Let uDefFaciSeqNo(ByVal vData As String)
    strDefFaciSeqNo = vData
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

Public Property Get uFaciSeqNo() As String
    uFaciSeqNo = strFaciSeqNo
End Property

Public Property Get uSmartCardNo() As String
    uSmartCardNo = strSmartCardNo
End Property

Public Property Let uCustid(ByVal vData As Long)
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

Public Property Let uInstRsCopy(ByVal vData As Boolean)
    blnInstRsCopy = vData
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
        Select Case KeyCode
            Case vbKeyEscape
                Unload Me
        End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
        '#3593 ServiceType�ھڶǨӪ��ѼƬ���,�S�Ǯɤ~�ϥ�gServiceType By Kin 2007/12/14
        If strServiceType = "" Then
            strServiceType = gServiceType
        End If
        If lngCustId = 0 Then
            lngCustId = gCustId
        End If
        Call InitData
        Call subGrd
        Call OpenData
        Call DefaultValue
End Sub

Private Sub InitData()
    On Error Resume Next
        Me.Move objParentForm.Left, objParentForm.Top + objParentForm.Height - Me.Height
        strFaciSeqNo = ""
        strFaciSNo = ""
End Sub

Private Sub OpenData()
    On Error GoTo ChkErr
    Dim strWhere As String
    Dim strFaciWhere As String
    Dim strSQL As String
        strFaciWhere = " A.FaciCode In (Select CodeNo From " & GetOwner & "CD022 Where "
        
        '#3593 ServiceType�ھڶǨӪ��ѼƬ���,�S�Ǯɤ~�ϥ�gServiceType By Kin 2007/12/14
        Select Case strServiceType
            Case "I"
                '2007/02/01 Jacky 2697 HandSome
                '�Щ�I�A�ȭ쬰�L�o2,5�W�[���L�o2,5,7,8
                strWhere = " And " & strFaciWhere & " RefNo In (2,5,7,8" & IIf(strRefNo = "", "", "," & strRefNo)
            Case "D"
                strWhere = " And " & strFaciWhere & " RefNo In (3" & IIf(strRefNo = "", "", "," & strRefNo)
            Case "P"
                strWhere = " And " & strFaciWhere & " RefNo In (6" & IIf(strRefNo = "", "", "," & strRefNo)
            Case Else
                If strRefNo <> "" Then strWhere = " And " & strFaciWhere & " RefNo In (" & strRefNo
        End Select
        '2016/09/02 Jacky 7285 Jacy
        If strFilter = "" Then
            If strWhere <> "" Then strWhere = strWhere & "))"
        Else
            strWhere = " And " & strFilter
        End If
'        strSQL = "Select '' as SelectStatus,'' as Status,A.SeqNo,a.FaciSNo,A.FaciCode, a.FaciName, a.InstDate, a.PRDate, b. Description InitPlace, a.ModelName, a.BuyName, a.SmartCardNo,A.CustId,A.CompCode,A.ServiceType From " & _
'            GetOwner & "SO004 A, " & GetOwner & "CD056 B Where A.InitPlaceNo=B.CodeNo(+) And A.CustId=" & gCustId & _
'            " And A.ServiceType='" & gServiceType & "' and a.CompCode=" & gCompCode & strWhere

        '#3593 ServiceType�ھڶǨӪ��ѼƬ���,�S�Ǯɤ~�ϥ�gServiceType By Kin 2007/12/14
        strSQL = "Select '' as SelectStatus,A.*, b. Description InitPlace,Nvl(C.RefNo,0) RefNo From " & _
            GetOwner & "SO004 A, " & GetOwner & "CD056 B," & GetOwner & "CD022 C Where A.InitPlaceNo=B.CodeNo(+) And A.FaciCode=C.CodeNo And A.CustId=" & lngCustId & _
            " And A.ServiceType='" & strServiceType & "' and a.CompCode=" & gCompCode & strWhere & " And Nvl(C.RefNo,0)<> 4"
        
        If str004Filter <> "" Then strSQL = strSQL & " And " & str004Filter
        
        If blnShowClctDate Then strSQL = "Select A.*,B.ClctDate,B.CitemName From (" & strSQL & ") A," & GetOwner & "SO003 B Where A.CompCode= B.CompCode(+) And A.ServiceType= B.ServiceType(+) And A.CustId = B.CustId(+) And A.SeqNo = B.FaciSeqNo(+) And B.CitemCode(+) = " & strCitem
        If blnPPVInfo Then strSQL = "Select A.*,B.FuncFlag01 From (" & strSQL & ") A," & GetOwner & "CD043 B Where A.ModelCode= B.CodeNo(+)"
        
        If Not GetRS(rsData, strSQL, , adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        
        Set rsData.ActiveConnection = Nothing
        If blnInst Then Call InsertInstRs
        Set ggrData.Recordset = rsData
        ggrData.Refresh
    Exit Sub
ChkErr:
    ErrSub Me.Name, "OpenData"
End Sub

Private Sub DefaultValue()
    On Error Resume Next
        If rsData.RecordCount = 0 Then Exit Sub
        rsData.MoveFirst
        rsData.Find "SeqNo = '" & strDefFaciSeqNo & "'"
        If Not rsData.EOF Then
            strFaciSNo = rsData("FaciSNo") & ""
            strFaciSeqNo = rsData("SeqNo") & ""
        End If
        If rsData.EOF Then rsData.MoveFirst
        
End Sub

Private Sub InsertInstRs()
    On Error GoTo ChkErr
    Dim rsClone As New ADODB.Recordset
    Dim strRefFaciCode As String
    Dim intLoop As Integer
    Dim intRefNo As Integer
        If ChkRecordsetIsOk(rsInstRS) = False Then
            Exit Sub
        End If
        Set rsClone = rsInstRS.Clone
        On Error Resume Next
        If strRefNo <> "" Then strRefFaciCode = "(" & gcnGi.Execute("Select CodeNo From " & GetOwner & "CD022 Where RefNo In (" & strRefNo & ")").GetString(, , "", "),(")
        On Error GoTo ChkErr
        If rsClone.RecordCount > 0 Then rsClone.MoveFirst
        Do While Not rsClone.EOF
            With rsData
                .AddNew
                If blnInstRsCopy Then
                    For intLoop = 0 To .Fields.Count - 1
                        If .Fields(intLoop).Name <> "SELECTSTATUS" And .Fields(intLoop).Name <> "STATUS" And .Fields(intLoop).Name <> "INITPLACE" And .Fields(intLoop).Name <> "REFNO" Then
                            On Error Resume Next
                            .Fields(intLoop) = rsClone(.Fields(intLoop).Name)
                        End If
                    Next
                    If rsClone("InitPlaceNo") & "" <> "" Then .Fields("InitPlace") = GetRsValue("Select Description From " & GetOwner & "CD056 Where CodeNo = " & rsClone("InitPlaceNo"))
                Else
                    .Fields("SeqNo") = rsClone("FaciSeqNo")
                    .Fields("FaciSNo") = Null
                    .Fields("FaciCode") = rsClone("FaciNo")
                    .Fields("FaciName") = rsClone("FaciName")
                    .Fields("InstDate") = Null
                    .Fields("PRDate") = Null
                    .Fields("InitPlace") = Null
                    .Fields("ModelName") = Null
                    '.Fields("Status") = Null
                End If
                .Fields("BuyName") = rsClone("BuyName")
                .Fields("SmartCardNo") = Null
                If InStr(1, strRefFaciCode, "(" & .Fields("FaciCode") & ")") > 0 Or strRefNo = "" Then
                    intRefNo = GetRsValue("Select Nvl(RefNo,0) From " & GetOwner & "CD022 Where CodeNo = " & .Fields("FaciCode"))
                    If intRefNo <> 4 Then .Update
                End If
                .CancelUpdate
            End With
            rsClone.MoveNext
        Loop
        
        Call CloseRecordset(rsClone)
    Exit Sub
ChkErr:
    ErrSub Me.Name, "InsertInstRS"
End Sub

Private Sub subGrd()
    On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
      '���~�Ǹ�, ���ئW��, �˾����, ������, �˸m�I, �����W��, �R��覡, ���z�d�Ǹ�
      
        mFlds.Add "SelectStatus", , , , , "�� "
        mFlds.Add "PRDate", , , , , "���A   "
        mFlds.Add "FaciSNo", , , , , "���~�Ǹ�                  "
        mFlds.Add "DeclarantName", , , , , "�ӽФH�m�W"
        mFlds.Add "FaciCode", , , , , "�N�X"
        mFlds.Add "FaciName", , , , , "���ئW��          "
        
        '#3593 ServiceType�ھڶǨӪ��ѼƬ���,�S�Ǯɤ~�ϥ�gServiceType By Kin 2007/12/14
        If strServiceType = "I" Or blnCMBaudRate Then
            mFlds.Add "CMBaudNo", , , , , "�N�X"
            mFlds.Add "CMBaudRate", , , , , "CM�t�v"
            mFlds.Add "FixIPCount", , , , , "�T�wIP"
            mFlds.Add "DynIPCount", , , , , "�ʺAIP"
        End If
        mFlds.Add "InstDate", giControlTypeDate, , , , "  �˾����  "
        mFlds.Add "PRDate", giControlTypeDate, , , , "  ������  "
        mFlds.Add "InitPlace", , , , , "�˸m�I   "
        mFlds.Add "ModelName", , , , , "�����W��"
        'mflds.Add "BuyCode", , , , , "���O�覡"
        mFlds.Add "BuyName", , , , , "�R��覡"
        mFlds.Add "SmartCardNo", , , , , "���z�d�Ǹ�                 "
        If blnPPVInfo Then
            mFlds.Add "ReFaciSNo", , , , , "�����Ǹ�"
            mFlds.Add "HintName", , , , , "�K�X���ܦW��"
            mFlds.Add "ParentPassword", , , , , "���D����"
            mFlds.Add "DueDate", giControlTypeDate, , , , "���ɨ����"
            mFlds.Add "ContractCust", , , , , "�j����"
            mFlds.Add "ContractDate", giControlTypeDate, , , , "�j�������"
            mFlds.Add "FuncFlag01", , , , , "�^�ǼҲ�"
        End If
        mFlds.Add "SeqNo", , , , , "�]�Ƭy����     "
        If blnShowClctDate Then
            mFlds.Add "CitemName", , , , , "���O����    "
            mFlds.Add "ClctDate", giControlTypeDate, , , , "   ������   "
        End If
        '�B�^�ǼҲ�(cd042.FuncFlag01)�B
        ggrData.AllFields = mFlds
        ggrData.UseCellForeColor = True
        ggrData.SetHead
    Exit Sub
ChkErr:
    ErrSub Me.Name, "subGrd"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
'        blnInst = False
'        blnInstRsCopy = False
'        strRefNo = ""
'        lngEditMode = 0
'        str004Filter = ""
'        strDefFaciSeqNo = ""
'        strCitem = ""
'        blnShowClctDate = False
        lngCustId = 0
        strServiceType = ""
        Call CloseRecordset(rsData)
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        ReleaseCOM frmSO1131E
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
        strFaciSNo = rsData("FaciSNo") & ""
        strFaciSeqNo = rsData("SeqNo") & ""
        strSmartCardNo = rsData("SmartCardNo") & ""
        Unload Me
End Sub

Private Function IsDataOk() As Boolean
    On Error GoTo ChkErr
        If rsData.RecordCount > 0 Then
            If rsData("PRDate") & "" <> "" Then
                If blnPRCanClick Then
                    If MsgBox("�ӳ]�Ƥw�,�нT�w�O�_���??", vbQuestion + vbYesNo, gimsgPrompt) = vbNo Then Exit Function
                Else
                    MsgBox "����]�Ƥ�����!!", vbExclamation, gimsgPrompt: Exit Function
                End If
            End If
            If rsData("FaciCode") & "" <> "" Then
                If GetRsValue("Select Nvl(RefNo,0) From " & GetOwner & "CD022 Where CodeNo = " & rsData("FaciCode")) = 4 Then
                    MsgBox "�]�ưѦҸ�4������!!", vbExclamation, gimsgPrompt: Exit Function
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
        If giFld.FieldName = "PRDate" And Trim(giFld.HeadName) = "���A" Then
            Value = GetFaciStatus(ggrData.Recordset, Value)
'        End If

'    If giFld.FieldName = "Status" Then
'        If ggrData.Recordset("PRDate") & "" = "" And ggrData.Recordset("PRSNo") & "" <> "" Then
'            Value = "���"
'        ElseIf ggrData.Recordset("PRDate") & "" = "" And ggrData.Recordset("InstDate") & "" = "" And ggrData.Recordset("ReFaciSNo") & "" <> "" Then
'            Value = "���ˤ�"
'        ElseIf ggrData.Recordset("PRDate") & "" = "" And ggrData.Recordset("InstDate") & "" <> "" Then
'            Value = "���`"
'        ElseIf ggrData.Recordset("PRDate") & "" = "" And ggrData.Recordset("InstDate") & "" = "" Then
'            Value = "�˾���"
'        ElseIf ggrData.Recordset("PRDate") & "" <> "" Then
'            Value = "�"
'        End If
    ElseIf giFld.FieldName = "SelectStatus" Then
        If ggrData.Recordset("SeqNo") & "" = strDefFaciSeqNo Then Value = "V"
    End If
End Sub
