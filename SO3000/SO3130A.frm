VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO3130A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "����վ� [SO3130A]"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3130A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin prjGiList.GiList gilNewElctEn 
      Height          =   315
      Left            =   1530
      TabIndex        =   7
      Top             =   1650
      Visible         =   0   'False
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F2Corresponding =   -1  'True
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   5025
      Left            =   330
      TabIndex        =   6
      Top             =   2100
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8864
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   9990
      TabIndex        =   2
      Top             =   240
      Width           =   1185
   End
   Begin VB.Frame fraData 
      Height          =   6885
      Left            =   90
      TabIndex        =   3
      Top             =   480
      Width           =   9435
      Begin VB.CommandButton cmdOk 
         Caption         =   "�}�l"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Left            =   4320
         TabIndex        =   1
         Top             =   720
         Width           =   825
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "���"
         Height          =   375
         Left            =   4320
         TabIndex        =   0
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblNewElctEn 
         AutoSize        =   -1  'True
         Caption         =   "�s���O�H��"
         Height          =   195
         Left            =   300
         TabIndex        =   8
         Top             =   1260
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblStep1 
         AutoSize        =   -1  'True
         Caption         =   "�B�J1: ��ܭ즬�O���P��D���X����R���"
         Height          =   195
         Left            =   300
         TabIndex        =   5
         Top             =   315
         Width           =   3690
      End
      Begin VB.Label lblStep2 
         AutoSize        =   -1  'True
         Caption         =   "�B�J2: �վ㦬�O����, �A���T�w��}�l����"
         Height          =   195
         Left            =   300
         TabIndex        =   4
         Top             =   780
         Width           =   3585
      End
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   930
      TabIndex        =   9
      Top             =   60
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth1       =   800
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   5430
      TabIndex        =   10
      Top             =   75
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth1       =   800
      F2Corresponding =   -1  'True
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "���q�O"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   150
      Width           =   585
   End
   Begin VB.Label lblServiceType 
      AutoSize        =   -1  'True
      Caption         =   "�A�����O"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4530
      TabIndex        =   11
      Top             =   150
      Width           =   780
   End
End
Attribute VB_Name = "frmSO3130A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���\��i�N�s��������ƼȦs��(So032)����ơA������վ�A��l��,Grid���ť�
Option Explicit
Private rsOrcl As New ADODB.Recordset
'Private ObjopenDb As New giOpenDBObj.OpenDBObj
Private cnnMdb As New ADODB.Connection
Private rsMdb As New ADODB.Recordset
Private blnChkrs As Boolean
Private blnUpdate As Boolean

Private Sub cmdCancel_Click()
    On Error GoTo chkErr
        Unload Me
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo chkErr
'�E  �ˬdgrid���O�_�����, �Y�L, �h�T��: "�L���", �õL�H�U�ʧ@
'�E  �ˬd�C�@�����s���O�H���O�_����, �Y�L�h�T��: "����쥲�ݦ���", ��Ц^������, �õL�H�U�ʧ@
'�E  Loop grid���C�@�����, ������B�z, SQL�p�U:
' "update SO032 set ClctEn=<�s���O�H���N��> where StrtCode=<��D�N�X> and ClctEn=<�즬�O�H���N��>"
'�E  ��:<�즬�O�H��>���i�ରnull��
Dim strSQL As String
Dim intCount As Long
Dim intTotalCount As Long
Dim rsTmp As New ADODB.Recordset
    'ggddata.SaveNow
    If MsgBox("�T�w�}�l����?", vbQuestion + vbYesNo + vbDefaultButton2, "����") = vbNo Then Exit Sub
    If blnChkrs <> True Then MsgNoRcd: Exit Sub
    Screen.MousePointer = vbHourglass
    If Not rsMdb.EOF Then
        rsMdb.MoveFirst
        gcnGi.BeginTrans
        Set rsTmp = rsMdb.Clone
        If rsTmp.RecordCount > 0 Then
        End If
        Do While Not rsTmp.EOF
            With rsTmp
                If IsNull(.Fields("NewElctEn").Value) Then GoTo lblNext
                strSQL = "Update " & GetOwner & "So032 Set ClctEn='" & .Fields("NewElctEn").Value & "',ClctName='" & .Fields("NewEmpName").Value & "'," & _
                            "oldClctEn='" & .Fields("NewElctEn").Value & "',OldClctName='" & .Fields("NewEmpName").Value
                strSQL = strSQL & "' Where StrtCode='" & .Fields("StrtCode").Value & "' And ClctEn" & IIf((.Fields("ClctEn").Value & "") = "", " Is Null", "='" & .Fields("ClctEn").Value & "'")
                gcnGi.Execute strSQL, intCount
                intTotalCount = intTotalCount + intCount
            End With
lblNext:
                rsTmp.MoveNext
        Loop
'        picmsg.Visible = False
        MsgBox "�@���� " & intTotalCount & " ��", vbInformation, "�T���I"
        gcnGi.CommitTrans
        Screen.MousePointer = vbDefault
        Call cmdQuery_Click
        'Unload Me
    Else
        MsgBox "�L��ơI�Ч�s��ܫ�A���ʡI", vbExclamation, "�T���I"
    End If
    Screen.MousePointer = vbDefault
Exit Sub
chkErr:
    Screen.MousePointer = vbDefault
    MsgBox "������ѡI�Э��s�ާ@�I", vbExclamation, "�T���I"
    gcnGi.RollbackTrans
End Sub

Private Sub cmdQuery_Click()
On Error GoTo chkErr
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
        'ggddata.SaveNow
        '���M���ثe�s�u
        ggrData.Visible = False
        Call subChoose
        strSQL = "Select StrtCode,StrtName,ClctEn,D.EmpName,intCount From " & _
                " (Select A.StrtCode,B.Description StrtName,A.ClctEn,A.intCount From " & _
                        "(Select StrtCode,ClctEn,Count(*) as intCount From " & GetOwner & "So032 " & strChoose & _
                            " Group By StrtCode,ClctEn) A," & _
                            GetOwner & "CD017 B Where A.StrtCode=B.CodeNo) C," & GetOwner & "CM003 D " & _
             " Where C.ClctEn=D.EmpNo(+) Order by ClctEn,StrtCode"
        
        If rsOrcl.State = adStateOpen Then rsOrcl.Close
'        rsOrcl.CursorLocation = adUseClient
'        rsOrcl.Open strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly
        If Not GetRS(rsOrcl, strSQL, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly) Then Exit Sub
        If rsOrcl.EOF Then ggrData.Visible = True: MsgNoRcd
        Call SaveRecToTmp
        
        Set ggrData.Recordset = rsMdb
        ggrData.Refresh
        ggrData.Blank = False
        ggrData.Visible = True
        If Not rsMdb.EOF Then
            rsMdb.MoveFirst
            ggrData.Enabled = True
            cmdQuery.Enabled = False
            cmdOk.Enabled = True
            ggrData.SetFocus
        End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdQuery_Click")
End Sub

Private Sub subChoose()
    On Error Resume Next
        strChoose = ""
        If gilCompCode.GetCodeNo <> "" Then Call subAnd2(strChoose, "CompCode =" & gilCompCode.GetCodeNo)
        If gilServiceType.GetCodeNo <> "" Then Call subAnd2(strChoose, "ServiceType ='" & gilServiceType.GetCodeNo & "'")
        If strChoose <> "" Then strChoose = " Where " & strChoose
End Sub


Private Sub SaveRecToTmp()
    On Error GoTo chkErr
    Dim strSQL As String
        '�O����Rs
        cnnMdb.BeginTrans
        DefaultRs
        Do While Not rsOrcl.EOF
            With rsMdb
             .AddNew
             .Fields("ClctEn") = NoZero(rsOrcl("ClctEn"))
             .Fields("EmpName") = NoZero(rsOrcl("EmpName"))
             .Fields("StrtCode") = NoZero(rsOrcl("StrtCode"))
             .Fields("StrtName") = NoZero(rsOrcl("StrtName"))
             .Fields("Count") = NoZero(rsOrcl("intCount"), True)
             .Update
            End With
            rsOrcl.MoveNext
        Loop
        cnnMdb.CommitTrans
        If Not GetRS(rsMdb, "Select * from tmpSo3130", cnnMdb, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        blnChkrs = True
    Exit Sub
chkErr:
    cnnMdb.RollbackTrans
    Call ErrSub(Me.Name, "SaveRecToTmp")
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Activate")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo chkErr
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyEscape Then Call cmdCancel_Click: Exit Sub
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
    On Error GoTo chkErr
        '���o�@�ΰѼ�
        If Alfa2 = True Then Call GetGlobal
        Call subGil
        Call InitData
        Call subGrd
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub InitData()
    On Error Resume Next
    Dim strPath As String
    Dim strConn As String
        blnChkrs = False
        strPath = ReadGICMIS1("TmpMDBPath")
        'cnnMdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Password=;Data Source=" & strPath & "\Tmp1111.MDB" & ";Persist Security Info=True"
        cnnMdb.Open "Provider=" & GetOleDbProvider & ";Password=;Data Source=" & strPath & "\Tmp1111.MDB" & ";Persist Security Info=True"
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
End Sub

Private Sub ClearS()
    On Error Resume Next
        cmdQuery.Enabled = True
        cmdOk.Enabled = False
        ggrData.Blank = True
End Sub

Private Sub subGil()
    On Error Resume Next
        SetgiList gilNewElctEn, "EmpNo", "EmpName", "CM003", , , , , 1, 12, , True
        SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)

End Sub

Private Sub subGrd()
    On Error GoTo chkErr
    Dim mFlds As New prjGiGridR.GiGridFlds
        '(�즬�O��(�N��/�m�W),��D,�i��,�s���O��(�N��/�m�W))
        With mFlds
            .Add "ClctEn", , , , False, "���O�H���N��", vbLeftJustify
            .Add "EmpName", , , , False, "���u�m�W", vbLeftJustify
            .Add "StrtCode", , , , False, "��D�N�X", vbRightJustify
            .Add "StrtName", , , , False, "     ��D�W��    ", vbLeftJustify
            .Add "Count", , , , False, "�i��", vbRightJustify
            .Add "NewElctEn", , , , True, "�s���O�H���N��", vbLeftJustify, , , , "EmpNO", "EmpName", "CM003", 10, 20
            .Add "NewEmpName", , , , True, "�s���O�H���m�W", vbLeftJustify, , , , "EmpNO", "EmpName", "CM003", 10, 20
        End With
        ggrData.AllFields = mFlds
        ggrData.SetHead
        ggrData.Enabled = False
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGrd")
End Sub

Private Function DefaultRs() As Boolean
    On Error GoTo chkErr
    Dim rsTmp As New ADODB.Recordset
        If Not GetRS(rsTmp, "Select A.ClctEn,A.ClctName EmpName,A.StrtCode,B.Description StrtName,Count(*) Count," & _
            "A.ClctEn NewElctEn,A.ClctName NewEmpName From SO032 A,CD014 B Where A.StrtCode=B.CodeNo And " & _
            "A.BillNo = '' And Item=-1 Group By A.ClctEn,A.ClctName,A.StrtCode,B.Description") Then Exit Function
        If Not CreateMDBTable(rsTmp, "tmpSo3130", cnnMdb) Then Exit Function
          
        If Not GetRS(rsMdb, "Select * From tmpSo3130 Where 1=0", cnnMdb, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        Call CloseRecordset(rsTmp)
    Exit Function
chkErr:
    ErrSub Me.Name, "DefaultRS"
End Function
 
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        cnnMdb.Execute "Delete From CM003"
        Set rsOrcl = Nothing
        Set rsMdb = Nothing
        Set cnnMdb = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)
End Sub

Private Sub ggrData_DblClick()
    On Error GoTo chkErr
        If ggrData.Recordset.BOF Or ggrData.Recordset.EOF Then Exit Sub
        gilNewElctEn.Visible = True
        lblNewElctEn.Visible = True
        gilNewElctEn.SetFocus
    Exit Sub
chkErr:
    ErrSub Me.Name, "ggrData_DblClick"
End Sub


Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3100", "SO3130") Then Exit Sub
        Call subGil
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.SetCodeNo ""
        gilServiceType.Query_Description
        gilServiceType.ListIndex = 1
        Call GiListFilter(gilNewElctEn, , gilCompCode.GetCodeNo)
        ClearS
End Sub

Private Sub gilNewElctEn_GotFocus()
    On Error GoTo chkErr
        With ggrData.Recordset
                If .BOF Or .EOF Then Exit Sub
                gilNewElctEn.SetCodeNo .Fields("NewElctEn") & ""
                gilNewElctEn.SetDescription .Fields("NewEmpName") & ""
        End With
    Exit Sub
chkErr:
    ErrSub Me.Name, "gilNewElctEn_GotFocus"
End Sub

Private Sub gilNewElctEn_Validate(Cancel As Boolean)
    On Error GoTo chkErr
        With rsMdb
                If .Fields("ClctEn") = gilNewElctEn.GetCodeNo Then
                    MsgBox "�s���O�H�����o�P�즬�O�H���ۦP", vbExclamation, "ĵ�i"
                    gilNewElctEn.GotoCodeNo
                    Cancel = True
                    Exit Sub
                Else
                    .Fields("NewElctEn") = NoZero(gilNewElctEn.GetCodeNo2)
                    .Fields("NewEmpName") = NoZero(gilNewElctEn.GetDescription2)
                    .Update
                End If
        End With
        Set ggrData.Recordset = rsMdb
        ggrData.Refresh
        lblNewElctEn.Visible = False
        gilNewElctEn.Visible = False
    Exit Sub
chkErr:
    ErrSub Me.Name, "gilNewElctEn_Validate"
End Sub
