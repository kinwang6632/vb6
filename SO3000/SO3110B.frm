VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSO3110B 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "���X��@���� [SO3110B]"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   11235
   Icon            =   "SO3110B.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   11235
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame Frame1 
      Height          =   6195
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   11025
      Begin TabDlg.SSTab sstData 
         Height          =   5715
         Left            =   150
         TabIndex        =   4
         Top             =   300
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   10081
         _Version        =   393216
         Tabs            =   4
         Tab             =   1
         TabsPerRow      =   4
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "�ݦ��O"
         TabPicture(0)   =   "SO3110B.frx":0442
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "ggrView1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "�w���O"
         TabPicture(1)   =   "SO3110B.frx":045E
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "ggrView2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "�b�b"
         TabPicture(2)   =   "SO3110B.frx":047A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "ggrView3"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "���`���"
         TabPicture(3)   =   "SO3110B.frx":0496
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "ggrView4"
         Tab(3).ControlCount=   1
         Begin prjGiGridR.GiGridR ggrView1 
            Height          =   5085
            Left            =   -74820
            TabIndex        =   5
            Top             =   480
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   8969
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin prjGiGridR.GiGridR ggrView2 
            Height          =   5055
            Left            =   180
            TabIndex        =   6
            Top             =   480
            Width           =   10365
            _ExtentX        =   18283
            _ExtentY        =   8916
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin prjGiGridR.GiGridR ggrView3 
            Height          =   4995
            Left            =   -74820
            TabIndex        =   7
            Top             =   480
            Width           =   10365
            _ExtentX        =   18283
            _ExtentY        =   8811
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin prjGiGridR.GiGridR ggrView4 
            Height          =   5085
            Left            =   -74820
            TabIndex        =   8
            Top             =   480
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   8969
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5.�C�L"
      Height          =   435
      Left            =   450
      TabIndex        =   0
      Top             =   6450
      Width           =   1275
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "F2.�ץX"
      Height          =   435
      Left            =   2430
      TabIndex        =   1
      Top             =   6450
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "����(&X)"
      Height          =   435
      Left            =   9270
      TabIndex        =   2
      Top             =   6420
      Width           =   1275
   End
End
Attribute VB_Name = "frmSO3110B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#5169 �ݨD�A�n�N���X�檺��Ƥ�������� By Kin 2009/09/17
Option Explicit
Private rsView1 As New ADODB.Recordset
Private rsView2 As New ADODB.Recordset
Private rsView3 As New ADODB.Recordset
Private rsView4 As New ADODB.Recordset
Private blnExcel As Boolean
Private Sub cmdExit_Click()
 On Error GoTo chkErr
   Unload Me
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdExport_Click()
  On Error GoTo chkErr
    blnExcel = True
    cmdPrint.Value = True
  Exit Sub
chkErr:
  Call ErrSub(Me.Name, "cmdExport_Click")
End Sub

Private Sub cmdPrint_Click()
 On Error GoTo chkErr
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    If cnn.State = adStateClosed Then Set cnn = GetTmpMdbCn
    ReadyGoPrint
    Select Case sstData.Tab
        Case 0
            If rsView1.RecordCount = 0 Then
                MsgBox "�L��ơI", vbInformation, "�T��"
                GoTo 88
            Else
                If Not CreateMDBTable(rsView1, "SO3110A1", cnn) Then Exit Sub
                If Not InsMDBData(rsView1, "SO3110A1") Then Exit Sub
                SendSQL rsView1.Source, True
            End If
        Case 1
            If rsView2.RecordCount = 0 Then
                MsgBox "�L��ơI", vbInformation, "�T��"
                GoTo 88
            Else
                If Not CreateMDBTable(rsView2, "SO3110A1", cnn) Then Exit Sub
                If Not InsMDBData(rsView2, "SO3110A1") Then Exit Sub
                SendSQL rsView2.Source, True
            End If
        Case 2
            If rsView3.RecordCount = 0 Then
                MsgBox "�L��ơI", vbInformation, "�T��"
                GoTo 88
            Else
                If Not CreateMDBTable(rsView3, "SO3110A1", cnn) Then Exit Sub
                If Not InsMDBData(rsView3, "SO3110A1") Then Exit Sub
                SendSQL rsView3.Source, True
            End If
        Case 3
            If rsView4.RecordCount = 0 Then
                MsgBox "�L��ơI", vbInformation, "�T��"
                GoTo 88
            Else
                If Not CreateMDBTable(rsView4, "SO3110A1", cnn) Then Exit Sub
                If Not InsMDBData(rsView4, "SO3110A1") Then Exit Sub
                SendSQL rsView4.Source, True
            End If
    End Select
    If blnExcel Then
        Call toExcel("Select * From SO3110A1")
    Else
        PrintRpt2 GetPrinterName(5), "SO3110A2.Rpt", "SO3110", "���O�楼���ͤ@���� [SO3110A]", "SELECT * From SO3110A1", "������", , True, "Tmp0000.MDB", , , GiPaperLandscape
    End If
88:
    blnExcel = False
    SendSQL , , True
    Me.Enabled = True
    Screen.MousePointer = vbDefault
  Exit Sub
chkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Sub Form_Activate()
 On Error GoTo chkErr
        Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Activate")
End Sub
Private Function InsMDBData(ByRef rs As ADODB.Recordset, ByVal strTable As String) As Boolean
  On Error GoTo chkErr
    Dim rsMdb As New ADODB.Recordset
    Dim i As Long
    If cnn.State = adStateClosed Then Set cnn = GetTmpMdbCn
    If rs.EOF Then InsMDBData = True: Exit Function
    If Not GetRS(rsMdb, "Select * From " & strTable & " Where 1=0", cnn, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    cnn.BeginTrans
    ggrView1.Visible = False
    ggrView2.Visible = False
    ggrView3.Visible = False
    ggrView4.Visible = False
    rs.MoveFirst
    Do While Not rs.EOF
        rsMdb.AddNew
        For i = 0 To rsMdb.Fields.Count - 1
            rsMdb.Fields(rsMdb.Fields(i).Name) = rs.Fields(rsMdb.Fields(i).Name)
        Next i
        rsMdb.Update
        rs.MoveNext
    Loop
    cnn.CommitTrans
    rs.MoveFirst
  On Error Resume Next
  ggrView1.Visible = True
  ggrView2.Visible = True
  ggrView3.Visible = True
  ggrView4.Visible = True
  InsMDBData = True
  Exit Function
chkErr:
  Call ErrSub(Me.Name, "InsMDBData")
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo chkErr
    If KeyCode = vbKeyF2 Then
        cmdExport.SetFocus
        cmdExport.Value = True
        Exit Sub
    End If
    If KeyCode = vbKeyF5 Then
        cmdPrint.SetFocus
        cmdPrint.Value = True
        Exit Sub
    End If
    If KeyCode = vbKeyEscape Then
        cmdExit.SetFocus
        cmdExit.Value = True
        Exit Sub
    End If
    Exit Sub
chkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
  On Error GoTo chkErr
    sstData.Tab = 0
    Call SetGridData1
    Call SetGridData2
    Call SetGridData3
    Call SetGridData4
  Exit Sub
chkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub
Private Sub SetGridData1()
    On Error GoTo chkErr
        Dim mFlds As New prjGiGridR.GiGridFlds
            With mFlds
                .Add "ServiceType", , , , False, "�A�ȧO", vbLeftJustify
                .Add "BillNo", , , , False, "��ڽs��             ", vbLeftJustify
                .Add "Item", , , , False, "����", vbLeftJustify
                .Add "FaciSNo", , , , False, "�]�ƧǸ�       ", vbLeftJustify
                .Add "CitemName", , , , False, "���O���ئW��" & Space(10), vbLeftJustify
                .Add "CustId", , , , False, "�Ȥ�s��", vbLeftJustify
                .Add "MduId", , , , False, "�j�ӽs��", vbLeftJustify
                .Add "MduName", , , , False, "�j�ӦW��" & Space(20), vbLeftJustify
                .Add "ShouldDate", , , , False, "�������", vbLeftJustify
                .Add "ShouldAmt", , , , False, "�������B", vbLeftJustify
                .Add "StartDate", giControlTypeDate, , , False, "�_�l���", vbLeftJustify
                .Add "StopDate", giControlTypeDate, , , False, "�I����", vbLeftJustify
                .Add "Notes", , , , False, "�Ƶ�" & Space(25), vbLeftJustify
            End With
        ggrView1.AllFields = mFlds
        ggrView1.SetHead
        If rsView1.EOF = False Then rsView1.MoveFirst
        Set ggrView1.Recordset = rsView1
        ggrView1.Refresh
    Exit Sub
chkErr:
    ErrSub Me.Name, "SetGridData1"
End Sub
Private Sub SetGridData2()
    On Error GoTo chkErr
        Dim mFlds As New prjGiGridR.GiGridFlds
            With mFlds
                .Add "ServiceType", , , , False, "�A�ȧO", vbLeftJustify
                .Add "BillNo", , , , False, "��ڽs��             ", vbLeftJustify
                .Add "Item", , , , False, "����", vbLeftJustify
                .Add "FaciSNo", , , , False, "�]�ƧǸ�       ", vbLeftJustify
                .Add "CitemName", , , , False, "���O���ئW��" & Space(10), vbLeftJustify
                .Add "CustId", , , , False, "�Ȥ�s��", vbLeftJustify
                .Add "MduId", , , , False, "�j�ӽs��", vbLeftJustify
                .Add "MduName", , , , False, "�j�ӦW��" & Space(20), vbLeftJustify
                .Add "ShouldDate", , , , False, "�������", vbLeftJustify
                .Add "ShouldAmt", , , , False, "�������B", vbLeftJustify
                .Add "StartDate", giControlTypeDate, , , False, "�_�l���", vbLeftJustify
                .Add "StopDate", giControlTypeDate, , , False, "�I����", vbLeftJustify
                .Add "Notes", , , , False, "�Ƶ�" & Space(25), vbLeftJustify
            End With
        ggrView2.AllFields = mFlds
        ggrView2.SetHead
        If rsView2.EOF = False Then rsView2.MoveFirst
        Set ggrView2.Recordset = rsView2
        ggrView2.Refresh
    Exit Sub
chkErr:
    ErrSub Me.Name, "SetGridData2"
End Sub
Private Sub SetGridData3()
    On Error GoTo chkErr
        Dim mFlds As New prjGiGridR.GiGridFlds
            With mFlds
                .Add "ServiceType", , , , False, "�A�ȧO", vbLeftJustify
                .Add "BillNo", , , , False, "��ڽs��             ", vbLeftJustify
                .Add "Item", , , , False, "����", vbLeftJustify
                .Add "FaciSNo", , , , False, "�]�ƧǸ�       ", vbLeftJustify
                .Add "CitemName", , , , False, "���O���ئW��" & Space(10), vbLeftJustify
                .Add "CustId", , , , False, "�Ȥ�s��", vbLeftJustify
                .Add "MduId", , , , False, "�j�ӽs��", vbLeftJustify
                .Add "MduName", , , , False, "�j�ӦW��" & Space(20), vbLeftJustify
                .Add "ShouldDate", , , , False, "�������", vbLeftJustify
                .Add "ShouldAmt", , , , False, "�������B", vbLeftJustify
                .Add "StartDate", giControlTypeDate, , , False, "�_�l���", vbLeftJustify
                .Add "StopDate", giControlTypeDate, , , False, "�I����", vbLeftJustify
                .Add "Notes", , , , False, "�Ƶ�" & Space(25), vbLeftJustify
            End With
        ggrView3.AllFields = mFlds
        ggrView3.SetHead
        If rsView3.EOF = False Then rsView3.MoveFirst
        Set ggrView3.Recordset = rsView3
        ggrView3.Refresh
    Exit Sub
chkErr:
    ErrSub Me.Name, "SetGridData3"
End Sub
Private Sub SetGridData4()
    On Error GoTo chkErr
        Dim mFlds As New prjGiGridR.GiGridFlds
            With mFlds
                .Add "ServiceType", , , , False, "�A�ȧO", vbLeftJustify
                .Add "BillNo", , , , False, "��ڽs��             ", vbLeftJustify
                .Add "Item", , , , False, "����", vbLeftJustify
                .Add "FaciSNo", , , , False, "�]�ƧǸ�       ", vbLeftJustify
                .Add "CitemName", , , , False, "���O���ئW��" & Space(10), vbLeftJustify
                .Add "CustId", , , , False, "�Ȥ�s��", vbLeftJustify
                .Add "MduId", , , , False, "�j�ӽs��", vbLeftJustify
                .Add "MduName", , , , False, "�j�ӦW��" & Space(20), vbLeftJustify
                .Add "ShouldDate", , , , False, "�������", vbLeftJustify
                .Add "ShouldAmt", , , , False, "�������B", vbLeftJustify
                .Add "StartDate", giControlTypeDate, , , False, "�_�l���", vbLeftJustify
                .Add "StopDate", giControlTypeDate, , , False, "�I����", vbLeftJustify
                .Add "Notes", , , , False, "�Ƶ�" & Space(25), vbLeftJustify
            End With
        ggrView4.AllFields = mFlds
        ggrView4.SetHead
        If rsView4.EOF = False Then rsView4.MoveFirst
        Set ggrView4.Recordset = rsView4
        ggrView4.Refresh
    Exit Sub
chkErr:
    ErrSub Me.Name, "SetGridData3"
End Sub
Private Sub OpenData()
  On Error GoTo chkErr
  If cnn.State = adStateClosed Then Set cnn = GetTmpMdbCn
  If Not GetRS(rsView1, "Select * From SO3110A1", cnn, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
  If Not GetRS(rsView2, "Select * From SO3110A2", cnn, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
  If Not GetRS(rsView3, "Select * From SO3110A2", cnn, adUseClient, adOpenKeyset, adLockPessimistic) Then Exit Sub
  Exit Sub
chkErr:
  Call ErrSub(Me.Name, "OpenData")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
    Call Close3Recordset(rsView1)
    Call Close3Recordset(rsView2)
    Call Close3Recordset(rsView3)
    Call Close3Recordset(rsView4)
    cnn.Close
    Set cnn = Nothing
End Sub
Private Sub toExcel(ByVal strSQL As String)
  On Error GoTo chkErr
    Dim rsExcel As New ADODB.Recordset
    If Not GetRS(rsExcel, strSQL, cnn) Then Exit Sub
    RptToTxt "SO3110AE.rpt", , strSQL, , , "Tmp0000.MDB", , , , gTempPath & "\SO3110A"
    If Not Get_RS_From_Txt(gTempPath, "SO3110A.txt", rsExcel) Then blnExcel = False: Exit Sub
    Call UseProperty(rsExcel, "���X�沧�`����", "�Ĥ@��", , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , "������")
    blnExcel = False
  On Error Resume Next
  CloseRecordset rsExcel
  Exit Sub
chkErr:
  Call ErrSub(Me.Name, "subExcel")
End Sub
Public Property Set rs1(ByRef rs As ADODB.Recordset)
  On Error Resume Next
  Set rsView1 = rs
End Property
Public Property Set rs2(ByRef rs As ADODB.Recordset)
  On Error Resume Next
  Set rsView2 = rs
End Property
Public Property Set rs3(ByRef rs As ADODB.Recordset)
  On Error Resume Next
  Set rsView3 = rs
End Property
Public Property Set rs4(ByRef rs As ADODB.Recordset)
  On Error Resume Next
  Set rsView4 = rs
End Property

Private Sub sstData_Click(PreviousTab As Integer)
  'TABS����W�ŤjBug��..���q���ள��..���MKeyDown�ƥ�|�QĲ�o2�� By Kin 2009/09/17
  On Error Resume Next
    Select Case sstData.Tab
        Case 0
            ggrView1.SetFocus
        Case 1
            ggrView2.SetFocus
        Case 2
            ggrView3.SetFocus
        Case 3
            ggrView4.SetFocus
    End Select
    
End Sub

