VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO19N0A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�ʱ����� [SO19N0A]"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11835
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO19N0A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11835
   StartUpPosition =   1  '���ݵ�������
   Begin VB.Timer tmrAutoUpdate 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   11250
      Top             =   120
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "�w���R�O�^��u��"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5370
      TabIndex        =   25
      Top             =   7110
      Width           =   1785
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5. �C�L"
      Height          =   375
      Left            =   2730
      TabIndex        =   24
      Top             =   7110
      Width           =   1095
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "F6. �]�w"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4050
      TabIndex        =   9
      Top             =   7110
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   10080
      TabIndex        =   11
      Top             =   7110
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "F4. �M��"
      Height          =   375
      Left            =   1410
      TabIndex        =   8
      Top             =   7110
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "F3. �M��"
      Height          =   375
      Left            =   90
      TabIndex        =   7
      Top             =   7110
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "F10.�����R�O"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   7110
      Width           =   1575
   End
   Begin VB.Frame fraData 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   105
      TabIndex        =   12
      Top             =   0
      Width           =   11640
      Begin VB.TextBox txtClock 
         Height          =   285
         Left            =   10710
         TabIndex        =   32
         Text            =   "60000"
         Top             =   750
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox txtCustId 
         Height          =   315
         Left            =   4365
         MaxLength       =   200
         TabIndex        =   30
         Top             =   2250
         Width           =   3300
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "SO19N0A.frx":0442
         Left            =   6930
         List            =   "SO19N0A.frx":0455
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   29
         Top             =   150
         Width           =   2865
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "�۰ʦ^��"
         Enabled         =   0   'False
         Height          =   255
         Left            =   9960
         TabIndex        =   28
         Top             =   180
         Width           =   1125
      End
      Begin VB.Frame fraCommand 
         Caption         =   "���O�B�z����"
         Height          =   705
         Left            =   6990
         TabIndex        =   19
         Top             =   630
         Width           =   3435
         Begin VB.OptionButton optHandleAll 
            Caption         =   "����"
            Height          =   195
            Left            =   2460
            TabIndex        =   22
            Top             =   330
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optHandleYes 
            Caption         =   "�w�B�z"
            Height          =   195
            Left            =   1320
            TabIndex        =   21
            Top             =   330
            Width           =   885
         End
         Begin VB.OptionButton optHandleNo 
            Caption         =   "���B�z"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   330
            Width           =   885
         End
      End
      Begin CS_Multi.CSmulti gimExecEntry 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1830
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   661
         ButtonCaption   =   "�B�z�H��"
      End
      Begin Gi_Time.GiTime gdtUpdTime1 
         Height          =   315
         Left            =   1290
         TabIndex        =   1
         Top             =   660
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Time.GiTime gdtUpdTime2 
         Height          =   315
         Left            =   3615
         TabIndex        =   2
         Top             =   660
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Time.GiTime gdtProcessingDate1 
         Height          =   315
         Left            =   1290
         TabIndex        =   3
         Top             =   1065
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Time.GiTime gdtProcessingDate2 
         Height          =   315
         Left            =   3615
         TabIndex        =   4
         Top             =   1065
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Gi_Multi.GiMulti gimCommand 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1470
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   661
         ButtonCaption   =   "�]�w�Ҧ�"
      End
      Begin prjGiList.GiList gilCompCode 
         Height          =   285
         Left            =   1290
         TabIndex        =   0
         Top             =   195
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth1       =   750
         FldWidth2       =   1600
      End
      Begin prjGiList.GiList gilReturnCode 
         Height          =   300
         Left            =   9090
         TabIndex        =   26
         Top             =   2250
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   529
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth1       =   600
         FldWidth2       =   1500
         F2Corresponding =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "�Ȥ�s��(�H,�۹j�ΥH-���d��)�A�Ҧp: 1,5,11,20-33"
         Height          =   195
         Left            =   150
         TabIndex        =   31
         Top             =   2310
         Width           =   4155
      End
      Begin VB.Label lblReturnCode 
         AutoSize        =   -1  'True
         Caption         =   "�h���]"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   8130
         TabIndex        =   27
         Top             =   2310
         Width           =   780
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "�B�z����"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6090
         TabIndex        =   23
         Top             =   210
         Width           =   780
      End
      Begin VB.Label lblCompCode 
         AutoSize        =   -1  'True
         Caption         =   "���q�O"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   450
         TabIndex        =   17
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   3330
         TabIndex        =   16
         Top             =   1095
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   195
         Left            =   3330
         TabIndex        =   15
         Top             =   705
         Width           =   195
      End
      Begin VB.Label lblProcessingDate 
         Alignment       =   1  '�a�k���
         Caption         =   "�w���ɶ�"
         Height          =   300
         Left            =   330
         TabIndex        =   14
         Top             =   1095
         Width           =   855
      End
      Begin VB.Label lblUpdTime 
         Alignment       =   1  '�a�k���
         Caption         =   "�]�w�ɶ�"
         Height          =   300
         Left            =   330
         TabIndex        =   13
         Top             =   705
         Width           =   855
      End
   End
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   4395
      Left            =   75
      TabIndex        =   6
      Top             =   2670
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   7752
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
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
Attribute VB_Name = "frmSO19N0A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#3353�ݨD(�ʱ�����) By Kin
Option Explicit
Private intCMType As Integer
Dim rsDefTmp As New ADODB.Recordset
Dim strCMOwner As String
Dim strReportSQL As String
Dim intClickCount As Long
Private intAutoUpd As Integer
Private lngTimer As Long
Private Sub subGil()
  On Error GoTo ChkErr
    gilReturnCode.Clear
    SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 20, GetCompCodeFilter(gilCompCode)
    SetgiList gilReturnCode, "CodeNo", "Description", "CD015", , , , , 3, 20, " Where RefNo=3", True
    Exit Sub
ChkErr:
  
  ErrSub Me.Name, "subGil"
End Sub

Private Sub cboType_Click()
    Call ChangeMode
    Call OpenData(True)
End Sub

Private Sub chkAll_Click()
  On Error Resume Next
    Dim i As Long
    If chkAll.Value = 1 Then
        If MsgBox("��� [�۰ʦ^��] �N�|���s�j�M��ơI�O�_�T�w?", vbQuestion + vbYesNo, "�߰ݰT��") = vbYes Then
            Call OpenData(True)
            tmrAutoUpdate.Interval = lngTimer
            tmrAutoUpdate.Enabled = True
            cmdOK.Enabled = False
            cboType.Enabled = False
            For i = 0 To Controls.Count - 1
                If TypeName(Controls(i)) = "CommandButton" Then
                    Controls(i).Enabled = False
                End If
            Next i
            cmdExit.Enabled = True
        Else
            chkAll.Value = 0
        End If
    Else
        tmrAutoUpdate.Enabled = False
        intAutoUpd = 0
        cboType.Enabled = True
        For i = 0 To Controls.Count - 1
                If TypeName(Controls(i)) = "CommandButton" Then
                    Controls(i).Enabled = True
                End If
        Next i
        'cmdExit.Enabled = True
        If cboType.ListIndex = 0 Then
            cmdSetup.Enabled = False
            cmdDelete.Enabled = False
            'cmdOK.Enabled = True
        End If
    End If

End Sub

Private Sub cmdClear_Click()
  On Error GoTo ChkErr
    OpenData (True)
  Exit Sub
ChkErr:
  ErrSub Me.Name, "cmdClear"
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo ChkErr
    Dim StrTableName As String
    Dim strUpdSQL1 As String
    Dim strUpdSQL2 As String
    Dim rsSO0789 As New ADODB.Recordset
    Dim rsSO033 As New ADODB.Recordset
    Dim rsSO004D As New ADODB.Recordset
    Dim rsPRSO004 As New ADODB.Recordset
    Dim rsSo004 As New ADODB.Recordset
    Dim strSO033SQL As String
    Dim strUpdTime As String
    Dim blnSO033 As Boolean
    Dim blnUCCode As Boolean
    Dim strClsName As String
    Dim objclsAlterWip As Object
    Dim strRealDate As String
    
    If Not IsDataOk Then Exit Sub
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    
    If intCMType = 2 Then
        StrTableName = strCMOwner & "SO307 A"
        strUpdSQL1 = "Update " & StrTableName & " Set A.CancelCode=" & gilReturnCode.GetCodeNo & _
                    ",A.CanCelName='" & gilReturnCode.GetDescription & "'" & _
                    ",A.CommandStatus='D'" & _
                    ",A.HandleEn='" & garyGi(0) & "'" & _
                    ",A.HandleName='" & garyGi(1) & "'" & _
                    ",A.HandleTime=to_Date(" & Format(RightNow, "YYYYMMDDHHMMSS") & ",'YYYY/MM/DD HH24:MI:SS')" & _
                    ",A.HandleNote='�ѱ���x�����A�ӵ���ƵL������u��!!'" & _
                    ",A.StopFlag=1" & _
                    ",A.StopDate=to_Date(" & Format(RightNow, "YYYYMMDDHHMMSS") & ",'YYYY/MM/DD HH24:MI:SS')"
        strUpdSQL2 = "Update " & StrTableName & " Set A.CancelCode=" & gilReturnCode.GetCodeNo & _
                    ",A.CanCelName='" & gilReturnCode.GetDescription & "'" & _
                    ",A.HandleEn='" & garyGi(0) & "'" & _
                    ",A.HandleName='" & garyGi(1) & "'" & _
                    ",A.HandleTime=to_Date(" & Format(RightNow, "YYYYMMDDHHMMSS") & ",'YYYY/MM/DD HH24:MI:SS')" & _
                    ",A.StopFlag=1" & _
                    ",A.StopDate=to_Date(" & Format(RightNow, "YYYYMMDDHHMMSS") & ",'YYYY/MM/DD HH24:MI:SS')" & _
                    ",A.HandleNote="
    Else
        StrTableName = strCMOwner & "SO314 A"
        strUpdSQL1 = "Update " & StrTableName & " Set A.CancelCode=" & gilReturnCode.GetCodeNo & _
                    ",A.CanCelName='" & gilReturnCode.GetDescription & "'" & _
                    ",A.StopFlag=1" & _
                    ",A.CancelDate=to_Date(" & Format(RightNow, "YYYYMMDDHHMMSS") & ",'YYYY/MM/DD HH24:MI:SS')" & _
                    ",A.HandleEn='" & garyGi(0) & "'" & _
                    ",A.HandleName='" & garyGi(1) & "'" & _
                    ",A.HandleTime=to_Date(" & Format(RightNow, "YYYYMMDDHHMMSS") & ",'YYYY/MM/DD HH24:MI:SS')" & _
                    ",A.HandleNote='�ѱ���x�����A�ӵ���ƵL������u��!!'"
        strUpdSQL2 = "Update " & StrTableName & " Set A.CancelCode=" & gilReturnCode.GetCodeNo & _
                    ",A.CanCelName='" & gilReturnCode.GetDescription & "'" & _
                    ",A.StopFlag=1" & _
                    ",A.CancelDate=to_Date(" & Format(RightNow, "YYYYMMDDHHMMSS") & ",'YYYY/MM/DD HH24:MI:SS')" & _
                    ",A.HandleEn='" & garyGi(0) & "'" & _
                    ",A.HandleName='" & garyGi(1) & "'" & _
                    ",A.HandleTime=to_Date(" & Format(RightNow, "YYYYMMDDHHMMSS") & ",'YYYY/MM/DD HH24:MI:SS')" & _
                    ",A.HandleNote="
    End If
    
    gcnGi.BeginTrans
    With ggrData.Recordset
        .MoveFirst
        Do While Not .EOF
            If .Fields("Choice") = "1" Then
                '�S�������u��
                If IsNull(.Fields("SNO")) Or .Fields("SNO") & "" = "" Then
                    gcnGi.Execute strUpdSQL1 & " Where RowId='" & .Fields("RowId") & "'"
                Else  '�������u��
'                    strSO033SQL = "Select Count(*) From (Select A.BillNo From " & GetOwner & "SO033 A Where A.BillNo='" & .Fields("SNO") & "' And A.UCCode is Not NULL" & _
'                                " Union All Select B.BillNo From " & GetOwner & "SO034 B Where B.BillNo='" & .Fields("SNO") & "' And B.UCCode is Not Null" & _
'                                " Union All Select C.BillNo From " & GetOwner & "SO035 C Where C.BillNo='" & .Fields("SNO") & "' And C.UCCode is Not Null ) A"
                    strSO033SQL = "Select 0 as Type,RowId," & ChargeField & " From " & GetOwner & "SO033  Where BillNO ='" & .Fields("SNO") & "' Union  All " & _
                                "Select  1 as Type,RowId," & ChargeField & " From " & GetOwner & "SO034 Where BillNO ='" & .Fields("SNO") & "' Union All " & _
                                "Select  2 as Type,RowId," & ChargeField & " From " & GetOwner & "SO035 Where BillNO ='" & .Fields("SNO") & "'"

'                    blnSO033 = Len(gcnGi.Execute(strSO033SQL)("UCCODE")) > 0
                    
                    '�P�_��SO007��SO008��SO009
                    Select Case Mid(.Fields("SNo"), 7, 1)
                        Case "I"
                            StrTableName = GetOwner & "SO007"
                            strClsName = "clsAlterWip1"
                        Case "M"
                            StrTableName = GetOwner & "SO008"
                            strClsName = "clsAlterWip2"
                        Case "P"
                            StrTableName = GetOwner & "SO009"
                            strClsName = "clsAlterWip3"
                    End Select
                    
                    If Not GetRS(rsSO0789, "Select * From " & StrTableName & " Where SNO='" & .Fields("SNO") & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                    
                    '�����u��
                    If Not rsSO0789.EOF Then
                        If Not GetRS(rsSO033, strSO033SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                        If Not GetRS(rsSO004D, "Select A.RowId,A.* From " & GetOwner & "SO004D A Where A.SNO='" & .Fields("SNO") & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                        If Not GetRS(rsPRSO004, "Select A.RowId,A.* From " & GetOwner & "SO004 A Where A.PRSNO='" & .Fields("SNO") & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                        If Not GetRS(rsSo004, "Select A.RowId,A.* From " & GetOwner & "SO004 A Where A.SNO='" & .Fields("SNO") & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                        Set rsSO0789.ActiveConnection = Nothing
                        Set rsSO004D.ActiveConnection = Nothing
                        Set rsPRSO004.ActiveConnection = Nothing
                        Set rsSO033.ActiveConnection = Nothing
                        Set rsSo004.ActiveConnection = Nothing
                        blnSO033 = Not rsSO033.EOF
                        
                        '����즬�O��h�P�_������]�O�_���ȡA�u�n���䤤�@�����ȴN�n���X�A��@�����O����ƳB�z
                        If blnSO033 Then
                            rsSO033.MoveFirst
                            Do While Not rsSO033.EOF
                                If strRealDate = Empty Then
                                    If Not IsNull(rsSO033("RealDate")) Then strRealDate = rsSO033("RealDate")
                                End If
                                If Not IsNull(rsSO033("UCCode")) Then blnUCCode = True: Exit Do
                                rsSO033.MoveNext
                            Loop
                        End If
                        '�u�槹�u�ɶ��L�ȦӥB������]�]�S��
                        If IsNull(rsSO0789("FinTime")) And IsNull(rsSO0789("ReturnCode")) Then
                                
                            Set objclsAlterWip = CreateObject("csAlterWip4." & strClsName)
                            objclsAlterWip.uOwnerName = GetOwner
                            strUpdTime = Format(RightNow, "YYYY/MM/DD HH:MM")
                                
                            '����즬�O���
                            If blnSO033 Then
                                '�u����������O���إ�����]����
                                If blnUCCode Then
                                    gcnGi.Execute (strUpdSQL2 & "'�ӵ���Ƥw�Ѻʱ����ߨ���!!' Where RowId='" & .Fields("RowID") & "'")
                                Else
                                    gcnGi.Execute (strUpdSQL2 & "'�ӵ���Ƥw�Ѻʱ����ߨ����A�������O��Ƥw��" & strRealDate & "����' Where RowId='" & .Fields("RowId") & "'")
                                End If
                            Else '�S�����O���
                                gcnGi.Execute (strUpdSQL2 & "'�ӵ���Ƥw�Ѻʱ����ߨ���!!' Where RowId='" & .Fields("RowID") & "'")
                            End If
                            rsSO0789.Fields("ReturnCode") = gilReturnCode.GetCodeNo
                            rsSO0789.Fields("ReturnName") = gilReturnCode.GetDescription
                            rsSO0789.Fields("SignDate") = .Fields("ResvTime")
                            rsSO0789.Fields("SignEn") = garyGi(0)
                            rsSO0789.Fields("SignName") = garyGi(1)
                            rsSO0789.Fields("UpdEn") = garyGi(1)
                            rsSO0789.Fields("UpdTime") = GetDTString(strUpdTime)
                            'rsSO0789.Update
                            Select Case strClsName
                                Case "clsAlterWip1"
                                    If Not objclsAlterWip.Action(giEditModeEdit, gcnGi, rsSO0789, rsSo004, rsSO033, , , , , False, , , , , rsPRSO004, , , , , , , , , , , , , rsSO004D) Then GoTo lblCallFail
                                Case "clsAlterWip2"
                                    If Not objclsAlterWip.Action(giEditModeEdit, gcnGi, rsSO0789, rsSo004, rsSO033, , , , , False, , , , , rsPRSO004, , , , , , , , , , , , rsSO004D) Then GoTo lblCallFail
                                Case "clsAlterWip3"
                                    If Not objclsAlterWip.Action(giEditModeEdit, gcnGi, rsSO0789, rsSo004, rsSO033, , , , , False, , , , , rsPRSO004, , , , , , , , , , , , , , , rsSO004D) Then GoTo lblCallFail
                            End Select
                        Else '�u�槹�u�ɶ��Ψ�����]����
                            If blnSO033 Then  '����즬�O���
                                If blnUCCode Then
                                    gcnGi.Execute (strUpdSQL2 & "'�ѱ���x�����A" & .Fields("SNO") & "�u��w���u�ΰh��A�����ݦ���ƽЦۦ�B�z!!' Where RowId='" & .Fields("RowID") & "'")
                                Else '�S�����O���
                                    gcnGi.Execute (strUpdSQL2 & "'�ѱ���x�����A" & .Fields("SNO") & "�u��w���u�ΰh��A �������O��Ƥw��" & strRealDate & "����' Where RowId='" & .Fields("RowId") & "'")
                                End If
                            Else
                                gcnGi.Execute (strUpdSQL2 & "'�ӵ���Ƥw�Ѻʱ����ߨ���!!' Where RowId='" & .Fields("RowID") & "'")
                            End If
                        
                        End If
                    Else  '�����ɪ�SNO���ȡA���䤣���ڪ��u��渹
                        gcnGi.Execute strUpdSQL1 & " Where RowId='" & .Fields("RowId") & "'"
                    End If
                    
                End If
            End If
66:
            .MoveNext
        Loop
    End With
    gcnGi.CommitTrans
    Me.Enabled = True
    Me.MousePointer = vbDefault
    On Error Resume Next
    Call CloseRecordset(rsSO0789)
    Call CloseRecordset(rsSO033)
    Call CloseRecordset(rsSO004D)
    Call CloseRecordset(rsPRSO004)
    Call CloseRecordset(rsSo004)
    Set objclsAlterWip = Nothing
    Call OpenData(True)
    Exit Sub
lblCallFail:
    On Error Resume Next
    gcnGi.RollbackTrans
    Me.Enabled = True
    Me.MousePointer = vbDefault
    Call CloseRecordset(rsSO0789)
    Call CloseRecordset(rsSO033)
    Call CloseRecordset(rsSO004D)
    Call CloseRecordset(rsPRSO004)
    Call CloseRecordset(rsSo004)
    'Call CloseRecordset(rsTmp)
    Set objclsAlterWip = Nothing
    'Set objCMControl = Nothing
    MsgBox "�����s��ƥ���!!", vbCritical, "�T��"
    Exit Sub
ChkErr:
    Me.Enabled = True
    Me.MousePointer = vbDefault
    gcnGi.RollbackTrans
    ErrSub Me.Name, "cmdDelete_Click"
End Sub
Private Sub cmdExit_Click()
  On Error Resume Next
  Unload Me
End Sub

Private Sub cmdFind_Click()
  On Error GoTo ChkErr
    Call OpenData(False)
    Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdFind_Click"
End Sub

Private Sub cmdOk_Click()
  On Error GoTo ChkErr
    Dim StrTableName As String
    Dim strUpdSQL1 As String
    Dim strUpdSQL2 As String
    Dim rsSO0789 As New ADODB.Recordset
    Dim rsSO033 As New ADODB.Recordset
    Dim rsSO004D As New ADODB.Recordset
    Dim rsPRSO004 As New ADODB.Recordset
    Dim rsSo004 As New ADODB.Recordset
    Dim strSO033SQL As String
    Dim strUpdTime As String
    Dim strClsName As String
    Dim objclsAlterWip As Object
    Dim objCMControl As Object
    Dim blnCallWip As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim strSO004 As String
    Dim strErrPath As String
    Dim blnAutoFinType As Boolean
    Dim strAutoFinType As String
    blnCallWip = False
    strErrPath = Empty
    strErrPath = ReadGICMIS1("ErrLogPath")
    If ggrData.Recordset.RecordCount = 0 And chkAll.Value = 1 Then Exit Sub
    If Not IsDataOk Then Exit Sub
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    
    
    If intCMType = 2 Then
        StrTableName = strCMOwner & "SO307 A"
        Set objCMControl = CreateObject("csSPMconsole.clsCommand")
    Else
        StrTableName = strCMOwner & "SO314 A"
        Set objCMControl = CreateObject("csSCMcontrol.clsExeCommand")
    End If
    strUpdSQL1 = "Update " & StrTableName & " Set A.HandleEn='" & garyGi(0) & "'" & _
                    ",A.HandleName='" & garyGi(1) & "'" & _
                    ",A.HandleTime=to_Date(" & Format(RightNow, "YYYYMMDDHHMMSS") & ",'YYYY/MM/DD HH24:MI:SS')" & _
                    ",A.HandleNote='�L�����u��!!'"
                    
    strUpdSQL2 = "Update " & StrTableName & " Set A.HandleEn='" & garyGi(0) & "'" & _
                    ",A.HandleName='" & garyGi(1) & "'" & _
                    ",A.HandleTime=to_Date(" & Format(RightNow, "YYYYMMDDHHMMSS") & ",'YYYY/MM/DD HH24:MI:SS')" & _
                    ",A.HandleNote="
    gcnGi.BeginTrans
    With ggrData.Recordset
        .MoveFirst
        Do While Not .EOF
            blnCallWip = False
            If .Fields("Choice") = "1" Then
                '�S�������u��
                If IsNull(.Fields("SNO")) Or .Fields("SNO") & "" = "" Then
                    gcnGi.Execute strUpdSQL1 & " Where RowId='" & .Fields("RowId") & "'"
                    '**************************************************************************
                    '#3353 ���դ�OK,�S�������u��]�n�I�s����x By Kin 2008/04/14
                    '#3353 ���դ�Ok,�n�NSO307.NewModemMAC��: Trim���~���������SO004 By Kin 2008/04/23
                    strSO004 = "Select RowId,A.* From " & GetOwner & "SO004 A" & _
                             " Where A.PrDate is Null" & _
                             " And A.CustId=" & .Fields("CustId") & _
                             " And A.CompCode=" & gilCompCode.GetCodeNo & _
                             " And A.FaciSno='" & Replace(.Fields("DeviceSNo1"), ":", "") & "'"
                    If Not GetRS(rsSo004, strSO004, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                    If rsSo004.RecordCount > 0 Then
                        rsSo004.MoveFirst
                        GoTo CallObj
                    End If
                    '**************************************************************************
                '�������u��
                Else
                    
                    '�P�_��SO007��SO008��SO009
                    Select Case Mid(.Fields("SNo"), 7, 1)
                        Case "I"
                            StrTableName = GetOwner & "SO007"
                            strClsName = "clsAlterWip1"
                        Case "M"
                            StrTableName = GetOwner & "SO008"
                            strClsName = "clsAlterWip2"
                        Case "P"
                            StrTableName = GetOwner & "SO009"
                            strClsName = "clsAlterWip3"
                    End Select
                    
                    If Not GetRS(rsSO0789, "Select * From " & StrTableName & " Where SNO='" & .Fields("SNO") & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                    '*******************************************************************************************************************
                    '#4109 �P�_�n���u�W�^���ɶ��Χ��u�ɶ� By Kin 2008/10/06
                    blnAutoFinType = GetRsValue("Select nvl(AutoFinType,0) From " & GetOwner & "SO042 Where CompCode=" & _
                                                rsSO0789("CompCode") & " And ServiceType='" & rsSO0789("ServiceType") & "'", gcnGi)
                    If blnAutoFinType Then
                        strAutoFinType = "FinTime"
                    Else
                        strAutoFinType = "CallOkTime"
                    End If
                    '*******************************************************************************************************************
                    '�����u��
                    If Not rsSO0789.EOF Then
                        'If Not GetRS(rsSO033, "Select * From " & GetOwner & "SO033 Where BillNo='" & .Fields("SNO") & "'" & IIf(blnSO033, " And UCCode is Not Null", " And UCCode Is Null"), gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                        If Not GetRS(rsSO033, "Select A.RowId,A.* From " & GetOwner & "SO033 A Where 1=0", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                        If Not GetRS(rsSO004D, "Select A.RowId,A.* From " & GetOwner & "SO004D A Where A.SNO='" & .Fields("SNO") & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                        If Not GetRS(rsPRSO004, "Select A.RowId,A.* From " & GetOwner & "SO004 A Where A.PRSNO='" & .Fields("SNO") & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                        If Not GetRS(rsSo004, "Select RowId,A.* From " & GetOwner & "SO004 A Where A.SNO='" & .Fields("SNO") & "'", gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
                        Set rsSO0789.ActiveConnection = Nothing
                        Set rsSO004D.ActiveConnection = Nothing
                        Set rsPRSO004.ActiveConnection = Nothing
                        Set rsSO033.ActiveConnection = Nothing
                        Set rsSo004.ActiveConnection = Nothing
                        '�u�槹�u�ɶ��L�ȦӥB������]�]�S��
                        '#4109 �쥻�O�P�_���u�ɶ�,�{�b��ѰѼư��P�_ By Kin 2008/10/06
                        If IsNull(rsSO0789(strAutoFinType)) And IsNull(rsSO0789("ReturnCode")) Then
                            Set objclsAlterWip = CreateObject("csAlterWip4." & strClsName)
                            objclsAlterWip.uOwnerName = GetOwner
                            strUpdTime = Format(RightNow, "YYYY/MM/DD HH:MM")
                            gcnGi.Execute (strUpdSQL2 & "'�ѱ���x" & IIf(blnAutoFinType, "���u", "�u�W�^��") & ",�����ʦ��O���!!' Where RowId='" & .Fields("RowID") & "'")
                            '#4109 �ѼƬ����u�~�i�񧹤u�ɶ� By Kin 2008/10/06
                            If blnAutoFinType Then
                                rsSO0789.Fields("FinTime") = .Fields("ResvTime")
                                rsSO0789.Fields("SignDate") = .Fields("ResvTime")
                            End If
                            '**********************************************************************
                            '#3353 ���դ�OK �^���ɶ���null������J�w���ɶ� By Kin 2008/04/14
                            '#4434 �p�G�����u�ɶ�,�h���b�񧹤u�ɶ� By Kin 2009/04/10
                            If IsNull(rsSO0789("CallOkTime")) Then
                                rsSO0789.Fields("CallOkTime") = .Fields("ResvTime")
                            End If
                            '**********************************************************************
                            '#4109 �ѼƬ����u�~�i�񧹤u�ɶ� By Kin 2008/10/06
                            If blnAutoFinType Then
                                rsSO0789.Fields("SignEn") = garyGi(0)
                                rsSO0789.Fields("SignName") = garyGi(1)
                            End If
                            rsSO0789.Fields("UpdEn") = garyGi(1)
                            rsSO0789.Fields("UpdTime") = GetDTString(strUpdTime)
                            'rsSO0789.Update
                            blnCallWip = True
                            GoTo CallObj
                        End If
                        '#4109 �쥻�O�P�_���u�ɶ�,�{�b��ѰѼư��P�_ By Kin 2008/10/06
                        If Not IsNull(rsSO0789(strAutoFinType)) Then
                            gcnGi.Execute (strUpdSQL2 & "'�ѱ���x����," & rsSO0789("SNO") & "�渹�w" & IIf(blnAutoFinType, "���u", "�u�W�^��") & ",�����ʦ��O���!!' Where RowId='" & .Fields("RowID") & "'")
                            GoTo CallObj
                        End If
                        If Not IsNull(rsSO0789("ReturnCode")) Then
                            gcnGi.Execute (strUpdSQL2 & "'�ѱ���x����," & rsSO0789("SNO") & "�渹�w�h�� [�������ʸ��]!!' Where RowId='" & .Fields("RowID") & "'")
                            GoTo CallObj
                        End If
CallObj:
                        If intCMType = 2 Then
                            Do While Not rsSo004.EOF
                                If Not objCMControl.ExecCmd(gcnGi, rsSo004, .Fields("MsgName") & "", Join(garyGi, Chr(9)), strErrPath, , False, , , , , , , .Fields("CmdSeqNo")) Then GoTo lblCallCMFail
                                rsSo004.MoveNext
                            Loop
                        Else
                            Do While Not rsSo004.EOF
                                If Not objCMControl.ExeCMCommand(gcnGi, rsSo004, .Fields("MsgCode") & "", garyGi, , , , , , , , , , False, .Fields("CmdSeqNo")) Then GoTo lblCallFail
                                rsSo004.MoveNext
                            Loop
                        End If
                        If blnCallWip Then
                            Select Case strClsName
                                Case "clsAlterWip1"
                                    If Not objclsAlterWip.Action(giEditModeEdit, gcnGi, rsSO0789, rsSo004, Nothing, , , , , False, , , , , rsPRSO004, , , , , , , , , , , , , rsSO004D) Then GoTo lblCallFail
                                Case "clsAlterWip2"
                                    If Not objclsAlterWip.Action(giEditModeEdit, gcnGi, rsSO0789, rsSo004, Nothing, , , , , False, , , , , rsPRSO004, , , , , , , , , , , , rsSO004D) Then GoTo lblCallFail
                                Case "clsAlterWip3"
                                    If Not objclsAlterWip.Action(giEditModeEdit, gcnGi, rsSO0789, rsSo004, Nothing, , , , , False, , , , , rsPRSO004, , , , , , , , , , , , , , , rsSO004D) Then GoTo lblCallFail
                            End Select
                        End If
                        
                    Else  '�����ɪ�SNO���ȡA���䤣���ڪ��u��渹
                        gcnGi.Execute strUpdSQL1 & " Where RowId='" & .Fields("RowId") & "'"
                    End If
                    
                End If
            End If
            
66:
            .MoveNext
        Loop
    End With
    gcnGi.CommitTrans
    On Error Resume Next
    Me.Enabled = True
    Me.MousePointer = vbDefault
    Call CloseRecordset(rsSO0789)
    Call CloseRecordset(rsSO033)
    Call CloseRecordset(rsSO004D)
    Call CloseRecordset(rsPRSO004)
    Call CloseRecordset(rsSo004)
    Call CloseRecordset(rsTmp)
    Set objclsAlterWip = Nothing
    Set objCMControl = Nothing
    Call OpenData(True)
    
    Exit Sub
lblCallCMFail:
    On Error Resume Next
    gcnGi.RollbackTrans
    Me.Enabled = True
    Me.MousePointer = vbDefault
    Call CloseRecordset(rsSO0789)
    Call CloseRecordset(rsSO033)
    Call CloseRecordset(rsSO004D)
    Call CloseRecordset(rsPRSO004)
    Call CloseRecordset(rsSo004)
    Call CloseRecordset(rsTmp)
    Set objclsAlterWip = Nothing
    Set objCMControl = Nothing
    MsgBox "����x�����s��ƥ���!!", vbCritical, "�T��"
    Exit Sub

lblCallFail:
    On Error Resume Next
    gcnGi.RollbackTrans
    Me.Enabled = True
    Me.MousePointer = vbDefault
    Call CloseRecordset(rsSO0789)
    Call CloseRecordset(rsSO033)
    Call CloseRecordset(rsSO004D)
    Call CloseRecordset(rsPRSO004)
    Call CloseRecordset(rsSo004)
    Call CloseRecordset(rsTmp)
    Set objclsAlterWip = Nothing
    Set objCMControl = Nothing
    MsgBox "�u�椸���s��ƥ���!!", vbCritical, "�T��"
    Exit Sub
ChkErr:
    Me.MousePointer = vbDefault
    Me.Enabled = True
    gcnGi.RollbackTrans
    ErrSub Me.Name, "CmdOk_Click"
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo ChkErr
    If Not ChkDTok Then Exit Sub
    'If Not IsDataOk Then Exit Sub
    Screen.MousePointer = vbHourglass
      'cmdExit.SetFocus
    Me.Enabled = False
    ReadyGoPrint
    Call subChoose
    Call subPrint
    Me.Enabled = True
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub
Private Function subCreateView() As Boolean
  On Error Resume Next
    Dim strsql As String
    Dim StrTableName As String
    Dim strFields As String
    gcnGi.Execute "Drop Table " & strViewName
    If intCMType = 2 Then
        StrTableName = strCMOwner & "SO307 A"
        strFields = "A.CmdSeqNo,A.CompCode,A.MsgCode,A.MsgName,A.CustId,A.SNO,A.ExecTime,A.ResvTime,A.HandleTime,A.COMMANDSTATUS,HandleNote"
    Else
        StrTableName = strCMOwner & "SO314 A"
        strFields = "A.CmdSeqNo,A.CompCode,A.MsgCode,A.MsgName,To_Number(A.CustId) CustId,A.SNO,A.UpdTime ExecTime," & _
                    "A.ProcessingDate ResvTime,A.HandleTime,A.CmdStatus COMMANDSTATUS,HandleNote"
    End If
    On Error GoTo ChkErr
      strViewName = GetTmpViewName
      strsql = "Create Table " & strViewName & " as (" & _
             "Select " & strFields & " From " & StrTableName & " Where " & subChoose & ")"
    gcnGi.Execute strsql
    gcnGi.Execute "ALTER TABLE " & strViewName & " NOLOGGING"
    On Error GoTo ChkErr
    SendSQL strsql, True
    subCreateView = True
    Exit Function
ChkErr:
    ErrSub Me.Name, "subCreateView"
End Function

Private Sub subPrint()
  On Error GoTo ChkErr
    Dim strsql As String
    strGroupName = Empty
    If intCMType = 3 Then
        strGroupName = "Type=" & intCMType & ";SortA={V.MsgCode};GroupName={V.CustId}"
    Else
        strGroupName = "Type=" & intCMType & ";SortA={V.MsgName};GroupName={V.CustId}"
    End If
    If Not subCreateView Then Exit Sub
    strsql = "Select * From " & strViewName & " V "
    If gcnGi.Execute("Select Count(*) From " & strViewName)(0) = 0 Then
        MsgNoRcd
        SendSQL , , True
    Else
        Call PrintRpt2(GetPrinterName(5), RptName("SO19N0"), "SO19N0", Me.Caption, strsql, strChooseString, , True, , , strGroupName)
    End If
    
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subPrint")
End Sub

Private Sub cmdSetup_Click()
  On Error GoTo ChkErr
    Dim rsSource As New ADODB.Recordset
    Dim rsDest As New ADODB.Recordset
    Dim StrTableName As String
    Dim strDest As String
    Dim strUpdTime As String
    Dim i As Long
    
    'If ggrData.Recordset.RecordCount <= 0 Then MsgBox "�L�����ƥi����I", vbInformation, "�T��": Exit Sub
    If Not IsDataOk Then Exit Sub
    'If intClickCount = 0 Then MsgBox "�Цܤֿ�ܤ@����ơI", vbInformation, "�T��": Exit Sub
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    If intCMType = 2 Then
        StrTableName = strCMOwner & "SO307 A"
        strDest = "Select * From " & strCMOwner & "SO307 A Where 1=0"
    Else
        StrTableName = strCMOwner & "SO314 A"
        strDest = "Select * From " & strCMOwner & "SO314 A Where 1=0"
    End If
    If Not GetRS(rsDest, strDest, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    strUpdTime = Format(RightNow, "YYYY/MM/DD HH:MM:SS")
    ggrData.Recordset.MoveFirst
    gcnGi.BeginTrans
    Do While Not ggrData.Recordset.EOF
        If ggrData.Recordset("Choice") = "1" Then
            Set rsSource = gcnGi.Execute("Select * From " & StrTableName & " Where ROWID='" & ggrData.Recordset("RowId") & "'")
            rsDest.AddNew
            For i = 0 To rsSource.Fields.Count - 1
                If UCase(rsSource.Fields(i).Name) = "CMDSEQNO" Then
                    rsDest("CmdSeqNo") = Get_Cmd_Seq_No
                ElseIf UCase(rsSource.Fields(i).Name) = "COMMANDSTATUS" Then
                    rsDest("COMMANDSTATUS") = "W"
                ElseIf UCase(rsSource.Fields(i).Name) = "CMDSTATUS" Then
                    rsDest("CmdStatus") = "W"
                ElseIf UCase(rsSource.Fields(i).Name) = "OPERATORID" Then
                    rsDest("OPERATORID") = garyGi(0)
                ElseIf UCase(rsSource.Fields(i).Name) = "EXECENTRY" Then
                    rsDest("EXECENTRY") = garyGi(1)
                ElseIf UCase(rsSource.Fields(i).Name) = "COMPCODE" Then
                    rsDest("CompCode") = gilCompCode.GetCodeNo
                ElseIf UCase(rsSource.Fields(i).Name) = "PROCESSINGDATE" Then
                    rsDest("PROCESSINGDATE") = strUpdTime
                ElseIf UCase(rsSource.Fields(i).Name) = "UPDTIME " Then
                    rsDest("UPDTIME") = strUpdTime
                ElseIf UCase(rsSource.Fields(i).Name) = "RESVTIME" Then
                    rsDest("RESVTIME") = strUpdTime
                ElseIf UCase(rsSource.Fields(i).Name) = "EXECTIME" Then
                    rsDest("EXECTIME") = strUpdTime
                ElseIf UCase(rsSource.Fields(i).Name) = "HANDLEEN" Then
                    'rsDest("HANDLEEN") = garyGi(0)
                ElseIf UCase(rsSource.Fields(i).Name) = "HANDLENAME" Then
                    'rsDest("HANDLENAME") = garyGi(1)
                ElseIf UCase(rsSource.Fields(i).Name) = "HANDLETIME" Then
                    'rsDest("HANDLETIME") = strUpdTime
                ElseIf UCase(rsSource.Fields(i).Name) = "HANDLENOTE" Then
                    'rsDest("HANDLENOTE") = "���O�w���e!!"
                Else
                    If Not IsNull(rsSource.Fields(i).Value) Then
                        rsDest.Fields(rsSource.Fields(i).Name) = rsSource.Fields(i).Value & ""
                    End If
                End If
            Next i
            rsDest.Update
        End If
        '#3353 ���դ�OK,�w���e�����O�nUPDATE�U�C������� By Kin 2008/03/25
        gcnGi.Execute "Update " & StrTableName & _
                        " Set HandleEn='" & garyGi(0) & "'" & _
                        ",HandleName='" & garyGi(1) & "'" & _
                        ",HandleTime=To_Date(" & "'" & strUpdTime & "','YYYY/MM/DD HH24:MI:SS'" & ")" & _
                        ",HandleNote='���O�w���e!!'" & _
                        " Where CmdSeqNo=" & ggrData.Recordset("CmdSeqNo") & _
                        " AND COMPCODE=" & gilCompCode.GetCodeNo
        ggrData.Recordset.MoveNext
    Loop
    gcnGi.CommitTrans
    Call CloseRecordset(rsDest)
    Call CloseRecordset(rsSource)
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Call OpenData(True)
    
    Exit Sub
ChkErr:
    Me.MousePointer = vbDefault
    Me.Enabled = True
    gcnGi.RollbackTrans
    ErrSub Me.Name, "cmdSetup_Click"
End Sub
Private Function CmdDisp(Value As Variant) As Variant
  On Error GoTo ChkErr
    CmdDisp = Switch(Value = "C1", "�˾� (CM)", Value = "E1", "�˾� (EMTA)", _
                                    Value = "C2", "��� (CM�˾��h��)", Value = "E2", "��� (EMTA�˾��h��)", _
                                    Value = "E3", "CP �˾�", Value = "A1_P", "CP Only ���", _
                                    Value = "E4", "EMTA �[�� CP ����", Value = "E5", "EMTA �� CP ����", _
                                    Value = "A1_D", "�n��", Value = "A1_C", "�n�_", _
                                    Value = "A1_BU", "�t�v�ɭ���", Value = "A1_PIP", "�ӽаʺAIP", _
                                    Value = "A1_MIP", "�����ʺAIP", Value = "Q1", "CM reset", _
                                    Value = "Q2", "�d��CM���p", Value = "A7", "CM ���� / ���s", _
                                    Value = "A8", "EMTA ���� / ���s", Value = "A9", "����CM / ���sEMTA", _
                                    Value = "A10", "����EMTA / ���s CM", Value = "A6", "�ܧ� CPE MAC", _
                                    Value = "A5", "CM IP  CPE IP �]�w", Value = "A4", "EMTA IP Priv IP �]�w", _
                                    Value = "A3", "�����T�wIP", Value = "A2", "�ӽЩT�wIP", _
                                    Value = "C6", "CM ����", Value = "E6", "EMTA ����")
  Exit Function
ChkErr:
    ErrSub Me.Name, "CmdDisp"
End Function

Private Sub Form_Activate()
  On Error Resume Next
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If Not ChkGiList(KeyCode) Then Exit Sub
    If KeyCode = vbKeyS And Shift = 6 Then
        txtClock.Visible = Not txtClock.Visible
    End If
    Select Case KeyCode
        Case vbKeyF3
             If cmdFind.Enabled Then cmdFind.Value = True
        Case vbKeyF4
             If cmdClear.Enabled Then cmdClear.Value = True
        Case vbKeyF5
             If cmdPrint.Enabled Then cmdPrint.Value = True
        Case vbKeyF6
             If cmdSetup.Enabled Then cmdSetup.Value = True
        Case vbKeyF10
             If cmdDelete.Enabled Then cmdDelete.Value = True
        Case vbKeyEscape
             If Me.Enabled Then Unload Me
        
    End Select
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")

End Sub
Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If ggrData.Recordset.RecordCount = 0 Then
        MsgBox "�L�����ƥi����I", vbExclamation, "�T��"
        Exit Function
    End If
    If Not ChkDTok Then Exit Function
    If Not MustExist(gilCompCode, 2, "���q�O") Then Exit Function
    If cboType.ListIndex = 1 Or cboType.ListIndex = 3 Then
        If gilReturnCode.GetCodeNo = Empty Then
            MsgBox "�h���]�������ȡI�I", vbInformation, "�T��"
            Exit Function
        End If
    End If
    If chkAll.Value <> 1 Then
        If intClickCount = 0 Then MsgBox "�Цܤֿ�ܤ@����ơI", vbInformation, "�T��": Exit Function
    End If
    IsDataOk = True
    Exit Function
    
ChkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub Form_Load()
  On Error GoTo ChkErr
    gdtProcessingDate1.ShowSecond
    gdtProcessingDate2.ShowSecond
    gdtUpdTime1.ShowSecond
    gdtUpdTime2.ShowSecond
    lngTimer = 60000
    Call subGil
    Call subGim
    Call subGrd
    Call InitialCtl
    
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
    cboType.ListIndex = 0
    Exit Sub
ChkErr:
  ErrSub Me.Name, "Form_Load"
End Sub
Private Sub subGim()
  On Error GoTo ChkErr
    Dim strWhere As String
    gimCommand.Clear
    gimExecEntry.Clear
    If intCMType = 2 Then
        strWhere = " Where Exists(Select * From " & GetOwner & "SO307 Where CM003.EmpNo=SO307.execentry)"
    Else
        strWhere = " Where Exists(Select * From " & GetOwner & "SO314 Where CM003.EmpNo=SO314.OperatorId)"
    End If
    SetgiMulti gimExecEntry, "EmpNo", "EmpName", "CM003", "���u�N��", "���u�m�W", strWhere
    If intCMType = 2 Then
        SetgiMultiAddItem gimCommand, "C1,E1,C2,E2,A1_P,A1_D,A1_C,E3,A1_BU,A1_PIP,A1_MIP,A2,A3,E4,E5,A4,A5,A6,A7,A8,A9,A10,Q1,Q2,C6,E6", _
                                    "CM�}��,EMTA�}��,CM���,EMTA���,CP Only���,CM�n��,CM�n�_,CP���(�w��EMTA),�t�v�ɭ���,�ӽаʺAIP," & _
                                    "�����ʺAIP,�ӽЩT�wIP,�����T�wIP,CP�ˤ��u,CP����u,EMTA ��IP,CPE IP ��,�ܧ�CPE MAC,CM��CM,EMTA��EMTA," & _
                                    "CM��EMTA,EMTA��CM,���] CM reset,�d��CM���p,CM ����,EMTA ����", "�N�X", "�W��"
    Else
        SetgiMultiAddItem gimCommand, "10,11,12,13,14,15,16,20,21,22,23,24,30,31,40,41", _
                                    "�˾�,�}��,����,����(�n��),�_��(�n�_),���,���״���,�ܧ�򥻸��,�ܧ�ӽФH," & _
                                    "�ܧ�t�v,�ܧ����,�ܧ�K�X,�d��CM��T,�d�߱b����T,�]�ƤJ�w,�]�ƥX�w", "�N�X", "�W��"
    End If
    Exit Sub
ChkErr:
    
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub subGrd()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    With mFlds
        .Add "Choice", , , , , "���"
        .Add "CommandStatus", , , , , "���A" & Space(20)
        .Add "CustId", , , , , "�Ȥ�s�� "
        .Add "SNO", , , , , "�u��渹" & Space(10)
        .Add "ModemMAC", , , , , "CM MAC" & Space(10)
        .Add "MsgName", , , , , "����R�O"
        .Add "ResvTime", , , , , "�w���ɶ�" & Space(20)
        .Add "ExecEntry", , , , , "����H��"
        .Add "Exectime", , , , , "����ɶ�" & Space(20)
        '#3917 �W�[HandleNote��� By Kin 2008/04/30
        .Add "HandleNote", , , , , "�B�z����" & Space(80)
    End With
    ggrData.AllFields = mFlds
    ggrData.SetHead
    Exit Sub

ChkErr:
   
   Call ErrSub(Me.Name, "SubGrd")
End Sub
Private Sub InitialCtl()
  On Error Resume Next
    optHandleYes.Enabled = True
    optHandleNo.Enabled = False
    optHandleAll.Enabled = False
    optHandleYes.Value = True
    cmdOK.Enabled = True
    cmdSetup.Enabled = False
    gilReturnCode.Enabled = False
    cmdDelete.Enabled = False
    chkAll.Enabled = True
    tmrAutoUpdate.Interval = 60000
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  Call CloseRecordset(rsDefTmp)
  gcnGi.Execute "Drop Table " & strViewName
  
End Sub

Private Sub ggrData_ColorCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error Resume Next
    If UCase(giFld.FieldName) = "CHOICE" Then
        If Value = 1 Then Value = vbRed
    End If
End Sub

Private Sub ggrData_DblClick()
   On Error GoTo ChkErr
    Dim blnCanChoice As Boolean
    Static intChoice As Long
    blnCanChoice = True
    intChoice = intClickCount
    If cboType.ListIndex = 4 Then Exit Sub
'    '(2) ��CMType = 2 ��, COMMANDSTATUS <> 'W','C','S','D'�~�i����C(V0.2)
'    If cboType.ListIndex = 3 Then
'        If intCMType = 2 Then
'            Select Case ggrData.Recordset("CommandStatus")
'                Case "W"
'                    blnCanChoice = False
'                Case "C"
'                    blnCanChoice = False
'                Case "S"
'                    blnCanChoice = False
'                Case "D"
'                    blnCanChoice = False
'            End Select
'        Else
'            If UCase(ggrData.Recordset("CommandStatus")) <> "E" Then blnCanChoice = False
'        End If
'    End If
    If Not blnCanChoice Then Exit Sub
    If ggrData.Recordset.RecordCount > 0 Then
        If ggrData.Recordset("Choice") = "1" Then
            ggrData.Recordset("Choice") = "0"
            If intChoice > 0 Then intChoice = intChoice - 1: intClickCount = intChoice
        Else
            ggrData.Recordset("Choice") = "1"
            intChoice = intChoice + 1
            intClickCount = intChoice
        End If
        ggrData.Recordset.Update
    End If
    Exit Sub
ChkErr:
    ErrSub Me.Name, "ggrData_DblClick"
End Sub

Private Sub ggrdata_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error Resume Next
    If UCase(giFld.FieldName) = "CHOICE" Then
        Select Case Value
            Case 1
                Value = "V"
            Case Else
                Value = ""
        End Select
    End If
    Select Case intCMType
        Case 2
            If UCase(giFld.FieldName) = "COMMANDSTATUS" Then
                Select Case UCase(Value)
                    Case "W"
                        Value = "�ȪA�g�J�ݳB�z"
                    Case "E"
                        Value = "GateWay ���榳���~"
                    Case "S"
                        Value = "GateWay ���榨�\"
                    Case "F"
                        Value = "GateWay ���楢��"
                    Case "C"
                        Value = "GateWay �w����"
                    Case "D"
                        Value = "�ȪA�t�Ψ���"
                    Case "T"
                        Value = "�ȪA�t�� TimeOut GateWay �ݭnrollback"
                    Case "GT"
                        Value = "GatewWay �Pip command time out"
                    Case "TR"
                        Value = "GateWay rollback  ���\"
                    Case "TE"
                        Value = "GateWay rollback  ����"
                End Select
            End If
            If UCase(giFld.FieldName) = "MSGNAME" Then
                Value = CmdDisp(Value)
            End If
        Case 3
            If UCase(giFld.FieldName) = "COMMANDSTATUS" Then
                Select Case UCase(Value)
                    Case "W"
                       Value = "���B�z"
                    Case "E"
                       Value = "���~"
                    Case "P"
                       Value = "�B�z��"
                    Case "C"
                       Value = "���\"
                End Select
            End If
    End Select
    
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr:
    garyGi(16) = gilCompCode.GetOwner
    gCompCode = gilCompCode.GetCodeNo
    garyGi(9) = gilCompCode.GetCodeNo
    
    intCMType = Val(GetRsValue("Select Nvl(CMType,0) From " & GetOwner & "CD039 Where CodeNo=" & gCompCode, gcnGi))
    If intCMType <> 2 And intCMType <> 3 Then Exit Sub
    Call OpenData(True)
    Call subGil
    Call subGim

  Exit Sub
ChkErr:
  
  ErrSub Me.Name, "gilCompCode_Change"
End Sub
Private Function OpenData(ByVal blnClear As Boolean) As Boolean
  On Error GoTo ChkErr
    Dim strQrySQL As String
    Dim strField As String
    Dim StrTableName As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    strReportSQL = Empty
    OpenData = False
    strCMOwner = Empty
    strCMOwner = gcnGi.Execute("Select CMOwner From " & GetOwner & "SO041 Where CompCode=" & gCompCode)(0) & ""
    If strCMOwner <> Empty Then strCMOwner = strCMOwner & "."
    Select Case intCMType
        Case 2
            StrTableName = strCMOwner & "SO307 A "
            strField = "A.RowId,A.COMMANDSTATUS,A.CustID,A.SNO,A.ModemMAC,A.ResvTime,A.ExecEntry," & _
                       "A.Exectime,A.MsgName,A.MsgCode,A.CmdSeqNo,A.NEWMODEMMAC DeviceSNo1,A.HandleNote  "
            
        Case 3
            StrTableName = strCMOwner & "SO314 A "
            strField = "A.RowId,A.CmdStatus COMMANDSTATUS,A.CustId,A.SNO,A.DeviceSNo1 ModemMac," & _
                        "A.ProcessingDate ResvTime,A.OperatorId ExecEntry,A.UpdTime Exectime,A.MsgName," & _
                        "A.MsgCode,A.CmdSeqNo,A.DeviceSNo1,A.HandleNote "
        Case Else
            MsgBox "CM�������A ������2��3�I", vbInformation, "�T��"
            On Error Resume Next
            Unload Me
            Exit Function
    End Select
    
    Me.Enabled = False
    Me.MousePointer = vbHourglass
    DefineRs
    LockWindowUpdate Me.hwnd
    If Not blnClear Then
        strQrySQL = "Select " & strField & " From " & StrTableName & " Where " & subChoose
        'strReportSQL = strQrySQL
        If Not GetRS(rsTmp, strQrySQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
        Set rsTmp.ActiveConnection = Nothing
        If chkAll.Value = 0 Then
            If rsTmp.RecordCount <= 0 Then MsgBox "�䤣������ơI", vbInformation, "�T��"
        End If
        Do While Not rsTmp.EOF
            rsDefTmp.AddNew
            If chkAll.Value = 1 Then
                rsDefTmp.Fields("Choice") = "1"
            End If
            rsDefTmp.Fields("CommandStatus") = rsTmp.Fields("CommandStatus") & ""
            rsDefTmp.Fields("CustId") = rsTmp.Fields("CustId") & ""
            rsDefTmp.Fields("SNO") = rsTmp.Fields("SNO") & ""
            rsDefTmp.Fields("ModemMAC") = rsTmp.Fields("ModemMAC") & ""
            rsDefTmp.Fields("ResvTime") = rsTmp.Fields("ResvTime") & ""
            rsDefTmp.Fields("ExecEntry") = rsTmp.Fields("ExecEntry") & ""
            rsDefTmp.Fields("Exectime") = rsTmp.Fields("Exectime") & ""
            rsDefTmp.Fields("RowId") = rsTmp.Fields("RowId") & ""
            rsDefTmp.Fields("MsgName") = rsTmp.Fields("MsgName") & ""
            rsDefTmp.Fields("MsgCode") = rsTmp.Fields("MsgCode") & ""
            rsDefTmp.Fields("CmdSeqNo") = rsTmp.Fields("CmdSeqNo") & ""
            rsDefTmp.Fields("DeviceSNo1") = rsTmp.Fields("DeviceSNo1") & ""
            '#3917 �W�[HandleNote���
            rsDefTmp.Fields("HandleNote") = IIf(IsNull(rsTmp.Fields("HandleNote")), "", rsTmp.Fields("HandleNote") & "")
            rsDefTmp.Update
            rsTmp.MoveNext
        Loop
    Else
        intClickCount = 0
    End If
    ggrData.Blank = True
    Set ggrData.Recordset = rsDefTmp
    ggrData.Refresh
    ggrData.Blank = False
    Me.Enabled = True
    Me.MousePointer = vbDefault
    LockWindowUpdate 0
    
    Call CloseRecordset(rsTmp)
    OpenData = True
    Exit Function
ChkErr:
  ErrSub Me.Name, "OpenData"
End Function
Private Function subChoose() As String
  On Error GoTo ChkErr
    strChoose = Empty
    strChooseString = Empty
    If gilCompCode.GetCodeNo <> Empty Then Call subAnd("A.CompCode=" & gilCompCode.GetCodeNo)
    If gdtUpdTime1.GetValue <> Empty Then
        Select Case intCMType
            Case 2
                Call subAnd("A.Exectime>=To_Date(" & gdtUpdTime1.GetValue & ",'YYYY/MM/DD HH24:MI:SS')")
            Case 3
                Call subAnd("A.UpdTime>=To_Date(" & gdtUpdTime1.GetValue & ",'YYYY/MM/DD HH24:MI:SS')")
        End Select
    End If
    If gdtUpdTime2.GetValue <> Empty Then
        Select Case intCMType
            Case 2
                Call subAnd("A.Exectime<=To_Date(" & gdtUpdTime2.GetValue & ",'YYYY/MM/DD HH24:MI:SS')")
            Case 3
                Call subAnd("A.UpdTime<=To_Date(" & gdtUpdTime2.GetValue & ",'YYYY/MM/DD HH24:MI:SS')")
        End Select
    End If
    If gdtProcessingDate1.GetValue <> Empty Then
        Select Case intCMType
            Case 2
                Call subAnd("A.ResvTime>=To_Date(" & gdtProcessingDate1.GetValue & ",'YYYY/MM/DD HH24:MI:SS')")
            Case 3
                Call subAnd("A.ProcessingDate>=To_Date(" & gdtProcessingDate1.GetValue & ",'YYYY/MM/DD HH24:MI:SS')")
        End Select
    End If
    If gdtProcessingDate2.GetValue <> Empty Then
        Select Case intCMType
            Case 2
                Call subAnd("A.ResvTime<=To_Date(" & gdtProcessingDate2.GetValue & ",'YYYY/MM/DD HH24:MI:SS')")
            Case 3
                Call subAnd("A.ProcessingDate<=To_Date(" & gdtProcessingDate2.GetValue & ",'YYYY/MM/DD HH24:MI:SS')")
        End Select
    End If
    '#3353 ���դ�OK,�p�G�O���SO307,�h�R�O�W�٬�MsgName By Kin 2008/03/24
    If gimCommand.GetQueryCode <> Empty Then
        If intCMType = 2 Then
            Call subAnd("A.MsgName In(" & gimCommand.GetQueryCode & ")")
        Else
            Call subAnd("A.MsgCode In(" & gimCommand.GetQueryCode & ")")
        End If
    End If
    If gimExecEntry.GetQueryCode <> Empty Then
        Select Case intCMType
            Case 2
                Call subAnd("A.ExecEntry In(" & gimExecEntry.GetQueryCode & ")")
            Case 3
                Call subAnd("A.OperatorId In(" & gimExecEntry.GetQueryCode & ")")
        End Select
    End If
    If txtCustId.Text <> Empty Then
        Call TimetxtCustId(txtCustId)
    End If
            
    '�B�z����
    Select Case cboType.ListIndex
        Case 0  '�w���R�O������u��(�@�w�n���\���~�B�z)
            If intCMType = 2 Then
                Call subAnd("A.COMMANDSTATUS='S'")
            Else
                Call subAnd("A.CmdStatus='C'")
            End If
        Case 1
            
        Case 2  '���~�R�O���e(�@�w�n�����~���~�B�z)
            If intCMType = 2 Then
                Call subAnd("A.COMMANDSTATUS<>'W' And A.COMMANDSTATUS<>'S' And A.COMMANDSTATUS<>'S' And A.COMMANDSTATUS<>'D'")
            Else
                Call subAnd("A.CmdStatus='E'")
            End If
            
        Case 3  '�������~�R�O(�@�w�n�����~���~�B�z)
            If intCMType = 2 Then
                Call subAnd("A.COMMANDSTATUS<>'W' And A.COMMANDSTATUS<>'S' And A.COMMANDSTATUS<>'S' And A.COMMANDSTATUS<>'D'")
            Else
                Call subAnd("A.CmdStatus='E'")
            End If
    End Select
    
    '���O�B�z����
    If optHandleNo.Value Then   '���B�z
        If intCMType = 2 Then
            Call subAnd("A.COMMANDSTATUS='W'")
        Else
            Call subAnd("A.CMDSTATUS='W'")
        End If
    ElseIf optHandleYes.Value Then  '�w�B�z
        If intCMType = 2 Then
            Call subAnd("A.COMMANDSTATUS <>'W' And A.COMMANDSTATUS <>'C'")
        Else
            Call subAnd("A.CMDSTATUS IN('C','E')")
        End If
    End If
    If intCMType = 2 Then
        Call subAnd("A.ResvTime Is Not Null " & IIf(cboType.ListIndex <> 4, " And A.HandleTime is null", ""))
    Else
        Call subAnd("A.ProcessingDate Is Not Null " & IIf(cboType.ListIndex <> 4, " And A.HandleTime is null", ""))
    End If
    
    'SO307 �B�z���A
    'W: �ȪA�g�J�ݳB�z,
    'C: GateWay �w����
    'S: GateWay ���榨�\
    'D: �ȪA�t�Ψ���
    'E: GateWay ���榳���~
    'T: �ȪA�t�� TimeOut GateWay �ݭnrollback
    'GT: GatewWay �Pip command time out
    'TR: GateWay rollback  ���\
    'TE:GateWay rollback  ����)
    
    subChoose = strChoose
    strChooseString = "���q�O      :" & gilCompCode.GetDescription & ";" & _
                      "�B�z����    :" & cboType.Text & ";" & _
                      "�]�w�ɶ�    :" & subSpace(gdtUpdTime1.GetValue) & "~" & subSpace(gdtUpdTime1.GetValue) & ";" & _
                      "�w���ɶ�    :" & subSpace(gdtProcessingDate1.GetValue) & "~" & subSpace(gdtProcessingDate2.GetValue) & ";" & _
                      "�]�w�Ҧ�    :" & subSpace(gimCommand.GetDispStr) & ";" & _
                      "�B�z�H��    :" & subSpace(gimExecEntry.GetDispStr) & ";" & _
                      "���O�B�z����:" & IIf(optHandleAll.Value, optHandleAll.Caption, IIf(optHandleNo.Value, optHandleNo.Caption, optHandleYes.Caption)) & _
                      "�h���]    :" & gilReturnCode.GetDescription

    Exit Function
ChkErr:
    ErrSub Me.Name, "subChoose"
End Function
Private Function Get_Cmd_Seq_No() As String
  On Error GoTo ChkErr
    If intCMType = 2 Then
'        Get_Cmd_Seq_No = RPxx(gcnGi.Execute("SELECT " & strCMOwner & "s_so307_cmdseqno.NEXTVAL FROM DUAL").GetString & "")
'        Get_Cmd_Seq_No = Format(gCompCode, "000") & _
'                                Format(RightDate, "yyyymmdd") & _
'                                Format(Get_Cmd_Seq_No, "000000000")
                           
        Get_Cmd_Seq_No = RPxx(gcnGi.Execute("SELECT " & strCMOwner & "S_CMDSEQNO.NEXTVAL FROM DUAL").GetString & "")
        Get_Cmd_Seq_No = Format(gCompCode, "000") & _
                                Format(RightDate, "yyyymmdd") & _
                                Format(Get_Cmd_Seq_No, "000000000")
    Else
        Get_Cmd_Seq_No = RPxx(gcnGi.Execute("SELECT " & strCMOwner & "s_so314_cmdseqno.NEXTVAL FROM DUAL").GetString & "")
        Get_Cmd_Seq_No = Format(Get_Cmd_Seq_No, "00000000")
    End If
  Exit Function
ChkErr:
    ErrSub Me.Name, "Get_Cmd_Seq_No"
End Function

Private Sub DefineRs()
  On Error GoTo ChkErr
    With rsDefTmp
        If .State = adStateOpen Then
            .Close
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Fields.Append ("Choice"), adBSTR, 1, adFldIsNullable
        .Fields.Append ("CommandStatus"), adBSTR, 20, adFldIsNullable
        .Fields.Append ("CustId"), adBSTR, 20, adFldIsNullable
        .Fields.Append ("MsgName"), adBSTR, 30, adFldIsNullable
        .Fields.Append ("SNO"), adBSTR, 20, adFldIsNullable
        .Fields.Append ("ModemMAC"), adBSTR, 30, adFldIsNullable
        .Fields.Append ("ResvTime"), adBSTR, 20, adFldIsNullable
        .Fields.Append ("ExecEntry"), adBSTR, 20, adFldIsNullable
        .Fields.Append ("Exectime"), adBSTR, 20, adFldIsNullable
        .Fields.Append ("RowId"), adBSTR, 50, adFldIsNullable
        .Fields.Append ("MsgCode"), adBSTR, 20, adFldIsNullable
        .Fields.Append ("CmdSeqNo"), adBSTR, 30, adFldIsNullable
        .Fields.Append ("DeviceSNo1"), adBSTR, 20, adFldIsNullable
        '#3917 �W�[HandleNote���
        .Fields.Append ("HandleNote"), adBSTR, 1000
        .Open
    End With
    Exit Sub
ChkErr:
  
  ErrSub Me.Name, "DefineRs"
End Sub
Private Sub ChangeMode()
    
    With frmSO19N0A
        .cmdDelete.Caption = "F10.����"
        .chkAll.Value = 0
        .gilReturnCode.Clear
        .lblReturnCode.ForeColor = vbBlack
        Select Case .cboType.ListIndex
            Case 0
                .optHandleYes.Enabled = True
                .optHandleNo.Enabled = False
                .optHandleAll.Enabled = False
                .optHandleYes.Value = True
                .cmdOK.Enabled = True
                .cmdSetup.Enabled = False
                .gilReturnCode.Enabled = False
                .cmdDelete.Enabled = False
                .chkAll.Enabled = True
            Case 1
                .optHandleNo.Enabled = True
                .optHandleYes.Enabled = False
                .optHandleAll.Enabled = False
                .optHandleNo.Value = True
                .cmdOK.Enabled = False
                .cmdSetup.Enabled = False
                .cmdDelete.Enabled = True
                .gilReturnCode.Enabled = True
                .lblReturnCode.ForeColor = vbRed
                .chkAll.Enabled = False
            Case 2
                .optHandleNo.Enabled = False
                .optHandleYes.Enabled = True
                .optHandleAll.Enabled = False
                .optHandleYes.Value = True
                .cmdOK.Enabled = False
                .cmdSetup.Enabled = True
                .cmdDelete.Enabled = False
                .gilReturnCode.Enabled = False
                .chkAll.Enabled = False
            Case 3
                .optHandleNo.Enabled = False
                .optHandleYes.Enabled = True
                .optHandleAll.Enabled = False
                .optHandleYes.Value = True
                .cmdOK.Enabled = False
                .cmdSetup.Enabled = False
                .cmdDelete.Enabled = True
                .lblReturnCode.ForeColor = vbRed
                .gilReturnCode.Enabled = True
                .chkAll.Enabled = False
                .cmdDelete.Caption = "F10.�����R�O"
            Case 4
                 optHandleNo.Enabled = False
                .optHandleYes.Enabled = False
                .optHandleAll.Enabled = True
                .optHandleAll.Value = True
                .cmdOK.Enabled = False
                .cmdSetup.Enabled = False
                .cmdDelete.Enabled = False
                .gilReturnCode.Enabled = False
                .chkAll.Enabled = False
        End Select
    End With
End Sub

Private Sub tmrAutoUpdate_Timer()
  On Error Resume Next
    Static intAutoCount As Integer
    intAutoCount = intAutoUpd
    If Me.Enabled Then
        If intAutoCount = 0 Then
            intAutoCount = intAutoCount + 1
            intAutoUpd = intAutoCount
            cmdFind.Value = True
            
        Else
            intAutoCount = 0
            intAutoUpd = 0
            cmdOK.Value = True
        End If
    End If
End Sub

Private Sub txtClock_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyY And Shift = 2 Then lngTimer = CLng(txtClock.Text)
End Sub

Private Sub txtCustId_KeyPress(KeyAscii As Integer)
  On Error Resume Next
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 44 Or KeyAscii = 8 Or KeyAscii = 45 Then
        If KeyAscii = 44 Or KeyAscii = 45 Then
            If Asc(Right(" " & txtCustId.Text, 1)) = 44 Or Asc(Right(" " & txtCustId.Text, 1)) = 45 Or txtCustId = "" Then KeyAscii = 9
        End If
        If Not ChkMaxLengthOK(txtCustId) Then
           If Not (KeyAscii = 44 Or KeyAscii = 8 Or KeyAscii = 45) Then KeyAscii = 9
        End If
    Else
        KeyAscii = 9
    End If

End Sub
