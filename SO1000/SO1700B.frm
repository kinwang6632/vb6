VERSION 5.00
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO1700B 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�A�ȥӧi�s�� [SO1700B]"
   ClientHeight    =   5055
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
   Icon            =   "SO1700B.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNext 
      Caption         =   "�U�@��(&N)"
      Height          =   375
      Left            =   7500
      TabIndex        =   30
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "�W�@��(&P)"
      Height          =   375
      Left            =   6270
      TabIndex        =   29
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   9750
      TabIndex        =   13
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "F2. �s��"
      Height          =   375
      Left            =   390
      TabIndex        =   12
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame fraData 
      Height          =   4395
      Left            =   180
      TabIndex        =   14
      Top             =   90
      Width           =   11535
      Begin VB.TextBox txtNote 
         Height          =   1005
         Left            =   1530
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "SO1700B.frx":0442
         Top             =   3255
         Width           =   3825
      End
      Begin Gi_Time.GiTime gdtHandleTime 
         Height          =   345
         Left            =   7095
         TabIndex        =   10
         Top             =   1305
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
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
      Begin prjGiList.GiList gilHandleName 
         Height          =   315
         Left            =   7095
         TabIndex        =   9
         Top             =   930
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
         FilterStop      =   -1  'True
      End
      Begin VB.TextBox txtHandleNote 
         Height          =   2565
         Left            =   7080
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1710
         Width           =   3915
      End
      Begin Gi_Time.GiTime gdtAcceptTime 
         Height          =   345
         Left            =   1530
         TabIndex        =   6
         Top             =   2850
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
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
      Begin prjGiList.GiList gilAcceptEn 
         Height          =   315
         Left            =   1530
         TabIndex        =   5
         Top             =   2475
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
         FilterStop      =   -1  'True
      End
      Begin prjGiList.GiList gilServiceCode 
         Height          =   315
         Left            =   1530
         TabIndex        =   0
         Top             =   570
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
         FilterStop      =   -1  'True
      End
      Begin prjGiList.GiList gilServiceContent 
         Height          =   315
         Left            =   1530
         TabIndex        =   3
         Top             =   1680
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
      Begin prjGiList.GiList gilProcResultNo 
         Height          =   315
         Left            =   7095
         TabIndex        =   8
         Top             =   570
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
         FilterStop      =   -1  'True
      End
      Begin prjGiList.GiList gilMediaCode 
         Height          =   315
         Left            =   1530
         TabIndex        =   1
         Top             =   930
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
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
         F2Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin prjGiList.GiList gilPromCode 
         Height          =   315
         Left            =   1530
         TabIndex        =   2
         Top             =   1305
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
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
         F2Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin prjGiList.GiList gilContentCode 
         Height          =   315
         Left            =   1530
         TabIndex        =   4
         Top             =   2070
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�ӧi���e�Ӷ�"
         Height          =   195
         Left            =   225
         TabIndex        =   35
         Top             =   2130
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�ӧi���e"
         Height          =   195
         Left            =   225
         TabIndex        =   34
         Top             =   1755
         Width           =   780
      End
      Begin VB.Label lblMediaCode 
         AutoSize        =   -1  'True
         Caption         =   "���дC��"
         Height          =   195
         Left            =   225
         TabIndex        =   33
         Top             =   975
         Width           =   780
      End
      Begin VB.Label lblPromCode 
         AutoSize        =   -1  'True
         Caption         =   "�~�Ȭ���"
         Height          =   195
         Left            =   225
         TabIndex        =   32
         Top             =   1380
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�B�z���G"
         Height          =   195
         Left            =   6120
         TabIndex        =   31
         Top             =   645
         Width           =   780
      End
      Begin VB.Label lblCustId1 
         AutoSize        =   -1  'True
         Caption         =   "�Ȥ�s��:"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   255
         Width           =   825
      End
      Begin VB.Label lblCustName1 
         AutoSize        =   -1  'True
         Caption         =   "�Ȥ�m�W"
         Height          =   195
         Left            =   3165
         TabIndex        =   27
         Top             =   255
         Width           =   780
      End
      Begin VB.Label lbltel1 
         AutoSize        =   -1  'True
         Caption         =   "�q��:"
         Height          =   195
         Left            =   7410
         TabIndex        =   26
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lblCustId 
         AutoSize        =   -1  'True
         Caption         =   "99999999"
         Height          =   195
         Left            =   1335
         TabIndex        =   25
         Top             =   255
         Width           =   720
      End
      Begin VB.Label lblCustName 
         AutoSize        =   -1  'True
         Caption         =   "XXXXXXXXXXXXXXXXX"
         Height          =   195
         Left            =   4095
         TabIndex        =   24
         Top             =   255
         Width           =   2295
      End
      Begin VB.Label lblTel 
         AutoSize        =   -1  'True
         Caption         =   "XXXXXXXXXXXXXXXX"
         Height          =   195
         Left            =   8145
         TabIndex        =   23
         Top             =   240
         Width           =   2160
      End
      Begin VB.Label lblAcceptEn 
         AutoSize        =   -1  'True
         Caption         =   "���z�H��"
         Height          =   195
         Left            =   225
         TabIndex        =   22
         Top             =   2565
         Width           =   780
      End
      Begin VB.Label lblHandleEn 
         AutoSize        =   -1  'True
         Caption         =   "�B�z�H��"
         Height          =   195
         Left            =   6120
         TabIndex        =   21
         Top             =   975
         Width           =   780
      End
      Begin VB.Label lblHandleNote 
         AutoSize        =   -1  'True
         Caption         =   "�B�z����"
         Height          =   195
         Left            =   6120
         TabIndex        =   20
         Top             =   1755
         Width           =   780
      End
      Begin VB.Label lblServiceCode 
         AutoSize        =   -1  'True
         Caption         =   "�ӹq����"
         Height          =   195
         Left            =   225
         TabIndex        =   19
         Top             =   645
         Width           =   780
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "�Ƶ�"
         Height          =   195
         Left            =   225
         TabIndex        =   18
         Top             =   3375
         Width           =   390
      End
      Begin VB.Label lblAcceptTime 
         AutoSize        =   -1  'True
         Caption         =   "���z�ɶ�"
         Height          =   195
         Left            =   225
         TabIndex        =   16
         Top             =   2955
         Width           =   780
      End
      Begin VB.Label lblHandleTime 
         Caption         =   "�B�z�ɶ�"
         Height          =   255
         Left            =   6120
         TabIndex        =   15
         Top             =   1380
         Width           =   780
      End
   End
   Begin VB.Label lblEditMode 
      Alignment       =   2  '�m�����
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  '��u�T�w
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   11040
      TabIndex        =   17
      Top             =   3330
      Width           =   675
   End
End
Attribute VB_Name = "frmSO1700B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsSo1700B As New ADODB.Recordset
Private lngEditMode As giEditModeEnu
'�O�_�ϥ��ܼƱ���
Private Const blnYesNo = False
'�����W�h�ܼ�
Private strUserName As String

Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
End Sub

' �ۭq�ݩʡGEditMode
' �O���ثe�b�s��B�s�W���˵��Ҧ�
' giEditModeEnu(�ۭq�C�|�ȡA�]�w��Sys_Lib)
Public Property Get EditMode() As giEditModeEnu
   On Error GoTo ChkErr
   ' ���ثe�s��Ҧ�
   EditMode = lngEditMode
   Exit Property
ChkErr:
   Call ErrSub(Me.Name, "Get EditMode")
End Property

Public Property Let EditMode(ByVal vNewValue As giEditModeEnu)
  On Error GoTo ChkErr
    '�]�w�s��Ҧ�
    lngEditMode = vNewValue
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "Let EditMode")
End Property


'�ھڶǨӤ��s��Ҧ�, �]�w�U�����ݩ�
Private Sub ChangeMode(ByVal lngMode As giEditModeEnu)
 On Error GoTo ChkErr
   Dim blnFlag As Boolean  ' �O���O�_�b����s���Ҧ��A
   lngEditMode = lngMode
   Select Case lngMode
      Case giEditModeView
         blnFlag = True
         cmdSave.Enabled = False
         lblEditMode.Caption = "���"
         cmdCancel.Caption = "����(&X)"
      Case giEditModeEdit
         blnFlag = False
         cmdSave.Enabled = True
         lblEditMode.Caption = "�ק�"
         cmdCancel.Caption = "����(&X)"
   End Select
   fraData.Enabled = Not blnFlag
   cmdPrev.Visible = blnFlag
   cmdNext.Visible = blnFlag
 Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "ChangeMode")
End Sub

Private Sub SetGrid()
  On Error GoTo ChkErr
    '�]�w�ù��W�U��Gird����
    '    gilServiceCode  <ServiceCode, ServiceName>  �ӧi��]�N�X�ACD008(CodeNo�ADescription)
    '    gilAcceptEn <AcceptEn, AcceptName>  ���z�H���N���ACM003(EmpNo�AEmpName)
    '    gilHandleEn <HandleEn, HandleName>  �B�z�H���N�X�ACM003(EmpNo�AEmpName)
    '�ӧi��]
    gilServiceCode.SelectRefNo
    
    SetLst gilServiceCode, "CodeNo", "Description", 3, 12, , , "CD008", , True
    SetLst gilServiceContent, "CodeNo", "Description", 3, 12, , , "CD008C"

    '���z�H��
    SetLst gilAcceptEn, "EmpNo", "EmpName", 10, 20, , , "CM003", " WHERE COMPCODE=" & garyGi(9)
    
    '�B�z���G
    SetLst gilProcResultNo, "CodeNo", "Description", 3, 12, , , "CD008B"
    '�B�z�H��
    SetLst gilHandleName, "EmpNo", "EmpName", 10, 20, , , "CM003", " WHERE COMPCODE=" & garyGi(9)
  
    SetLst gilMediaCode, "CodeNo", "Description", 3, 12, , , "CD009", " WHERE RefNo=4 AND Nvl(SERVICETYPE,'" & gServiceType & "' ) ='" & gServiceType & "'"
    SetLst gilPromCode, "CodeNo", "Description", 3, 12, , , "CD042", " WHERE COMPCODE=" & garyGi(9)
  
  Exit Sub
ChkErr:
    ErrSub Me.Name, "SetGrid"
End Sub

Private Sub cmdNext_Click()
  On Error GoTo ChkErr
    If Not rsSo1700B.EOF Then
        rsSo1700B.MoveNext
        If Not rsSo1700B.EOF Then
            Call RcdToScr
        Else
            MsgBox "�w���ɧ��F�I", vbExclamation, "�T���I"
            rsSo1700B.MoveLast
        End If
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdNext_Click"
End Sub

Private Sub cmdPrev_Click()
  On Error GoTo ChkErr
    If Not rsSo1700B.BOF Then
        rsSo1700B.MovePrevious
        If Not rsSo1700B.BOF Then
            Call RcdToScr
        Else
            MsgBox "�w�����Y�F�I", vbExclamation, "�T���I"
            rsSo1700B.MoveFirst
        End If
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdPrev_Click"
End Sub

Private Sub cmdSave_Click()
  On Error GoTo ChkErr
    If IsDataOk Then
        Call ScrToRcd
        Unload Me
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdSave_Click")
End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = vbKeyF2 Then
        If Not ChkGiList(KeyCode) Then Exit Sub
        Call cmdSave_Click
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    If Crm Then
        Me.Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2)
    Else
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End If
    
    '#3495 �W�[�ӧi���e�Ӷ� By Kin 2008/02/19
    SetLst gilContentCode, "CodeNo", "Description", 3, 12, , , "CD008E"
    
    Call SetGrid
    Call ChangeMode(lngEditMode)
    Call RcdToScr
    Call gilServiceContent.SetCodeNo(rsSo1700B("ServDescCode").Value & "")
    Call gilServiceContent.SetDescription(rsSo1700B("ServDescName").Value & "")
    
    
    
    'If lngEditMode = giEditModeInsert Then gilServiceCode_Change
    
Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Public Property Let GetRecordSet(ByVal rsRcdNewValue As ADODB.Recordset)
  On Error GoTo ChkErr
    Set rsSo1700B = rsRcdNewValue
  Exit Property
ChkErr:
    Call ErrSub(Me.Name, "GetRecordSet")
End Property

Private Sub RcdToScr()
  On Error GoTo ChkErr
    '�ӧi��] SERVICECODE,SERVICENAME
    Call gilServiceCode.SetCodeNo(rsSo1700B("ServiceCode").Value & "")
    Call gilServiceCode.SetDescription(rsSo1700B("ServiceName").Value & "")
    '�ӧi���e ServDescCode,ServDescName
    Call gilServiceContent.SetCodeNo(rsSo1700B("ServDescCode").Value & "")
    Call gilServiceContent.SetDescription(rsSo1700B("ServDescName").Value & "")
    '���z�H��
    'AcceptEn,AcceptName
    Call gilAcceptEn.SetCodeNo(rsSo1700B("AcceptEn").Value & "")
    Call gilAcceptEn.SetDescription(rsSo1700B("AcceptName").Value & "")
    
    gilMediaCode.SetCodeNo rsSo1700B("MediaCode") & ""
    gilMediaCode.SetDescription rsSo1700B("MediaName") & ""
    
    gilPromCode.SetCodeNo rsSo1700B("PromCode") & ""
    gilPromCode.SetDescription rsSo1700B("PromName") & ""
    
    '���z�ɶ�
    gdtAcceptTime.Text = rsSo1700B("AcceptTime").Value & ""
    txtNote.Text = rsSo1700B("Note").Value & ""
    '�B�z���G
    Call gilProcResultNo.SetCodeNo(rsSo1700B("ProcResultNo").Value & "")
    Call gilProcResultNo.SetDescription(rsSo1700B("ProcResultName").Value & "")
    '�B�z�H��
    Call gilHandleName.SetCodeNo(rsSo1700B("HandleEn").Value & "")
    Call gilHandleName.SetDescription(rsSo1700B("HandleName").Value & "")
    '�B�z�ɶ�
    gdtHandleTime.Text = rsSo1700B("HANDLETIME").Value & ""
    '�B�z����
    txtHandleNote.Text = rsSo1700B("HandleNote").Value & ""
    '�Ȥ�s��,�m�W
    lblCustId.Caption = rsSo1700B("CustID").Value & ""
    lblCustName.Caption = rsSo1700B("CustName").Value & ""
    lblTel.Caption = rsSo1700B("Tel1").Value & ""
    
    '*******************************************************************************
    '#3495 �W�[�ӧi���e�Ӷ� By Kin 2008/02/19
    If Not IsNull(rsSo1700B("ContentCode")) Then
        Call gilContentCode.SetCodeNo(rsSo1700B("ContentCode").Value & "")
    End If
    If Not IsNull(rsSo1700B("ContentName")) Then
        Call gilContentCode.SetDescription(rsSo1700B("ContentName").Value & "")
    End If
    '********************************************************************************
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "RcdToScr")
End Sub

Private Sub ScrToRcd()
  On Error GoTo ChkErr
    If Not rsSo1700B.EOF Then
        rsSo1700B("ServiceCode").Value = gilServiceCode.GetCodeNo
        rsSo1700B("ServiceName").Value = gilServiceCode.GetDescription
        rsSo1700B("ServDescCode").Value = IIf(gilServiceContent.GetCodeNo = "", Null, gilServiceContent.GetCodeNo)
        rsSo1700B("ServDescName").Value = IIf(gilServiceContent.GetDescription = "", Null, gilServiceContent.GetDescription)
         
        rsSo1700B("MediaCode") = SolveNull(gilMediaCode.GetCodeNo & "")
        rsSo1700B("MediaName") = gilMediaCode.GetDescription & ""
        rsSo1700B("PromCode") = SolveNull(gilPromCode.GetCodeNo & "")
        rsSo1700B("PromName") = gilPromCode.GetDescription & ""
        
        rsSo1700B("AcceptEn").Value = gilAcceptEn.GetCodeNo
        rsSo1700B("AcceptName").Value = gilAcceptEn.GetDescription
        rsSo1700B("AcceptTime").Value = Format(gdtAcceptTime.Text, "YYYY/MM/DD HH:MM:SS")
        rsSo1700B("Note").Value = txtNote.Text
        rsSo1700B("ProcResultNo").Value = IIf(gilProcResultNo.GetCodeNo = "", Null, gilProcResultNo.GetCodeNo)
        rsSo1700B("ProcResultName").Value = IIf(gilProcResultNo.GetDescription = "", Null, gilProcResultNo.GetDescription)
        rsSo1700B("HandleEn").Value = gilHandleName.GetCodeNo
        rsSo1700B("HandleName").Value = gilHandleName.GetDescription
        rsSo1700B("HANDLETIME").Value = IIf(gdtHandleTime.Text = "", Null, Format(gdtHandleTime.Text, "YYYY/MM/DD HH:MM:SS"))
        rsSo1700B("HandleNote").Value = IIf(txtHandleNote.Text = "", Null, txtHandleNote.Text)
        rsSo1700B("UpdEn").Value = IIf(blnYesNo, strUserName, garyGi(1))
        rsSo1700B("UpdTime").Value = GetDTString(RightNow)
        
        '***********************************************************************
        '#3495�W�[�ӧi���e�Ӷ� By Kin 2008/02/19
        If gilContentCode.GetCodeNo <> "" Then
            rsSo1700B("ContentCode").Value = gilContentCode.GetCodeNo
            rsSo1700B("ContentName").Value = gilContentCode.GetDescription
        End If
        '************************************************************************
        rsSo1700B.Update
    End If
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "ScrToRcd")
End Sub

Private Function IsDataOk() As Boolean
'���n���: �ӧi��]�A���z�H���A���z�ɶ�
  On Error GoTo ChkErr
    IsDataOk = False
    If gdtAcceptTime.Text = "" Then
        gdtAcceptTime.SetFocus
        GoTo ErrMsg
    End If
    If gilAcceptEn.GetCodeNo = "" Then
        gilAcceptEn.SetFocus
        GoTo ErrMsg
    End If
    If gilServiceCode.GetCodeNo = "" Then
        gilServiceCode.SetFocus
        GoTo ErrMsg
    End If
    If gilServiceCode.GetRefNo = "9" Then
        If gilMediaCode.GetCodeNo = "" Or gilMediaCode.GetDescription = "" Then
            gilMediaCode.SetFocus
            GoTo ErrMsg
        End If
        If gilPromCode.GetCodeNo = "" Or gilPromCode.GetDescription = "" Then
            gilPromCode.SetFocus
            GoTo ErrMsg
        End If
    End If
    IsDataOk = True
  Exit Function
ErrMsg:
    MsgBox "�������n���I", vbExclamation, "�T���I"
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsDataOK")
End Function

Public Property Let uUserName(ByVal vNewValue As Variant)
  On Error GoTo ChkErr
    strUserName = vNewValue
  Exit Property
ChkErr:
    ErrSub Me.Name, "Let uUserName"
End Property

Private Sub gilServiceCode_Change()
  On Error GoTo ChkErr
    With gilServiceContent
        If Len(Trim(gilServiceCode.GetCodeNo)) > 0 Then
'            SetLst gilServiceContent, "CodeNo", "Description", 3, 12, , , "CD008C", " WHERE ServiceCode=" & gilServiceCode.GetCodeNo
            SetgiList gilServiceContent, "CodeNo", "Description", "CD008A", , , , , , , " Where (ServiceType='" & rsSo1700B("ServiceType") & "' Or ServiceType is Null) And CodeNo in (" & _
                            " Select CodeNo From " & GetOwner & "CD008C Where ServiceCode=" & gilServiceCode.GetCodeNo & " And (ServiceType='" & _
                            rsSo1700B("ServiceType") & "' Or ServiceType is Null))", True

            If gilServiceCode.GetRefNo = "9" Then
                gilMediaCode.Enabled = True
                gilPromCode.Enabled = True
            Else
                gilMediaCode.SetCodeNo ""
                gilMediaCode.SetDescription ""
                gilPromCode.SetCodeNo ""
                gilPromCode.SetDescription ""
                gilMediaCode.Enabled = False
                gilPromCode.Enabled = False
            End If
            If .GetListCount > 0 Then
               .Enabled = True
               .ListIndex = 1
            Else
               .SetCodeNo ""
               .SetDescription ""
               .Enabled = False
            End If
        Else
            gilMediaCode.Enabled = False
            gilPromCode.Enabled = False
           .Enabled = False
           SetgiList gilServiceCode, "CodeNo", "Description", "CD008A", , , , , , , " Where 1=0"
        End If
    End With
    '**************************************************************************************************************************
    '#3495 �ӧi���e�Ӷ� By Kin 2008/02/19
    If gilServiceContent.GetCodeNo <> "" Then
        gilContentCode.Filter = ""
        'SetLst gilContentCode, "CodeNo", "Description", 3, 12, , , "CD008E", " WHERE ContentCode=" & gilServiceContent.GetCodeNo
         SetLst gilContentCode, "CodeNo", "Description", 3, 12, , , "CD008E", " WHERE ServiceType='" & rsSo1700B("ServiceType") & "' Or ServiceType is Null And " & _
                                                                    "Exists (Select * From CD008D Where CD008E.CodeNo=CD008D.CodeNo And " & _
                                                                                "CD008E.ServiceType=CD008D.ServiceType And CD008D.StopFlag<>1 or CD008E.ServiceType is Null)"
                                                                             
    Else
        gilContentCode.Filter = ""
        SetLst gilContentCode, "CodeNo", "Description", 3, 12, , , "CD008E", "Where 1=0"
    End If
    '**************************************************************************************************************************
  Exit Sub
ChkErr:
    ErrSub Me.Name, "gilServiceCode_Change"
End Sub

Private Sub gilServiceContent_DropDown()
    gilServiceCode_Change
End Sub

Private Sub txtHandleNote_Change()
  On Error Resume Next
    CML
End Sub

'Private Sub gilServiceCode_Change()
'  Dim rs As New ADODB.Recordset
'  Dim strWHERE As String
'
'    If gilServiceCode.GetCodeNo = "" Then gilServiceCode.SetFocus: GoTo 99
'    strWHERE = " WHERE ServiceCode=" & gilServiceCode.GetCodeNo
'    SetLst gilServiceContent, "CodeNo", "Description", 3, 12, , , "CD008C", strWHERE
'    gilServiceContent.SetCodeNo ""
'    gilServiceContent.SetDescription ""
'  Exit Sub
'99:
'   MsgBox "�������n���I", vbExclamation, "�T���I"
'  Exit Sub
'ChkErr:
'  ErrSub Me.Name, "Let uUserName"
'End Sub


Private Sub txtNote_Change()
  On Error Resume Next
    CML
End Sub
