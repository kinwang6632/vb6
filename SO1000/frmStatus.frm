VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmStatus 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�Ȥ᪬�A"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   Icon            =   "frmStatus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '�t�ιw�]��
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   3405
      Left            =   0
      TabIndex        =   0
      Tag             =   "3495"
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6006
      QuickMode       =   0   'False
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseCellForeColor=   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "BISP����ɶ�"
      Height          =   180
      Left            =   9225
      TabIndex        =   6
      Top             =   3480
      Width           =   1080
   End
   Begin VB.Label lblBISPInstTime 
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5790
      TabIndex        =   5
      Top             =   3450
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "BISP�˾��ɶ�"
      Height          =   180
      Left            =   4575
      TabIndex        =   4
      Top             =   3480
      Width           =   1080
   End
   Begin VB.Label lblBISPStatus 
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   975
      TabIndex        =   3
      Top             =   3450
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�a�k���
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "BISP���A"
      Height          =   180
      Left            =   165
      TabIndex        =   2
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label lblBISPPRTime 
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   10395
      TabIndex        =   1
      Top             =   3450
      Width           =   1215
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
'  On Error Resume Next
'    Dim intLoop As Integer
'    If Len(Environ("OS")) > 0 Then
'        For intLoop = 0 To 255 Step 17
'            MakeTransparent Me.hwnd, intLoop
'            DoEvents
'        Next
'    End If
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    Dim rsStatus As New ADOR.Recordset
    With rsStatus
        .CursorLocation = adUseClient
'       .Open "SELECT COMPCODE,SERVICETYPE,CUSTSTATUSNAME,WIPNAME1,WIPNAME2,WIPNAME3,AccountOwner FROM " & GetOwner & "SO002 WHERE CUSTID=" & gCustId & " AND COMPCODE=" & gCompCode, gcnGi, adOpenForwardOnly, adLockReadOnly
'       .Open "SELECT A.COMPCODE,B.DESCRIPTION,A.CUSTSTATUSNAME,A.WIPNAME1,A.WIPNAME2,A.WIPNAME3,A.ACCOUNTOWNER,C.INSTCOUNT,C.BALANCE,A.INSTTIME,A.PRTIME,D.DisConnectDate FROM " & GetOwner & "SO002 A," & GetOwner & "CD046 B," & GetOwner & "SO001 C,SO031 D WHERE C.CUSTID=" & gCustId & " AND A.CUSTID=C.CUSTID AND A.COMPCODE=" & gCompCode & " AND B.CODENO=A.SERVICETYPE AND D.CUSTID(+)=A.CUSTID", gcnGi, adOpenForwardOnly, adLockReadOnly
        '#3387 �W�[�������O��� By Kin 2007/09/07
        .Open "SELECT A.COMPCODE,B.DESCRIPTION,A.CUSTSTATUSNAME,C.RiskName,A.WIPNAME1,A.WIPNAME2,A.WIPNAME3,A.ACCOUNTOWNER,A.INSTCOUNT,C.BALANCE,A.BALANCE,A.INSTTIME,A.PRTIME FROM " & GetOwner & "SO002 A," & GetOwner & "CD046 B," & GetOwner & "SO001 C WHERE C.CUSTID=" & gCustId & " AND A.CUSTID=C.CUSTID AND A.COMPCODE=" & gCompCode & " AND B.CODENO=A.SERVICETYPE", gcnGi, adOpenForwardOnly, adLockReadOnly
    End With
    With mFlds
        .Add "COMPCODE", , , , , "���q�O  "
        .Add "DESCRIPTION", , , , , "�A�ȧO  "
        .Add "CUSTSTATUSNAME", , , , , "�Ȥ᪬�A   "
        '#3387 �W�[�������O��� By Kin 2007/09/07
        .Add "RiskName", , , , , "�������O" & Space(12)
        .Add "WIPNAME1", , , , , "�˾������A"
        .Add "WIPNAME2", , , , , "���������A"
        .Add "WIPNAME3", , , , , "��������A"
'        .Add "BALANCE", , , "###,###,###,###", , "��O  ", vbRightJustify
        .Add "BALANCE", , , "###,###,###,###", , "��O     ", vbRightJustify
        .Add "INSTCOUNT", , , , , "�˾��x��", vbCenter
        .Add "INSTTIME", giControlTypeTime, , , , "�˾����         "
        .Add "PRTIME", giControlTypeTime, , , , "������         "
    End With
    With ggrData
        .AllFields = mFlds
        .SetHead
         Set .Recordset = rsStatus
        .Refresh
    End With
    If gcnGi.Execute("SELECT SHOWBISP FROM " & GetOwner & "SO041 WHERE COMPCODE=" & gCompCode & " AND SERVICETYPE LIKE '%" & gServiceType & "%'").GetString(adClipString, , "", "", 0) = 1 Then
        ggrData.Height = 3400
        With frmSO1100BMDI
             lblBISPstatus = .rsSo001!BISPStatus & ""
'             lblBISPInstTime = Format(.rsSo001!BISPInstTime & "", "EE/MM/DD")
'             lblBISPPRTime = Format(.rsSo001!BISPPRTime & "", "EE/MM/DD")
             lblBISPInstTime = GetDTString(.rsSo001!BISPInstTime & "")
             lblBISPPRTime = GetDTString(.rsSo001!BISPPRTime & "")
             
        End With
    Else
        ggrData.Height = 3750
    End If
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Activate"
End Sub

'FOR v3.00
'Schema�W�[����:
'   SO041 : ShowBISP Number(1) Default 0
'   SO001 : BISPStatus Varchar2(20) , BISPInstTime Date , BISPPRTime Date
'�e���B�z����:
'   1. SO7300Bp : �[�@CheckBox����B�z[���BISP���]
'   2. ��SO1100B��[���A]Button���Ȥ᪬�AForm���[���BISP���e. ��SO041.ShowBISP=1��,�h�b�Ȥ᪬�AForm���[Show BISP���A���(BISPStatus , BISPInstTime , BISPPRTime);
'       ��SO041.ShowBISP=0��,�h�^�Ь��ثe�Ȥ᪬�AForm���e
'   3.�Щ�����C�L����[��BISPStatus , BISPInstTime , BISPPRTime�T�����,��UESER�i��C�L�����ɦL�X.

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = 27 Then Unload Me
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    With Me
        .Top = ((Screen.Height - .Height) / 2) - 800
        .Left = (Screen.Width - .Width) / 2
'         If .Picture <> 0 Then Call SetAutoRgn(Me)
'         If Len(Environ("OS")) > 0 Then MakeTransparent Me.hwnd, 0
    End With
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'  On Error GoTo ChkErr
'    If Button = vbLeftButton Then
'        Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
'        Call ReleaseCapture
'    End If
'  Exit Sub
'ChkErr:
'    errHandle Me.Name, "Form_MouseMove"
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
''  On Error Resume Next
''    If Len(Environ("OS")) > 0 Then
''        Dim intLoop As Integer
''        For intLoop = 255 To 0 Step -17
''            MakeTransparent Me.hwnd, intLoop
''            DoEvents
''        Next
''    End If
''    Me.Hide
''    DeleteObject 0
''    ReleaseCapture
''    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
'End Sub

'Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
'  On Error GoTo ChkErr
'    If giFld.FieldName = "SERVICETYPE" Then
'        Select Case Value
'               Case "C"
'                    Value = "CATV"
'               Case "I"
'                    Value = "CAISP"
'        End Select
'    End If
'  Exit Sub
'ChkErr:
'    ErrSub Me.Name, "ggrData_ShowCellData"
'End Sub
