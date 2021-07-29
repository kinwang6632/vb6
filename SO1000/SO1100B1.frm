VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1100B1 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "������J�u�����s�� [SO1100B1]"
   ClientHeight    =   3510
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
   Icon            =   "SO1100B1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   2925
      Left            =   120
      TabIndex        =   0
      Top             =   450
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   5159
      QuickMode       =   0   'False
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
   Begin MSComctlLib.TabStrip tabData 
      Height          =   3315
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   5847
      Style           =   1
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   6773
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&A. �� / �[ / �_��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&B. ����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&C. �� / �� / ����"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSO1100B1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmer: Power Hammer
' Start Date: 2005/11/23
' Last Modify: 2005/11/23

Private rs789A As New ADODB.Recordset

Private Sub Form_Activate()
  On Error Resume Next
    FrmOnTop Me
End Sub

Public Sub Set_Normal_Frm_Act(Optional frmAct As Form = Nothing)
  On Error Resume Next
    Dim frmMain As Form
    Set frmMain = frmSO1100BMDI
    If frmAct Is Nothing Or IsMissing(frmAct) Then Set frmAct = Screen.ActiveForm
    With frmAct
        If Crm Then
            .Move frmMain.Left + ((frmMain.Width - .Width) / 2), frmMain.Top + ((frmMain.Height - .Height) / 2)
        Else
            .Move (Screen.Width - .Width) / 2, (Screen.Height - .Height) / 2
        End If
    End With
    Set ObjActiveForm = frmMain
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Set_Normal_Frm_Act Me
    OpenData 1
  Exit Sub
ChkErr:
   ErrSub Name, "Form_Load"
End Sub

Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
  On Error GoTo ChkErr
    If giFld.FieldName = "Type" Then Value = IIf(Val(Len(Value & "")) = 0, "���鵲", "�w�鵲")
  Exit Sub
ChkErr:
    ErrSub Name, "ggrData_ShowCellData"
End Sub

Private Sub tabData_Click()
  On Error GoTo ChkErr
    Dim strQry As String
    With tabData
            Select Case True
                    Case .Tabs(1).Selected
                            OpenData 1
                    Case .Tabs(2).Selected
                            OpenData 2
                    Case .Tabs(3).Selected
                            OpenData 3
            End Select
    End With
  Exit Sub
ChkErr:
    ErrSub Name, "tabData_Click"
End Sub

Private Sub OpenData(bytIdx As Byte)
  On Error GoTo ChkErr
    Dim strQry As String
    Select Case bytIdx
                Case 1
                        strQry = "SELECT A.CLSTIME AS TYPE,A.*,B.DESCRIPTION " & _
                                        " FROM " & GetOwner & "SO007A A," & GetOwner & "CD039 B " & _
                                        " WHERE A.CUSTID = " & gCustId & _
                                        " AND A.SERVICETYPE='" & gServiceType & "'" & _
                                        " AND A.COMPCODE=" & gCompCode & _
                                        " AND B.CODENO=A.COMPCODE " & _
                                        " ORDER BY TYPE DESC,FINTIME DESC,SIGNDATE DESC,RESVTIME DESC"
                        GetRS rs789A, strQry, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly, "( SO007A )", "tabData_Click"
                        SO007Grd
                Case 2
                        strQry = "SELECT A.CLSTIME AS TYPE,A.*,B.DESCRIPTION " & _
                                        " FROM " & GetOwner & "SO008A A," & GetOwner & "CD039 B " & _
                                        " WHERE A.CUSTID = " & gCustId & _
                                        " AND A.SERVICETYPE='" & gServiceType & "'" & _
                                        " AND A.COMPCODE=" & gCompCode & _
                                        " AND B.CODENO=A.COMPCODE " & _
                                        " ORDER BY TYPE DESC,FINTIME DESC,SIGNDATE DESC,RESVTIME DESC"
                        GetRS rs789A, strQry, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly, "( SO008A )", "tabData_Click"
                        SO008Grd
                Case 3
                        strQry = "SELECT A.CLSTIME AS TYPE,A.*,B.DESCRIPTION " & _
                                        " FROM " & GetOwner & "SO009A A," & GetOwner & "CD039 B " & _
                                        " WHERE A.CUSTID = " & gCustId & _
                                        " AND A.SERVICETYPE='" & gServiceType & "'" & _
                                        " AND A.COMPCODE=" & gCompCode & _
                                        " AND B.CODENO=A.COMPCODE " & _
                                        " ORDER BY TYPE DESC,FINTIME DESC,SIGNDATE DESC,RESVTIME"
                        GetRS rs789A, strQry, gcnGi, adUseClient, adOpenForwardOnly, adLockReadOnly, "( SO009A )", "tabData_Click"
                        SO009Grd
    End Select
  Exit Sub
ChkErr:
    ErrSub Name, "OpenData"
End Sub

Private Sub SO007Grd()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    With mFlds
            .Add "Description", , , , , "��J���q�O"
            .Add "Type", , , , , "����" & Space(3)
            .Add "SNO", , , , , "���u�渹" & Space(10)
            .Add "InstName", , , , , "�˾����O" & Space(4)
            .Add "AcceptTime", giControlTypeTime, , , , "���z�ɶ�" & Space(8)
            .Add "ResvTime", giControlTypeTime, , , , "�w���ɶ�" & Space(8)
            .Add "FinTime", giControlTypeTime, , , , "���u�ɶ�" & Space(8)
            .Add "ReturnName", , , , , "�h���]" & Space(4)
            .Add "CallOkTime", giControlTypeTime, , , , "�^�����u�ɶ�"
            .Add "AcceptName", , , , , "���z�H��"
            .Add "ClsTime", giControlTypeTime, , , , "�鵲�b�ɶ�" & Space(8)
            .Add "SignName", , , , , "ñ���H��"
            .Add "OrderNo", , , , , "�q��渹" & Space(10)
            .Add "Note", , , , , "�Ƶ�" & Space(8)
            .Add "UpdEn", , , , , "���ʤH��"
            .Add "UpdTime", , , , , "���ʮɶ�" & Space(8)
    End With
    RefreshGrd mFlds
Exit Sub
ChkErr:
    ErrSub Name, "SO007Grd"
End Sub

Private Sub SO008Grd()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    With mFlds
            .Add "Description", , , , , "��J���q�O"
            .Add "Type", , , , , "����" & Space(3)
            .Add "SNO", , , , , "���u�渹" & Space(10)
            .Add "ServiceName", , , , , "�A�ȥӧi" & Space(4)
            .Add "AcceptTime", giControlTypeTime, , , , "���z�ɶ�" & Space(8)
            .Add "ResvTime", giControlTypeTime, , , , "�w���ɶ�" & Space(8)
            .Add "FinTime", giControlTypeTime, , , , "���u�ɶ�" & Space(8)
            .Add "ReturnName", , , , , "�h���]" & Space(4)
            .Add "CallOkTime", giControlTypeTime, , , , "�^�����u�ɶ�"
            .Add "AcceptName", , , , , "���z�H��"
            .Add "ClsTime", giControlTypeTime, , , , "�鵲�b�ɶ�" & Space(8)
            .Add "SignName", , , , , "ñ���H��"
            .Add "MFName1", , , , , "�G�٭�]�@"
            .Add "MFName2", , , , , "�G�٭�]�G"
            .Add "Note", , , , , "�Ƶ�" & Space(8)
            .Add "UpdTime", , , , , "���ʮɶ�" & Space(8)
            .Add "UpdEn", , , , , "���ʤH��"
    End With
    RefreshGrd mFlds
  Exit Sub
ChkErr:
    ErrSub Name, "SO008Grd"
End Sub

Private Sub SO009Grd()
  On Error GoTo ChkErr
    Dim mFlds As New GiGridFlds
    With mFlds
            .Add "Description", , , , , "��J���q�O"
            .Add "Type", , , , , " ����" & Space(3)
            .Add "SNO", , , , , "���u�渹" & Space(10)
            .Add "PrName", , , , , "���u���O" & Space(4)
            .Add "ReasonName", , , , , "���u��]" & Space(4)
            .Add "AcceptTime", giControlTypeTime, , , , "���z�ɶ�" & Space(12)
            .Add "ResvTime", giControlTypeTime, , , , "�w���ɶ�" & Space(12)
            .Add "FinTime", giControlTypeTime, , , , "���u�ɶ�" & Space(12)
            .Add "ReturnName", , , , , "�h���]" & Space(4)
            .Add "CallOkTime", giControlTypeTime, , , , "�^�����u�ɶ�"
            .Add "AcceptName", , , , , "���z�H��"
            .Add "ClsTime", giControlTypeTime, , , , "�鵲�b�ɶ�" & Space(12)
            .Add "SignName", , , , , "ñ���H��"
            .Add "Note", , , , , "�Ƶ�" & Space(8)
            .Add "UpdTime", giControlTypeTime, , , , "���ʮɶ�" & Space(12)
            .Add "UpdEn", , , , , "���ʤH��"
    End With
    RefreshGrd mFlds
  Exit Sub
ChkErr:
    ErrSub Name, "SO009Grd"
End Sub

Private Sub RefreshGrd(FldObj As Object)
  On Error GoTo ChkErr
    With ggrData
'        LockWindowUpdate .GetHwnd
        .AllFields = FldObj
        .SetHead
        Set .Recordset = rs789A
'        LockWindowUpdate 0
'        UpdateWindow .GetHwnd
        .Refresh
    End With
  Exit Sub
ChkErr:
    ErrSub Name, "RefreshGrd"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    Set ObjActiveForm = frmSO1100BMDI
    CloseRS rs789A
    ReleaseCOM frmSO1100B1
    Set frmSO1100B1 = Nothing
End Sub

