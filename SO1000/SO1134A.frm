VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO1134A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�h�O�է����s�� [SO1134A]"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   1530
   ClientWidth     =   11880
   Icon            =   "SO1134A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   5100
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   8996
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
Attribute VB_Name = "frmSO1134A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmer: Power Hammer
' Last Modify: 2006/03/06

Option Explicit

Private rs As New ADODB.Recordset

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    If Crm Then
        Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Height) / 2)
    Else
        Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    End If
    OpenData
  Exit Sub
ChkErr:
    ErrSub Name, "Form_Load"
End Sub

Private Sub OpenData()
  On Error GoTo ChkErr
  
    'select���A�B���z����B�է��]�B�h�O�է怜�O���ئW�١B��h���B�B��h����B���ơB���İ_�l����B���ĺI����
    '   �B���z�H���B���z�H�B�~�����B���է蠟�渹�B���է蠟�����B���է蠟���O���ئW�١B�f�֤H���m�W1
    '   �B�f�֤H�B�~����1�B�֭���1�B�f�֤H���m�W2�B�f�֤H�B�~����2�B�֭���2�B�h�O�է怜�O�渹
    '   from so171 where CustId=[�ӫȤ�s��] and ServiceType=[�ثe�A�����O]
    '   order by AcceptTime desc
    
    GetRS rs, "SELECT * FROM " & GetOwner & "SO171" & _
                    " WHERE CUSTID=" & gCustId & _
                    " AND SERVICETYPE='" & gServiceType & "'" & _
                    " ORDER BY ACCEPTTIME DESC", _
                    gcnGi, adUseClient, adOpenStatic, adLockReadOnly
    GrdFmt
  Exit Sub
ChkErr:
    ErrSub Name, "OpenData"
End Sub

Private Sub GrdFmt()
  On Error GoTo ChkErr
  
    ' ���A�B���z����B�է��]�B�h�O�է怜�O���ئW�١B��h���B�B��h����B���ơB���İ_�l����B���ĺI����
    '�B���z�H���B���z�H�B�~�����B���է蠟�渹�B���է蠟�����B���է蠟���O���ئW�١B�f�֤H���m�W1�B�f�֤H�B�~����1
    '�B�֭���1�B�f�֤H���m�W2�B�f�֤H�B�~����2�B�֭���2�B�h�O�է怜�O�渹
    Dim mFlds As New GiGridFlds
    With mFlds
            .Add "PSTATUS", , , , , "���A", vbRightJustify
            .Add "ACCEPTTIME", giControlTypeDate, , , , "  ���z�ɶ�  ", vbLeftJustify
            .Add "STNAME", , , , , "�է��]", vbLeftJustify
            .Add "CITEMNAME", , , , , "�h�O�է怜�O���ئW��", vbLeftJustify
            .Add "REALAMT", , , , , "��h���B", vbRightJustify
            .Add "REALDATE", giControlTypeDate, , , , "  ��h���  ", vbLeftJustify
            .Add "REALPERIOD", , , , , "����", vbRightJustify
            .Add "REALSTARTDATE", giControlTypeDate, , , , "���İ_�l���", vbLeftJustify
            .Add "REALSTOPDATE", giControlTypeDate, , , , "���ĺI����", vbLeftJustify
            .Add "ACCEPTNAME", , , , , "���z�H���m�W", vbLeftJustify
            .Add "NOTES1", , , , , "���z�H�B�~����", vbLeftJustify
            .Add "SBILLNO", , , , , "���է蠟�渹", vbLeftJustify
            .Add "SITEM", , , , , "���է蠟����", vbRightJustify
            .Add "SCITEMNAME", , , , , "���է蠟���O���ئW��", vbLeftJustify
            .Add "PROCESSNAME1", , , , , "�f�֤H���m�W1", vbLeftJustify
            .Add "NOTES2", , , , , "�f�֤H�B�~����1", vbLeftJustify
            .Add "PROCESSDATE1", giControlTypeDate, , , , " �֭���1 ", vbLeftJustify
            .Add "PROCESSNAME2", , , , , "�f�֤H���m�W2", vbLeftJustify
            .Add "NOTES3", , , , , "�f�֤H�B�~����2", vbLeftJustify
            .Add "PROCESSDATE2", giControlTypeDate, , , , " �֭���2 ", vbLeftJustify
            .Add "BILLNO", , , , , "�h�O�է怜�O�渹", vbLeftJustify
    End With
    With ggrData
            .AllFields = mFlds
            Set .Recordset = rs
            .Refresh
    End With
  Exit Sub
ChkErr:
    ErrSub Name, "GrdFmt"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
    CloseRS rs
    Rlx rs
End Sub
