VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO1100D 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�Ȥ���/�Ȥ�A��--�Ȥ���P [SO1100D]"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "SO1100D.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1290
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4440
      Begin Gi_Date.GiDate gdaDelDate 
         Height          =   285
         Left            =   1470
         TabIndex        =   2
         Top             =   315
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
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
      Begin prjGiList.GiList gilDelete 
         Height          =   285
         Left            =   1470
         TabIndex        =   4
         Top             =   765
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth1       =   405
         FldWidth2       =   1620
         F2Corresponding =   -1  'True
         FilterStop      =   -1  'True
      End
      Begin VB.Label lblDelete 
         Caption         =   "���P��]"
         Height          =   195
         Left            =   660
         TabIndex        =   3
         Top             =   810
         Width           =   855
      End
      Begin VB.Label lblDelDate 
         Caption         =   "���P���"
         Height          =   225
         Left            =   660
         TabIndex        =   1
         Top             =   345
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "����(&X)"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1620
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�T�w"
      Height          =   375
      Left            =   450
      TabIndex        =   5
      Top             =   1620
      Width           =   1215
   End
End
Attribute VB_Name = "frmSO1100D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmer: Power Hammer
' Last Modify: 2000/12/30

Private Sub cmdCancel_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
    ErrSub Me.Name, "cmdCancel_Click"
End Sub

'Private Sub cmdOK_Click()
'  On Error GoTo CnErr
'    Dim cn2 As New ADODB.Connection
'    If Not MustExist(gdaDelDate, 1, "���P���") Then Exit Sub
'    If Not MustExist(gilDelete, 2, "���P��]") Then Exit Sub
'    With cn2
'        .CursorLocation = adUseServer
'        .Open gcnGi.ConnectionString
'        .BeginTrans
'    End With
'  On Error GoTo ProcErr
'    '(1) ��s�U�ؤ�Ʋέp���:
'    '.�Y�ӫȤ᪬�A�����` (1), �h:
'    With frmSO1100BMDI
'         If .rsSo002!CustStatusCode = 1 Then
'             '1.  �N��D�N�X��(CD017)��, �ӫȤ�˾��a�}�ҹ�������D, �N��D�w�ˤ�ƴ�1
'             If ExecuteSQL("UPDATE " & GetOwner & "CD017 SET INSTCNT = INSTCNT - 1 WHERE CODENO = " & .rsSo001!InstAddrNO, cn2) < giOK Then
'                 cn2.RollbackTrans
'                 cn2.Close
'                 Set cn2 = Nothing
'                 Exit Sub
'             End If
''             If ExecuteSQL("UPDATE " & GetOwner & "CD017 SET INSTCNT = INSTCNT - 1 WHERE CODENO = " & .rsSo001!InstAddrNO, cn2) < giOK Then
''                 cn2.RollbackTrans
''                 cn2.Close
''                 Set cn2 = Nothing
''                 Exit Sub
''             End If
'             '2.  �Y�ӫȤᬰ�j�Ӥ�(�ӵ��Ȥ��ƪ�MduId��즳��)
'             If Not IsNull(.rsSo001!MduId) Then
'                 If .rsSo001!MduId <> 0 Then ' �h�N�j�Ӹ���ɪ��Ӥj�Ӫ����Ĥ�ƴ�1
'                     If ExecuteSQL("UPDATE " & GetOwner & "SO017 SET INSTCNT = INSTCNT - 1 WHERE MDUID = '" & .rsSo001!MduId & "'", cn2) < giOK Then
'                         cn2.RollbackTrans
'                         cn2.Close
'                         Set cn2 = Nothing
'                         Exit Sub
'                     End If
'                 End If
'             End If
'         ' �Y�ӫȤ᪬�A���w���(3)�Τw����(2)
'         ElseIf .rsSo002!CustStatusCode = 3 Or .rsSo002!CustStatusCode = 2 Then              ' �B�ӫȤᬰ�j�Ӥ�
'             If Not IsNull(.rsSo001!MduId) Then
'                 If .rsSo001!MduId <> 0 Then ' �h�N�j�Ӹ���ɪ��Ӥj�Ӫ��j�Ӱ�����ݩ��ƴ�1
'                     If ExecuteSQL("UPDATE " & GetOwner & "SO017 SET UNINSTCNT = UNINSTCNT - 1 WHERE MDUID = '" & .rsSo001!MduId & "'", cn2) < giOK Then
'                         cn2.RollbackTrans
'                         cn2.Close
'                         Set cn2 = Nothing
'                         Exit Sub
'                     End If
'                 End If
'             End If
'         End If
'
'         cn2.Execute "INSERT INTO " & GetOwner & "SO101 (CUSTID,CUSTNAME,CUSTSTATUSCODE,CUSTSTATUSNAME,DELDATE,DELCODE,DELNAME,CUSTSTATUSCODEB,CUSTSTATUSNAMEB,UPDTIME,UPDEN,MODEID,COMPCODE,SERVICETYPE) VALUES " & _
'                     "(" & gCustId & ",'" & .rsSo001!Custname & "'," & _
'                     .rsSo002!CustStatusCode & ",'" & _
'                     .rsSo002!CustStatusName & "'," & _
'                      GetNullString(gdaDelDate.Text, giDateV) & "," & _
'                      gilDelete.GetCodeNo2 & ",'" & _
'                      gilDelete.GetDescription2 & "',4,'" & _
'                      GetRsValue("select Description FROM " & GetOwner & "CD035 WHERE CodeNo=4", gcnGi) & "','" & _
'                      GetDTString(RightNow) & "','" & _
'                      garyGi(1) & "',4," & _
'                      gCompCode & ",'" & _
'                      gServiceType & "')"
'
'         .rsSo002!DelDate = SolveDate(gdaDelDate.GetValue) & ""
'         .rsSo002!DelCode = gilDelete.GetCodeNo2 & ""
'         .rsSo002!DelName = gilDelete.GetDescription2 & ""
'         .rsSo002!OldStatusCode = Val(.rsSo002!CustStatusCode)
'         .rsSo002!CustStatusCode = 4
'         .rsSo002!CustStatusName = GetRsValue("select Description FROM " & GetOwner & "CD035 WHERE CodeNo=4", gcnGi) & ""
'         .rsSo001!UpdTime = GetDTString(RightNow)
'         .rsSo001!UpdEn = garyGi(1)
'         .rsSo001.Update
'         .rsSo002.Update
'    End With
'    cn2.CommitTrans
'    cn2.Close
'    Set cn2 = Nothing
'
'    frmSO1100BMDI.lblCustStatusString.Caption = frmSO1100BMDI.rsSo002!CustStatusName & ""
'   'frmSO1100BMDI.RcdToScr
'    ServiceCommon.CommonRcdToScr frmSO1100BMDI.ChildForm
'   'frmSO1100BMDI.RcdToScr
'
'    Unload Me '(3) �̫ᵲ�����e��,�^��Ȥ�򥻸�Ƶe��,�ç�s�e�����e
'  Exit Sub
'CnErr:
'    MsgBox "�s�u���~�ӵL�k���P�I���~��]�G" & Err.Description, vbCritical, "���~"
'  Exit Sub
'ProcErr:
'    MsgBox "��Ƶo�Ϳ��~�ӵL�k���P�I���~��]�G" & Err.Description, vbCritical, "���~"
'    On Error Resume Next
'    cn2.RollbackTrans
'    cn2.Close
'    Set cn2 = Nothing
'End Sub
Private Sub cmdOK_Click()
  On Error GoTo CnErr
    
    Dim cn2 As New ADODB.Connection
    
    If Not MustExist(gdaDelDate, 1, "���P���") Then Exit Sub
    If Not MustExist(gilDelete, 2, "���P��]") Then Exit Sub
    
    With cn2
        .CursorLocation = adUseServer
        .Open gcnGi.ConnectionString
        .BeginTrans
    End With
  
  On Error GoTo ProcErr
    
    ' QQ : 1405 , RA Doc : SO1100B_20070316_Debby_�վ���P�٭�W��.doc
    ' ��q cmdOK_Click Subroutine ���g
    
    ' 6.  ���P�٭�ɡA�ЧP�_�Y�ӫȽs�䥦�A�ȧO�����A�]�����P�A�h�@�ֱN���٭�C
    ' CustStatusCode : 1,���`    2,����  3,���  4,���P  5,�P�P  6,�����
  
    Dim strQry As String
    Dim strInst As String
    Dim strUninst As String
    Dim SvcTp As String
    Dim blnRd As Boolean
    Dim strMdu As String
    Dim intCustStatus As Integer
    Dim strCustStatus As String
    Dim strCust As String
    Dim rs2 As New ADODB.Recordset

    strQry = "SELECT CUSTSTATUSCODE,CUSTSTATUSNAME,OLDSTATUSCODE" & _
                    ",DELCODE,DELNAME,DELDATE,SERVICETYPE" & _
                    " FROM " & GetOwner & "SO002 " & _
                    " WHERE COMPCODE=" & gCompCode & _
                    " AND CUSTID=" & gCustId & _
                    " AND CUSTSTATUSCODE <> 4"

    blnRd = False
    strMdu = RPxx(frmSO1100BMDI.rsSo001!MduId & "")
    strCust = RPxx(frmSO1100BMDI.txtCustName & "")
    strCustStatus = GetRsValue("SELECT DESCRIPTION FROM " & GetOwner & "CD035 WHERE CODENO = 4") & ""
    
    If GetRS(rs2, strQry, cn2, adUseClient, adOpenKeyset, adLockOptimistic) Then
        
        With rs2
            
            While Not .EOF
            
                SvcTp = .Fields("ServiceType") & ""
                intCustStatus = Val(.Fields("CustStatusCode") & "")
                
                '(1) ��s�U�ؤ�Ʋέp���:
                If Not blnRd Then
                    If intCustStatus = 1 Then '.�Y�ӫȤ᪬�A�����` (1), �h:
                        '1.  �N��D�N�X��(CD017)��, �ӫȤ�˾��a�}�ҹ�������D, �N��D�w�ˤ�ƴ�1
                        If ExecuteSQL("UPDATE " & GetOwner & "CD017 SET INSTCNT = INSTCNT - 1" & _
                                                " WHERE CODENO = " & frmSO1100BMDI.rsSo001!InstAddrNO, cn2) = giOK Then
                            blnRd = True
                        Else
                            cn2.RollbackTrans
                            cn2.Close
                            Set cn2 = Nothing
                            Exit Sub
                        End If
                    End If
                End If
                
                If Len(strMdu) > 0 Then '2.  �Y�ӫȤᬰ�j�Ӥ�(�ӵ��Ȥ��ƪ�MduId��즳��)
                    strInst = Switch(SvcTp = "C", "InstCnt", SvcTp = "I", "InstCnt1", SvcTp = "P", "InstCnt3", SvcTp = "D", "InstCnt2")
                    strUninst = Switch(SvcTp = "C", "UnInstCnt", SvcTp = "I", "UnInstCnt1", SvcTp = "P", "UnInstCnt3", SvcTp = "D", "UnInstCnt2")
                    '.�Y�ӫȤ᪬�A�����` (1), �h:
                    If intCustStatus = 1 Then
                        ' �h�N�j�Ӹ���ɪ��Ӥj�Ӫ����Ĥ�ƴ�1
                        If ExecuteSQL("UPDATE " & GetOwner & "SO017" & _
                                                " SET " & strInst & " = " & strInst & " - 1" & _
                                                " WHERE MDUID = '" & strMdu & "'", cn2) < giOK Then
                             cn2.RollbackTrans
                             cn2.Close
                             Set cn2 = Nothing
                             Exit Sub
                         End If
                    ElseIf intCustStatus = 2 Or intCustStatus = 3 Then ' �Y�ӫȤ᪬�A���w���(3)�Τw����(2)
                        ' �h�N�j�Ӹ���ɪ��Ӥj�Ӫ��j�Ӱ�����ݩ��ƴ�1
                        If ExecuteSQL("UPDATE " & GetOwner & "SO017" & _
                                                " SET " & strUninst & " = " & strUninst & " - 1" & _
                                                " WHERE MDUID = '" & strMdu & "'", cn2) < giOK Then
                            cn2.RollbackTrans
                            cn2.Close
                            Set cn2 = Nothing
                            Exit Sub
                        End If
                    End If
                End If
                
                ' �Ȥ��Ʋ���Log��
                
                cn2.Execute "INSERT INTO " & GetOwner & "SO101" & _
                                    " (CUSTID,CUSTNAME,CUSTSTATUSCODE,CUSTSTATUSNAME" & _
                                    ",DELDATE,DELCODE,DELNAME,CUSTSTATUSCODEB,CUSTSTATUSNAMEB" & _
                                    ",UPDTIME,UPDEN,MODEID,COMPCODE,SERVICETYPE) VALUES " & _
                                    "(" & gCustId & ",'" & strCust & "'," & _
                                    intCustStatus & ",'" & .Fields("CustStatusName") & "'," & _
                                     GetNullString(gdaDelDate.Text, giDateV) & "," & _
                                     gilDelete.GetCodeNo2 & ",'" & gilDelete.GetDescription2 & "'," & _
                                     "4,'" & strCustStatus & "','" & _
                                     GetDTString(RightNow) & "','" & garyGi(1) & "',4," & _
                                     gCompCode & ",'" & SvcTp & "')"
                
                .Fields("DelDate") = SolveDate(gdaDelDate.GetValue) & ""
                .Fields("DelCode") = gilDelete.GetCodeNo2 & ""
                .Fields("DelName") = gilDelete.GetDescription2 & ""
                .Fields("OldStatusCode") = intCustStatus
                .Fields("CustStatusCode") = 4
                .Fields("CustStatusName") = strCustStatus
                .Update
                
                .MoveNext
                
            Wend
        End With
    End If
    
    With frmSO1100BMDI.rsSo001
         .Fields("UpdTime") = GetDTString(RightNow)
         .Fields("UpdEn") = garyGi(1)
         .Update
    End With
    
    cn2.CommitTrans
    cn2.Close
    
    Set cn2 = Nothing
    
    With frmSO1100BMDI
        .ProcessData
        ServiceCommon.CommonRcdToScr frmSO1100BMDI.ChildForm
        .lblCustStatusString = .rsSo002!CustStatusName & ""
    End With
    
    Unload Me '(3) �̫ᵲ�����e��,�^��Ȥ�򥻸�Ƶe��,�ç�s�e�����e
  Exit Sub
CnErr:
    MsgBox "�s�u���~�ӵL�k���P�I���~��]�G" & Err.Description, vbCritical, "���~"
  Exit Sub
ProcErr:
    MsgBox "��Ƶo�Ϳ��~�ӵL�k���P�I���~��]�G" & Err.Description, vbCritical, "���~"
    On Error Resume Next
    cn2.RollbackTrans
    cn2.Close
    Set cn2 = Nothing
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    If KeyCode = vbKeyF2 And cmdOK.Enabled Then
        If Not ChkGiList(KeyCode) Then Exit Sub
        cmdOK_Click
        KeyCode = 0
    End If
    If KeyCode = vbKeyEscape And cmdCancel.Enabled Then
        KeyCode = 0
        cmdCancel_Click
        Exit Sub
    End If
   Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
'   FrmOnTop Me
    If Crm Then
        Me.Move frmSO1100BMDI.Left + ((frmSO1100BMDI.Width - Me.Width) / 2), frmSO1100BMDI.Top + ((frmSO1100BMDI.Height - Me.Height) / 2)
    Else
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End If
    gdaDelDate.SetValue (Format(RightNow, "YYYYMMDD"))     ' �i��J���P���(�w�]������)
    SetLst gilDelete, "CodeNo", "Description", 3, 12, 405, 1620, "CD012", , True
    gilDelete.Filter = " Where (ServiceType ='" & gServiceType & "' OR ServiceType IS Null)"
    With gilDelete
'       .SetParent Me
        .SetCodeNo (GetRsValue("SELECT CODENO FROM " & GetOwner & "CD012", gcnGi))
        .Query_Description
    End With
    strActFrmName = Me.Name
   Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo ChkErr
    frmSO1100BMDI.ProcessData
    On Error Resume Next
    ReleaseCOM Me
  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Unload"
End Sub
