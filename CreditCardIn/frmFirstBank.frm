VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.1#0"; "GiGridR.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.2#0"; "GiList.ocx"
Begin VB.Form frmFirstBank 
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   Icon            =   "frmFirstBank.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   9345
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      Height          =   465
      Left            =   5775
      TabIndex        =   6
      Top             =   3585
      Width           =   1410
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�T�w"
      Default         =   -1  'True
      Height          =   465
      Left            =   2010
      TabIndex        =   5
      Top             =   3585
      Width           =   1410
   End
   Begin VB.Frame fra2 
      Caption         =   "�п�ܦ��O�H���Φ��O�覡����Ȥ覡"
      Height          =   3045
      Left            =   105
      TabIndex        =   7
      Top             =   240
      Width           =   9120
      Begin VB.OptionButton optByParameter 
         Caption         =   "�̦��O�Ѽƿ�J"
         Height          =   315
         Left            =   4605
         TabIndex        =   3
         Top             =   285
         Width           =   1620
      End
      Begin VB.OptionButton optByScreen 
         Caption         =   "�̵e����J"
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Top             =   285
         Width           =   1620
      End
      Begin prjGiGridR.GiGridR ggrData 
         Height          =   2145
         Left            =   4905
         TabIndex        =   4
         Top             =   750
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   3784
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
      Begin prjGiList.GiList gilCMCode 
         Height          =   315
         Left            =   1530
         TabIndex        =   1
         Top             =   750
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilClctEn 
         Height          =   315
         Left            =   1530
         TabIndex        =   2
         Top             =   1155
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F2Corresponding =   -1  'True
      End
      Begin VB.Label lblCMCode 
         AutoSize        =   -1  'True
         Caption         =   "���O�覡"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   435
         TabIndex        =   9
         Top             =   795
         Width           =   780
      End
      Begin VB.Label lblClctEn 
         AutoSize        =   -1  'True
         Caption         =   "���O�H��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   435
         TabIndex        =   8
         Top             =   1200
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmFirstBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnUpdate As Boolean        '��s�{���O�_���\
Private TmpFilePath As String
Private ErrFilePath As String
Private strFilePath As String       'INI�ɮ׸��|
Private strRealDate As String       '�J�b���
Private strSourcePath As String     '��l�ɮ׸��|
Private strUpdEn As String          '�O�����ʤH��
Private strPrgName As String        '��b�{���W��
Private strCMCode As String         '���O�覡�N�X
Private strCMName As String         '���O�覡�W��
Private strClctEn As String         '���O�H���N�X
Private strClctName As String       '���O�H���W��
Private strPTCode As String         '�I�ں���
Private strPTName As String         '�I�ں����W��
Private strServiceType As String    '�A�����O
Private strCompCode As String       '���q�O
Private strOption As String          '�P�_���eoption�����A
Private rsTmp As New ADODB.Recordset

Private Sub IntoData()
On Error GoTo ChkErr

   With objStorePara
     strFilePath = .FilePath
     strRealDate = .RealDate
     strSourcePath = .SourcePath
     strUpdEn = .UpdEn
     strPrgName = .PrgName
     strPTCode = .PTCode
     strPTName = .PTName
     strServiceType = .ServiceType
     strCompCode = .CompCode
   End With
   
   If InitData = False Then
      File.Close
      ErrFile.Close
      Set Fso = Nothing
      Exit Sub
   End If
  Exit Sub
ChkErr:
  ErrSub Me.Name, "IntoData"
End Sub


Private Function IntoAccount() As Boolean
On Error GoTo ChkErr
 Dim strBillNo As String
 Dim strRealAmt As String
 Dim strData As String
 Dim strYYMM As String
 Dim rsTmp As New ADODB.Recordset
 Dim strsql As String
 
        lngCount = 0
        lngErrCount = 0
        TmpFilePath = strFilePath & "\Tmp" & strPrgName & "In.log"
        ErrFilePath = strFilePath & "\" & strPrgName & "Err.log"
        Call GetFilePath(strFilePath)
        Fso.CopyFile strSourcePath, TmpFilePath
        Set File = Fso.OpenTextFile(TmpFilePath, ForReading)
        Set ErrFile = Fso.CreateTextFile(ErrFilePath, True)
        
        strNowTime = Now
        lngTime = Timer
        gcnGi.BeginTrans
        
        While Not File.AtEndOfLine   'Ū����b���
           strData = File.ReadLine   'Ū���@�C���
           
           '�Y�D���ӦC�����H�U�ʧ@
           'If Mid(strData, 1, 1) <> "2" Then GoTo Nextloop
           
           '���s��s���e15�X(�Ȥ�����b��)(B80~B95)
           strBillNo = Mid(strData, 80, 15)
           '�Y���b�����\�A�h�g�J���~log�ɡA���~���ƥ[��
           If Mid(strData, 96, 1) <> "0" Then
              ErrFile.WriteLine ("��ڽs���G " & strBillNo & " ATM���b�����\")
              lngErrCount = lngErrCount + 1
              GoTo Nextloop
           End If
           '���ӵ��J/���b���B(B49~B63)
           strRealAmt = Mid(strData, 49, 13)
           
           '�p�G�e����ܨ̰Ѽƿ�J,�h�~���o������Φ��O�覡���
           If strOption = "Parameter" Then
              
              '���ӵ�������(B55~B62)�Φ�����W��
              'strClctEn = Mid(strData, 10, 8)
              'strClctName = strClctEn
              '���ӵ�����N��(B23~B26)
              strCMCode = Trim(Mid(strData, 23, 4))
              strsql = "Select A.ModeNo,B.Description From CD031A A,CD031 B Where A.ModeNo=B.CodeNo And A.CodeNo = '" & strCMCode & "' And A.ServiceType = '" & strServiceType & "'"
              
              If Not GetRS(rsTmp, strsql, gcnGi) Then
                 MsgBox "�}��Recordset����!", vbExclamation, "ĵ�i..."
                 Exit Function
              End If
              If rsTmp.RecordCount > 0 Then
                 strCMCode = rsTmp("ModeNo")
                 strCMName = rsTmp("Description")
              End If
           Else
              strClctEn = gilClctEn.GetCodeNo
              strClctName = gilClctEn.GetDescription
              strCMCode = gilCMCode.GetCodeNo
              strCMName = gilCMCode.GetDescription
           End If
           
           '�Y�J�b������ťաA�h��(B12~B18)���J�b���YYYMMDD
           If Len(strRealDate) = 0 Then
              strRealDate = Mid(strData, 12, 7)
              strRealDate = Trim(Str(1911 + Val(Mid(strRealDate, 1, 3))) & Mid(strRealDate, 4))
           End If
           
           '����ChkData�ˬd��ƬO�_���~
           If ChkData(strBillNo, strRealAmt, strClctEn, strClctName, strRealDate, strUpdEn, strCMCode, strCMName, strPTCode, strPTName, "") = False Then
              IntoAccount = False
               GoTo Roolback
           End If
Nextloop:
        Wend
        
        gcnGi.CommitTrans

        IntoAccount = True
        On Error Resume Next
        rsTmp.Close
        Set rsTmp = Nothing
  Exit Function
Roolback:
        gcnGi.RollbackTrans
        frmFirstBank.MousePointer = vbDefault
        MsgBox "��s����!", vbExclamation, "ĵ�i..."
        File.Close
        ErrFile.Close
        Set Fso = Nothing
  Exit Function
ChkErr:
  ErrSub Me.Name, "IntoAccount"
End Function

Private Sub cmdCancel_Click()
On Error Resume Next
   Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ChkErr
       
       If Not IsDataOk Then Exit Sub
       frmFirstBank.MousePointer = vbHourglass
       blnUpdate = IntoAccount
       If blnUpdate = False Then  '��s���ѡA�N�����ɮ�����
          objStorePara.UpdState = False
          Exit Sub
       Else
          objStorePara.UpdState = True
       End If
      
       Call DefineRS
       Call ScrToRcd
    
       Call MsgResult       '��ܵ��G�T��
       '�Y���~���Ƥj��0�h��NotePad������ɮפ��e
       If lngErrCount > 0 Then Call Shell("NotePad " & ErrFilePath, vbNormalFocus)
       frmFirstBank.MousePointer = vbDefault
       File.Close
       ErrFile.Close
       Set Fso = Nothing
  
   Exit Sub
ChkErr:
   ErrSub Me.Name, "cmdOK_Click"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error Resume Next
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyEscape Then Unload Me
        If KeyCode = vbKeyF2 Then If cmdOK.Enabled = True Then cmdOK.Value = True
End Sub

Private Sub Form_Load()
  On Error Resume Next
      
      Me.Caption = objStorePara.BankName & ""
      Call IntoData
      Call subGil
      Call RcdToScr

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
      File.Close
      ErrFile.Close
      Set Fso = Nothing
End Sub

'�̵e����J
Private Sub optByScreen_Click()
On Error GoTo ChkErr
      
      strOption = "Screen"
      gilCMCode.Enabled = True
      gilClctEn.Enabled = True
      ggrData.Enabled = False
   Exit Sub
ChkErr:
   ErrSub Me.Name, "optByScreen_Click"
End Sub

'�̦��O�Ѽƿ�J
Private Sub optByParameter_Click()
On Error GoTo ChkErr

      strOption = "Parameter"

     ggrData.Enabled = True
     gilCMCode.Enabled = False
     gilClctEn.Enabled = False
     Call subGrd
     Call OpenRec
   Exit Sub
ChkErr:
   ErrSub Me.Name, "optByParameter_Click"
End Sub

Private Sub subGrd()
   '�]�wgrid�ݩ�
On Error GoTo ChkErr
   Dim mFlds As New prjGiGridR.GiGridFlds
   With mFlds
         .Add "CodeNo", , , , False, "���O�ӷ��N�X", vbLeftJustify
         .Add "Description", , , , False, "���O�ӷ��W��", vbLeftJustify
         .Add "ModeNo", , , , False, "��J�N�X", vbLeftJustify
         .Add "ModeName", , , , False, "��J�W��", vbLeftJustify
   End With
   ggrData.AllFields = mFlds
   ggrData.SetHead
   Exit Sub
ChkErr:
   Call ErrSub(Me.Name, "SubGrd")
End Sub

Private Sub OpenRec()
On Error GoTo ChkErr
  Dim strsql As String
  Dim rsData As New ADODB.Recordset
  
     If rsData.State = adStateOpen Then rsData.Close
     strsql = "Select A.CodeNo,A.Description,A.ModeNo,B.Description ModeName From CD031A A,CD031 B Where A.ModeNo=B.CodeNo And A.StopFlag=0 And A.CompCode = " & strCompCode & " And (A.ServiceType = '" & strServiceType & "' Or A.ServiceType Is Null) Order By A.CodeNo"
     If GetRS(rsData, strsql) Then
        Set ggrData.Recordset = rsData
        ggrData.Refresh
     Else
        MsgBox "�L�k�}��Recodeset!!", vbExclamation, "ĵ�i..."
     End If
     If rsData.State = adStateOpen Then
        rsData.Close
        Set rsData = Nothing
     End If
   Exit Sub
ChkErr:
   ErrSub Me.Name, "OpenRec"
End Sub

Private Sub subGil()
On Error GoTo ChkErr
    SetLst gilClctEn, "EmpNo", "EmpName", 10, 20, , , "CM003", " Where CompCode = " & strCompCode
    SetLst gilCMCode, "CodeNo", "Description", 3, 12, , , "CD031", " Where ServiceType Is Null Or ServiceType = '" & strServiceType & "'"
 
   Exit Sub
ChkErr:
   ErrSub Me.Name, "subGil"
End Sub

Private Sub RcdToScr()
On Error GoTo ChkErr
    
    If Dir(strFilePath & "\" & strPrgName & "In2.log") = "" Then
       optByScreen.Value = True
       gilClctEn.SetCodeNo ""
       gilClctEn.SetDescription ""
       gilCMCode.SetCodeNo ""
       gilCMCode.SetDescription ""
    Else
       With rsTmp
            If .State = adStateOpen Then .Close
            .Open strFilePath & "\" & strPrgName & "In2.log"
            If .Fields("Option") = "Screen" Then
               optByScreen.Value = True
               gilClctEn.SetCodeNo .Fields("ClctEn") & ""
               gilClctEn.Query_Description
               gilCMCode.SetCodeNo .Fields("CMCode") & ""
               gilCMCode.Query_Description
            Else
               optByParameter.Value = True
            End If
       End With
    End If

   Exit Sub
ChkErr:
   ErrSub Me.Name, "RcdToScr"
End Sub

Private Sub ScrToRcd()
On Error GoTo ChkErr
    With rsTmp
       .AddNew
       .Fields("Option") = strOption
       If strOption = "Screen" Then
          .Fields("ClctEn") = gilClctEn.GetCodeNo
          .Fields("ClctName") = gilClctEn.GetDescription
          .Fields("CMCode") = gilCMCode.GetCodeNo
          .Fields("CMName") = gilCMCode.GetDescription
       Else
          .Fields("ClctEn") = ""
          .Fields("ClctName") = ""
          .Fields("CMCode") = ""
          .Fields("CMName") = ""
       End If
       .Save strFilePath & "\" & strPrgName & "In2.log"
    End With
  Exit Sub
ChkErr:
  ErrSub Me.Name, "ScrToRcd"
End Sub

Private Sub DefineRS()
On Error GoTo ChkErr
  If rsTmp.State = adStateOpen Then rsTmp.Close
    With rsTmp
       .LockType = adLockOptimistic
       .Fields.Append "Option", adBSTR, 12
       .Fields.Append "ClctEn", adBSTR, 3
       .Fields.Append "ClctName", adBSTR, 12
       .Fields.Append "CMCode", adBSTR, 3
       .Fields.Append "CMName", adBSTR, 12
       .Open
    End With
  Exit Sub
ChkErr:
  ErrSub Me.Name, "DefineRS"
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo ChkErr

     IsDataOk = False
     If strOption = "" Then
        optByScreen.SetFocus
        MsgBox "�п�ܦ��O�H���Φ��O�覡����Ȥ覡!", vbExclamation, "ĵ�i..."
        Exit Function
     ElseIf strOption = "Screen" Then
        If Len(gilCMCode.GetCodeNo & "") = 0 Then
           gilCMCode.SetFocus
           MsgBox "�п�ܦ��O�覡!", vbExclamation, "ĵ�i..."
           Exit Function
        ElseIf Len(gilClctEn.GetCodeNo & "") = 0 Then
           gilClctEn.SetFocus
           MsgBox "�п�ܦ��O�H��!", vbExclamation, "ĵ�i..."
           Exit Function
        End If
     End If
     IsDataOk = True
   Exit Function
ChkErr:
   ErrSub Me.Name, "IsDataOk"
End Function
