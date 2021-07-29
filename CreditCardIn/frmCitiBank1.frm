VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.4#0"; "GiList.ocx"
Begin VB.Form frmCitiBank 
   BorderStyle     =   1  '��u�T�w
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   Icon            =   "frmCitiBank.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6105
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�T�w"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   630
      TabIndex        =   1
      Top             =   1965
      Width           =   1410
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "����(&X)"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3900
      TabIndex        =   0
      Top             =   1950
      Width           =   1410
   End
   Begin prjGiList.GiList gilCMCode 
      Height          =   315
      Left            =   1980
      TabIndex        =   2
      Top             =   630
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
      Left            =   1980
      TabIndex        =   3
      Top             =   1080
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
   Begin VB.Label Label1 
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
      Left            =   885
      TabIndex        =   5
      Top             =   675
      Width           =   780
   End
   Begin VB.Label lblClctEn 
      AutoSize        =   -1  'True
      Caption         =   "�I�ں���"
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
      Left            =   885
      TabIndex        =   4
      Top             =   1125
      Width           =   780
   End
End
Attribute VB_Name = "frmCitiBank"
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
Private strBankName As String       '�Ȧ�W��
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
 Dim rsTmp1 As New ADODB.Recordset
 Dim strsql As String
 Dim strCM As String
 
        lngCount = 0
        lngErrCount = 0
        
        TmpFilePath = strFilePath & "\Tmp" & strPrgName & "In.log"
        ErrFilePath = strFilePath & "\" & strPrgName & "Err.log"
        
        Call GetFilePath(strFilePath)
        
        Fso.CopyFile strSourcePath, TmpFilePath
        Set File = Fso.OpenTextFile(TmpFilePath, ForReading)
        Set ErrFile = Fso.CreateTextFile(ErrFilePath)
        
        strNowTime = Now
        lngTime = Timer
        gcnGi.BeginTrans
        While Not File.AtEndOfLine   'Ū����b���
        strData = File.ReadLine
        
        '' �J�b�渹 ���^�n�渹 (B097~B112)
        
          strBillNo = Trim(Mid(strData, 97, 16))
        
           '  �ˬd���v�X(B46~B53)  �A�Y���b�����\�A�h�g�J���~log�ɡA���~���ƥ[��
           ''  ���v�X�O�Ʀr�~�i�H
           
           If Not IsNumeric(Mid(strData, 46, 8)) Or Len(Trim(Mid(strData, 46, 8))) = 0 Then
              ErrFile.WriteLine ("��ڽs���G " & strBillNo & " �H�Υd���b����")
              lngErrCount = lngErrCount + 1
              GoTo Nextloop
           End If
           
           ''�J�b���B
           
             '''   strRealAmt = CStr(CInt(Mid(strData, 38, 8)))
            If Not IsNumeric(Mid(strData, 38, 8)) Then
              strRealAmt = 0
            Else
             strRealAmt = CStr(CInt(Mid(strData, 38, 8)))
            End If
          
           '�J�b��� (B85 ~ B90)
           
              strRealDate = Mid(strData, 85, 6)
              strRealDate = Trim(CStr(2000 + Val(Mid(strRealDate, 1, 2))) & Mid(strRealDate, 3, 4))
              
          ''   strCM = Trim(Mid(strData, 55, 8))
          
              strClctEn = gilClctEn.GetCodeNo
              strClctName = gilClctEn.GetDescription
              strCMCode = gilCMCode.GetCodeNo
              strCMName = gilCMCode.GetDescription
              
              If ChkData( _
                              strBillNo, strRealAmt, strRealDate, strUpdEn, _
                              strCMCode, strCMName, strClctEn, strClctName) = False Then
                              
                    IntoAccount = False
                    GoTo Roolback
              End If
           
Nextloop:

        Wend
        gcnGi.CommitTrans
        IntoAccount = True
        On Error Resume Next
        rsTmp1.Close
        Set rsTmp1 = Nothing
        
  Exit Function
Roolback:
        gcnGi.RollbackTrans
        frmCitiBank.MousePointer = vbDefault
        MsgBox "��s����!", vbExclamation, "ĵ�i..."
        File.Close
        ErrFile.Close
        Set Fso = Nothing
  Exit Function
ChkErr:
  ErrSub "agenCitiBank", "IntoAccount"
End Function

Private Sub cmdCancel_Click()
On Error Resume Next
   Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ChkErr
       
       If Not IsDataOk Then Exit Sub
       frmCitiBank.MousePointer = vbHourglass
       
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
       frmCitiBank.MousePointer = vbDefault
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

Private Sub OpenRec()
On Error GoTo ChkErr
  Dim strsql As String
  Dim rsData As New ADODB.Recordset
  
     If rsData.State = adStateOpen Then rsData.Close
     strsql = _
          "Select " & _
          "A.CodeNo,A.Description,A.ModeNo,B.Description ModeName " & _
          "From " & _
          "CD031A A,CD031 B " & _
          "Where " & _
          "A.ModeNo=B.CodeNo And A.StopFlag=0 And A.CompCode = " & strCompCode & " And (A.ServiceType = '" & strServiceType & "' Or A.ServiceType Is Null) Order By A.CodeNo"
  
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
    SetLst gilClctEn, "CodeNo", "Description", 10, 20, , , "CD032", " Where CompCode = " & strCompCode
    SetLst gilCMCode, "CodeNo", "Description", 10, 20, , , "CD031", " Where ServiceType Is Null Or ServiceType = '" & strServiceType & "'"
Exit Sub
ChkErr:
   ErrSub Me.Name, "subGil"
End Sub

Private Sub RcdToScr()
On Error GoTo ChkErr
    
    If Dir(strFilePath & "\" & strPrgName & "In2.log") = "" Then
       
       gilClctEn.SetCodeNo ""
       gilClctEn.SetDescription ""
       gilCMCode.SetCodeNo ""
       gilCMCode.SetDescription ""
    Else
       With rsTmp
            If .State = adStateOpen Then .Close
            .Open strFilePath & "\" & strPrgName & "In2.log"
               gilClctEn.SetCodeNo .Fields("ClctEn") & ""
               gilClctEn.Query_Description
               gilCMCode.SetCodeNo .Fields("CMCode") & ""
               gilCMCode.Query_Description
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

        If Len(gilCMCode.GetCodeNo & "") = 0 Then
           gilCMCode.SetFocus
           MsgBox "�п�ܦ��O�覡!", vbExclamation, "ĵ�i..."
           Exit Function
        ElseIf Len(gilClctEn.GetCodeNo & "") = 0 Then
           gilClctEn.SetFocus
           MsgBox "�п�ܦ��O�H��!", vbExclamation, "ĵ�i..."
           Exit Function
        End If
     IsDataOk = True
   Exit Function
ChkErr:
   ErrSub Me.Name, "IsDataOk"
End Function

