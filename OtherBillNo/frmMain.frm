VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "���~���"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   3660
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3660
   StartUpPosition =   1  '���ݵ�������
   Begin VB.CommandButton cmdVendor 
      Caption         =   "�t�Ӹ��"
      Height          =   855
      Left            =   420
      Picture         =   "frmMain.frx":0442
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   1
      Top             =   1110
      Width           =   2745
   End
   Begin VB.CommandButton cmdBill 
      Caption         =   "�b����"
      Height          =   855
      Left            =   420
      Picture         =   "frmMain.frx":1084
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   0
      Top             =   90
      Width           =   2715
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private gstrStationCount As String
Private strAppIniName As String


Private Sub cmdBill_Click()
    frmBillA.Show 1, Me
End Sub

Private Sub cmdVendor_Click()
  On Error Resume Next
    frmVenderComp.Show 1
End Sub

Private Sub Form_Initialize()
  On Error GoTo ChkErr
    SetLocaleInfo GetSystemDefaultLCID, &H1F, "yyyy/MM/dd"
    gstrCurDir = CurPath
    strAppIniName = App.Path
    If Right(strAppIniName, 1) = "\" Then
        strAppIniName = strAppIniName & "Diff.INI"
    Else
        strAppIniName = strAppIniName & "\Diff.INI"
    End If
    
    If OpenDB Is Nothing Then Set OpenDB = CreateObject("giOpenDBObj.OpenDBObj")
    Call ReadINI ' �� GICMIS2.INI Ū����T
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Initialize")
End Sub


Private Function ReadINI() ' �� GICMIS2.INI Ū���]�w���
  On Error GoTo ChkErr
    Dim SysPath As String
    Dim strTmp As String
    Dim lngLen As Long
    Dim pstrDataSource As String
    Dim pstrUserId As String
    Dim pstrMDBPath As String
    Dim pstrPassword As String
    Dim cpCode As String

    SysPath = GetIniPath2(strAppIniName)
    
    If Dir(SysPath) = "" Then ' �p�G�٬O�䤣�� GICMIS.INI �A�h�����t��
        MsgBox "�t�����ҰѼƳ]�w���~�A�L�k�}��" & SysPath & " !", vbCritical, "���~"
        Unload Me
        End
    End If
    
    If Command <> Empty Then
        cpCode = Command()
        cpCode = Replace(cpCode, "CRM", "", 1, , vbTextCompare)
        If cpCode <> Empty Then
            cpCode = Val(Trim(cpCode))
            If cpCode = 0 Then cpCode = ReadFROMINI(SysPath, "DEFAULT", "DEFAULTDB", False) ' �� GICMIS2.INI Ū�� �]�w��
        Else
            cpCode = ReadFROMINI(SysPath, "DEFAULT", "DEFAULTDB", False) ' �� GICMIS2.INI Ū�� �]�w��
        End If
    Else
        cpCode = ReadFROMINI(SysPath, "DEFAULT", "DEFAULTDB", False) ' �� GICMIS2.INI Ū�� �]�w��
    End If
    
    garyGi(14) = cpCode
    pstrDataSource = ReadFROMINI(SysPath, cpCode, "1", True)
    pstrUserId = ReadFROMINI(SysPath, cpCode, "2", True)
    pstrPassword = ReadFROMINI(SysPath, cpCode, "3", True)
    gstrStationCount = ReadFROMINI(SysPath, cpCode, "4", True) ' ���\�h�֤u�@���n�J
    
    If Len(pstrUserId) = 0 Or Len(pstrPassword) = 0 Or Len(gstrStationCount) = 0 Then
        MsgBox "�t �� �� �� �� �� �] �w �� �~ !!", , "���~�T��..!!"
        Unload Me ' �p�G ��ƨӷ� �� �ϥΪ̥N�X �ťաA
        End
    End If
   
    If Alfa Then
        GetGlobal
    Else
        garyGi(9) = cpCode
        garyGi(3) = "Provider=MSDAORA.1;" & _
                     IIf(pstrDataSource = "", "", "Data Source=" + pstrDataSource + ";") & _
                    "Password=" & pstrPassword + ";" & _
                    "User ID=" & pstrUserId + ";" & _
                    "Persist Security Info=True"
        
    End If
    Exit Function
ChkErr:
    If ErrHandle(, True, , "ReadINI") Then Resume 0
    Unload Me
    End
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo ChkErr
    
    
    If Not Alfa Then
        gcnGi.ConnectionString = garyGi(3)
        gcnGi.CursorLocation = adUseClient
        gcnGi.Open
        
        'garyGi(9) = Command
        gCompCode = garyGi(9)
    Else
        gCompCode = garyGi(9)
        MsgBox "Alfa=Ture �L�k���`������!", vbCritical, "�T��"
    End If
   
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

