VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "atmCitibank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private strFilePath As String       'INI�ɮ׸��|
Private strSourcePath As String     '��l�ɮ׸��|
Private strRealDate As String       '�J�b���
Private strPrgName As String        '��b�{���W��
Private strUpdEn As String          '�O�����ʤH��
Private blnUpdate As Boolean        '��s�{���O�_���\
Private strClctEn As String           '���O�H���N�X
Private strClctName As String         '���O�H���W��
Private strCMCode As String           '���O�覡�N�X
Private strCMName As String           '���O�覡�W��
Private strPTCode As String           '�I�ں���
Private strPTName As String           '�I�ں����W��
Private strServiceType As String      '�A�����O

Public Sub Action()
On Error GoTo ChkErr
Dim TmpFilePath As String
Dim ErrFilePath As String

   TmpFilePath = strFilePath & "\Tmp" & strPrgName & "In.log"
   ErrFilePath = strFilePath & "\" & strPrgName & "Err.log"
   Call GetFilePath(strFilePath)
   Fso.CopyFile strSourcePath, TmpFilePath
   Set File = Fso.OpenTextFile(TmpFilePath, ForReading)
   Set ErrFile = Fso.CreateTextFile(ErrFilePath)
   
   If InitData = False Then
      File.Close
      ErrFile.Close
      Set Fso = Nothing
      Exit Sub
   End If
   blnUpdate = IntoAccount
   If blnUpdate = False Then  '��s���ѡA�N�����ɮ�����
      File.Close
      ErrFile.Close
      Set Fso = Nothing
      Exit Sub
   End If

   Call MsgResult       '��ܵ��G�T��
   '�Y���~���Ƥj��0�h��NotePad������ɮפ��e
   If lngErrCount > 0 Then Call Shell("NotePad " & ErrFilePath, vbNormalFocus)
              
   File.Close
   ErrFile.Close
   Set Fso = Nothing
  Exit Sub
ChkErr:
  ErrSub "clsCThrust", "AgencyIn"
End Sub


Private Function IntoAccount() As Boolean
On Error GoTo ChkErr
 Dim strBillNo As String
 Dim strRealAmt As String
 Dim intClctEn As Integer
 Dim strData As String
 Dim strYYMM As String
 Dim rstmp As New ADODB.Recordset
 Dim strSQL As String
 
 
        strNowTime = Now
        lngTime = Timer
        gcnGi.BeginTrans
        While Not File.AtEndOfLine   'Ū����b���
           strData = File.ReadLine   'Ū���@�C���
           '�Y�D���ӦC�����H�U�ʧ@
           If Mid(strData, 1, 1) <> "2" Then GoTo Nextloop
           
           '����ڽs��--��ú�~��(B65~B68)+�b�����O(B69~B69)+�A�����O+�Ȥ�s��(B70~B75)
           strYYMM = Trim(CStr(Val(Mid(strData, 65, 4)) + 191100))
           If strServiceType = "" Then
              strBillNo = strYYMM & Mid(strData, 69, 1) & "I" & GetString(Mid(strData, 70, 6), 7, giRight, True)
           Else
              strBillNo = strYYMM & Mid(strData, 69, 1) & strServiceType & GetString(Mid(strData, 70, 6), 7, giRight, True)
           End If
           
           '�Y���b�����\�A�h�g�J���~log�ɡA���~���ƥ[��
           If Mid(strData, 63, 2) <> "00" And Mid(strData, 63, 2) <> "  " Then
              ErrFile.WriteLine ("��ڽs���G " & strBillNo & " ATM���b�����\")
              lngErrCount = lngErrCount + 1
              GoTo Nextloop
           End If
           '���ӵ��J/���b���B(B41~B52)
           strRealAmt = Mid(strData, 41, 12)
           '���ӵ�������(B55~B62)�Φ�����W��
           If strClctEn = "" Then
                strClctEn = Val(Mid(strData, 55, 8))
                
                Select Case strClctEn
                    Case 1
                      strClctName = "�l���۰���b"
                    Case 2
                      strClctName = "�l���d�i�N��"
                    Case 3
                      strClctName = "7-ELEVEn"
                    Case 4
                      strClctName = "���a"
                    Case 5
                      strClctName = "ATM"
                End Select
           End If
           '���ӵ����O�覡
           If strCMCode = "" Then
              strCMCode = "4"
              strSQL = "Select Description From CD031 Where CodeNo = '4'"
              strCMName = GetRsValue(strSQL, gcnGi) & ""
           End If
           '�Y�J�b������ťաA�h��(B21~B26)���J�b���
           If Len(strRealDate) = 0 Then
              strRealDate = Mid(strData, 21, 6)
              strRealDate = Trim(Str(1911 + Val(Mid(strRealDate, 1, 2))) & Mid(strRealDate, 3, 4))
           End If
           
           '����ChkData�ˬd��ƬO�_���~
           If ChkData(strBillNo, strRealAmt, strClctEn, strClctName, strRealDate, strUpdEn, strCMCode, strCMName, strPTCode, strPTName, True) = False Then
              IntoAccount = False
              Exit Function
           End If
Nextloop:
        Wend
        gcnGi.CommitTrans
        IntoAccount = True
        Set rstmp = Nothing
  Exit Function
ChkErr:
  
  ErrSub "AgencyIn", "IntoAccount"
End Function

Public Property Let FilePath(ByVal vData As String)
    strFilePath = vData
End Property

Public Property Let RealDate(ByVal vData As Variant)
    strRealDate = vData
End Property

Public Property Let SourcePath(ByVal vData As Variant)
    strSourcePath = vData
End Property

Public Property Let UpdEn(ByVal vData As Variant)
    strUpdEn = vData
End Property

Public Property Get UpdState() As Variant
    UpdState = blnUpdate
End Property

Public Property Let PrgName(ByVal vData As Variant)
    strPrgName = vData
End Property

Public Property Set ugcnGi(ByVal vData As ADODB.Connection)
    Set gcnGi = vData
End Property
'���O�覡
Public Property Let uCMCode(ByVal vData As Variant)
    strCMCode = vData
End Property
'���O�覡�W��
Public Property Let uCMName(ByVal vData As Variant)
    strCMName = vData
End Property
'���O�H��
Public Property Let uClctEn(ByVal vData As Variant)
    strClctEn = vData
End Property
'���O�H���W��
Public Property Let uClctName(ByVal vData As Variant)
    strClctName = vData
End Property
'�I�ں���
Public Property Let uPTCode(ByVal vData As Variant)
    strPTCode = vData
End Property
'�I�ں����W��
Public Property Let uPTName(ByVal vData As Variant)
    strPTName = vData
End Property
'�A�Ⱥ���
Public Property Let uServiceType(ByVal vData As Variant)
    strServiceType = vData
End Property