VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '���u�T�w��ܤ��
   Caption         =   " CITI Bank Payment Gateway  IN TWCA .."
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11655
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":08CA
   ScaleHeight     =   6975
   ScaleWidth      =   11655
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton cmdCred 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�h��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8310
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2880
      Width           =   915
   End
   Begin VB.CommandButton cmdFunc 
      BackColor       =   &H00E0E0E0&
      Caption         =   "��Ƭy�{��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2370
      Width           =   1365
   End
   Begin VB.CommandButton cmdFlow 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�y�{�ܷN��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1950
      Width           =   1365
   End
   Begin VB.TextBox txtMsg 
      Alignment       =   2  '�m�����
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   2880
      Width           =   7935
   End
   Begin VB.TextBox txtResult 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3585
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  '��̬Ҧ�
      TabIndex        =   3
      Text            =   "frmMain.frx":1C748
      Top             =   3180
      Width           =   7935
   End
   Begin VB.Label lblBatchClose 
      BackStyle       =   0  '�z��
      Height          =   495
      Left            =   9810
      TabIndex        =   9
      Top             =   2910
      Width           =   1185
   End
   Begin VB.Label lblCapRev 
      BackStyle       =   0  '�z��
      Height          =   525
      Left            =   7020
      TabIndex        =   5
      Top             =   1590
      Width           =   1095
   End
   Begin VB.Label lblCap 
      BackStyle       =   0  '�z��
      Height          =   435
      Left            =   7230
      TabIndex        =   2
      Top             =   540
      Width           =   645
   End
   Begin VB.Label lblAuthRev 
      BackStyle       =   0  '�z��
      Height          =   525
      Left            =   1950
      TabIndex        =   1
      Top             =   1590
      Width           =   1095
   End
   Begin VB.Label lblAuth 
      BackStyle       =   0  '�z��
      Height          =   435
      Left            =   2190
      TabIndex        =   0
      Top             =   540
      Width           =   645
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��X��API �ѼƤ��� server �Ƚг]�w�p�U:
'
'���ժ�POS URL:   https://npg.hyweb.com.tw:2011
'
'��B�� POS URL:   https://210.66.126.14:2011

Private csPay As Object

Private strServerURL As String
Private strCurrencyType As String
Private intExponentType As Integer
Private intECItype As Integer
Private strMerID As String
Private strMerchantID As String
Private strOrderNo As String
Private strCardNo As String
Private strExpDate As String
Private lngAmt As Long

Private strXID As String
Private strAuthrrpid As String
Private strAuthcode As String
Private strTermseq As String

Private strBatchID As String
Private strBatchSeq As String

Private Sub Form_Load()

    Set csPay = CreateObject("csPayment.clsTWpay")
    
    strServerURL = "https://npg.hyweb.com.tw:2011"
    
    strCurrencyType = "901"
    intExponentType = 0
    intECItype = -1
    strMerID = "118"
    
    strOrderNo = "87654321"
    strExpDate = "200808"
    lngAmt = 1000
    strCardNo = "4444444444444444"
    'strCardNo = "1222222222222222"
    
    strXID = ""
    strAuthrrpid = ""
    strAuthcode = ""
    strTermseq = ""
    strMerchantID = ""
    
End Sub

Private Sub lblAuth_Click() ' �� �v

    Dim lngAPIreturn As Long
    txtResult = ""
    
    ' Function Approve(strHyPOS, strCurrencyType, intExponentType, intECItype, _
    '                             strMerchantID, strOrderNo, strCardNo, strDate, lngAmt) As Long
        
    'If .Approve("http://localhost:2011", "901", 0, 7, 1, "12345678", "4444444444444444", "200808", 1000) = 0 Then
    
    With csPay
        
        If .Approve(strServerURL, strCurrencyType, intExponentType, intECItype, _
                            strMerID, strOrderNo, strCardNo, strExpDate, lngAmt) = 0 Then
            
            txtResult = txtResult & "���o����ѧO�X GetXID : " & .GetXID & vbCrLf
            txtResult = txtResult & "���o���v������N�X GetAuthrrpid : " & .GetAuthrrpid & vbCrLf
            txtResult = txtResult & "���o������վ\�Ǹ� GetTermseq : " & .GetTermseq & vbCrLf
            txtResult = txtResult & "���o������վ\�s�� GetRetrref : " & .GetRetrref & vbCrLf
            txtResult = txtResult & "�妸�������A���� GetBatchclose : " & .GetBatchclose & vbCrLf
            txtResult = txtResult & "���o���v������v�X GetAuthcode : " & .GetAuthcode & vbCrLf
            txtResult = txtResult & "���o���v���B GetAmount : " & .GetAmount & vbCrLf
            txtResult = txtResult & "���o������ȥN�� GetCurrency : " & .GetCurrency & vbCrLf
            txtResult = txtResult & "���o���ȫ��� GetExponent : " & .GetExponent & vbCrLf
            txtMsg = "�H�Υd���v�@�~ ����"
            
            strXID = .GetXID
            strAuthrrpid = .GetAuthrrpid
            strAuthcode = .GetAuthcode
            strTermseq = .GetTermseq
        
        Else
            
            txtResult = txtResult & "Err Status : " & .GetErrStatus & vbCrLf
            txtResult = txtResult & "Err Code : " & .GetErrCode & vbCrLf
            txtResult = txtResult & "API Return : " & .GetAPIreturn & vbCrLf
            txtResult = txtResult & "API Return Description : " & .GetAPIreturn(True) & vbCrLf
            txtResult = txtResult & "API Return ( Byref ) : " & lngAPIreturn & vbCrLf
            txtMsg = "�H�Υd���v�@�~ ����"
        
        End If
    
    End With
    
' GetErrStatus : ���o��������檬�A�C
' GetErrCode : ���o��������~�N�X

End Sub

Private Sub lblAuthRev_Click() ' �������v

    txtResult = ""

    ' Function AUTH_REV(strHyPOS, strCurrencyType, intExponentType, _
                                     intECItype, strMerchantID, strXID, _
                                     strAuthrrpid, strAuthcode, strTermseq, lngAmt) As Long
    
    With csPay
    
        If .AUTH_REV(strServerURL, strCurrencyType, intExponentType, .GetECI(strCardNo), _
                                strMerID, strXID, strAuthrrpid, strAuthcode, strTermseq, lngAmt) = 0 Then
                                
            txtResult = txtResult & "���o������վ\�s�� GetRetrref : " & .GetRetrref & vbCrLf
            txtResult = txtResult & "�妸�������� GetBatchclose : " & .GetBatchclose & vbCrLf
            txtResult = txtResult & "���o������վ\�Ǹ� GetTermseq : " & .GetTermseq & vbCrLf
            txtMsg = "���� �H�Υd���v�@�~ (�Y�����q��) ����"
        
        Else
        
            txtResult = txtResult & "Err Status : " & .GetErrStatus & vbCrLf
            txtResult = txtResult & "Err Code : " & .GetErrCode & vbCrLf
            txtResult = txtResult & "API Return : " & .GetAPIreturn & vbCrLf
            txtResult = txtResult & "API Return Description : " & .GetAPIreturn(True) & vbCrLf
            txtMsg = "���� �H�Υd���v�@�~ (�Y�����q��) ����"
            
        End If
    
    End With
    
End Sub

Private Sub lblCap_Click() ' �д�

    txtResult = ""
    
    'Function CAP(strHyPOS, strCurrencyType, intExponentType, _
                        intECItype, strMerchantID, strXID, _
                        strAuthrrpid, strAuthcode, strTermseq, _
                        lngAmt, Optional lngOrgAmt = -1) As Long
                        
    With csPay
    
        If .CAP(strServerURL, strCurrencyType, intExponentType, .GetECI(strCardNo), _
                    strMerID, strXID, strAuthrrpid, strAuthcode, strTermseq, lngAmt, lngAmt) = 0 Then
                    
            txtResult = txtResult & "���o������վ\�s�� GetRetrref : " & .GetRetrref & vbCrLf
            txtResult = txtResult & "�妸�������� GetBatchclose : " & .GetBatchclose & vbCrLf
            txtResult = txtResult & "����妸�Ǹ� GetBatchId : " & .GetBatchId & vbCrLf
            txtResult = txtResult & "����妸�Ǹ� GetBatchSeq : " & .GetBatchSeq & vbCrLf
            txtMsg = "�H�Υd�дڧ@�~ ����"
            
            'strXID = .GetXID
            'strAuthrrpid = .GetAuthrrpid
            'strAuthcode = .GetAuthcode
            'strTermseq = .GetTermseq
            strBatchID = .GetBatchId
            strBatchSeq = .GetBatchSeq
            
        Else
        
            txtResult = txtResult & "Err Status : " & .GetErrStatus & vbCrLf
            txtResult = txtResult & "Err Code : " & .GetErrCode & vbCrLf
            txtResult = txtResult & "API Return : " & .GetAPIreturn & vbCrLf
            txtResult = txtResult & "API Return Description : " & .GetAPIreturn(True) & vbCrLf
            txtMsg = "�H�Υd�дڧ@�~ ����"
            
        End If
        
    End With
    
End Sub

Private Sub lblCapRev_Click() ' �����д�
    
    txtResult = ""
    
    'Function CAP_REV(strHyPOS, strCurrencyType, intExponentType, _
                                    intECItype, strMerchantID, strXID, _
                                    strAuthrrpid, strAuthcode, strTermseq, _
                                    lngAmt, strBatchID, strBatchSeq) As Long

    With csPay
    
        If .CAP_REV(strServerURL, strCurrencyType, intExponentType, .GetECI(strCardNo), _
                                strMerID, strXID, strAuthrrpid, strAuthcode, strTermseq, lngAmt, _
                                strBatchID, strBatchSeq) = 0 Then
                                
            txtResult = txtResult & "���o������վ\�s�� GetRetrref : " & .GetRetrref & vbCrLf
            txtResult = txtResult & "�妸�������� GetBatchclose : " & .GetBatchclose & vbCrLf
            txtMsg = "���� �H�Υd�дڧ@�~ ����"
            
        Else
        
            txtResult = txtResult & "Err Status : " & .GetErrStatus & vbCrLf
            txtResult = txtResult & "Err Code : " & .GetErrCode & vbCrLf
            txtResult = txtResult & "API Return : " & .GetAPIreturn & vbCrLf
            txtResult = txtResult & "API Return Description : " & .GetAPIreturn(True) & vbCrLf
            txtMsg = "���� �H�Υd�дڧ@�~ ����"
            
        End If
        
    End With
    
End Sub

Private Sub cmdCred_Click() ' �h��

    txtResult = ""
    
'    Function CRED(strHyPOS, strCurrencyType, intExponentType, _
                            intECItype, strMerchantID, strXID, _
                            strAuthrrpid, strAuthcode, lngAmt, _
                            strBatchID, strBatchSeq, Optional lngOrgAmt = -1) As Long

    With csPay
    
        If .CRED(strServerURL, strCurrencyType, intExponentType, .GetECI(strCardNo), _
                        strMerID, strXID, strAuthrrpid, strAuthcode, _
                        lngAmt, strBatchID, strBatchSeq, lngAmt) = 0 Then
                                
            txtResult = txtResult & "���o������վ\�s�� GetRetrref : " & .GetRetrref & vbCrLf
            txtResult = txtResult & "�妸�������� GetBatchclose : " & .GetBatchclose & vbCrLf
            txtResult = txtResult & "����妸�Ǹ� GetBatchId : " & .GetBatchId & vbCrLf
            txtResult = txtResult & "����妸�Ǹ� GetBatchSeq : " & .GetBatchSeq & vbCrLf
            txtMsg = "�i��h�ڰʧ@ ����"
            
        Else
        
            txtResult = txtResult & "Err Status : " & .GetErrStatus & vbCrLf
            txtResult = txtResult & "Err Code : " & .GetErrCode & vbCrLf
            txtResult = txtResult & "API Return : " & .GetAPIreturn & vbCrLf
            txtResult = txtResult & "API Return Description : " & .GetAPIreturn(True) & vbCrLf
            txtMsg = "�i��h�ڰʧ@ ����"
            
        End If
        
    End With
    
End Sub

Private Sub lblBatchClose_Click() ' ���b

    txtResult = ""
    
    ' Function BATCHCLOSE(strHyPOS, strMerchantID) As Long
    
    With csPay
    
        If .BATCHCLOSE(strServerURL, strMerID) = 0 Then
        
            txtResult = txtResult & "����妸�s�� GetBatchId : " & .GetBatchId & vbCrLf
            txtResult = txtResult & "�妸�����ɶ� GetCloseDateTime : " & .GetCloseDateTime & vbCrLf
            txtMsg = "�i��妸���b�ʧ@ ����"
            
        Else
        
            txtResult = txtResult & "Err Status : " & .GetErrStatus & vbCrLf
            txtResult = txtResult & "Err Code : " & .GetErrCode & vbCrLf
            txtResult = txtResult & "API Return : " & .GetAPIreturn & vbCrLf
            txtResult = txtResult & "API Return Description : " & .GetAPIreturn(True) & vbCrLf
            txtMsg = "�i��妸���b�ʧ@ ����"
            
        End If
        
    End With
    
End Sub

Private Sub cmdFlow_Click()
    Form3.Show 1
End Sub

Private Sub cmdFunc_Click()
    Form2.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set csPay = Nothing
End Sub

Private Sub lblAuth_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblAuth.BorderStyle = 1
End Sub

Private Sub lblAuthRev_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblAuthRev.BorderStyle = 1
End Sub

Private Sub lblCap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCap.BorderStyle = 1
End Sub

Private Sub lblCapRev_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCapRev.BorderStyle = 1
End Sub

Private Sub lblBatchClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblBatchClose.BorderStyle = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ctl As Control
    For Each ctl In Controls
        If TypeOf ctl Is Label Then ctl.BorderStyle = 0
    Next
End Sub

