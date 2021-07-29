VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '雙線固定對話方塊
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
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdCred 
      BackColor       =   &H00E0E0E0&
      Caption         =   "退款"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8310
      Style           =   1  '圖片外觀
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2880
      Width           =   915
   End
   Begin VB.CommandButton cmdFunc 
      BackColor       =   &H00E0E0E0&
      Caption         =   "函數流程圖"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      Style           =   1  '圖片外觀
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2370
      Width           =   1365
   End
   Begin VB.CommandButton cmdFlow 
      BackColor       =   &H00E0E0E0&
      Caption         =   "流程示意圖"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      Style           =   1  '圖片外觀
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1950
      Width           =   1365
   End
   Begin VB.TextBox txtMsg 
      Alignment       =   2  '置中對齊
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
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   3
      Text            =   "frmMain.frx":1C748
      Top             =   3180
      Width           =   7935
   End
   Begin VB.Label lblBatchClose 
      BackStyle       =   0  '透明
      Height          =   495
      Left            =   9810
      TabIndex        =   9
      Top             =   2910
      Width           =   1185
   End
   Begin VB.Label lblCapRev 
      BackStyle       =   0  '透明
      Height          =   525
      Left            =   7020
      TabIndex        =   5
      Top             =   1590
      Width           =   1095
   End
   Begin VB.Label lblCap 
      BackStyle       =   0  '透明
      Height          =   435
      Left            =   7230
      TabIndex        =   2
      Top             =   540
      Width           =   645
   End
   Begin VB.Label lblAuthRev 
      BackStyle       =   0  '透明
      Height          =   525
      Left            =   1950
      TabIndex        =   1
      Top             =   1590
      Width           =   1095
   End
   Begin VB.Label lblAuth 
      BackStyle       =   0  '透明
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

'花旗的API 參數中之 server 值請設定如下:
'
'測試的POS URL:   https://npg.hyweb.com.tw:2011
'
'營運的 POS URL:   https://210.66.126.14:2011

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

Private Sub lblAuth_Click() ' 授 權

    Dim lngAPIreturn As Long
    txtResult = ""
    
    ' Function Approve(strHyPOS, strCurrencyType, intExponentType, intECItype, _
    '                             strMerchantID, strOrderNo, strCardNo, strDate, lngAmt) As Long
        
    'If .Approve("http://localhost:2011", "901", 0, 7, 1, "12345678", "4444444444444444", "200808", 1000) = 0 Then
    
    With csPay
        
        If .Approve(strServerURL, strCurrencyType, intExponentType, intECItype, _
                            strMerID, strOrderNo, strCardNo, strExpDate, lngAmt) = 0 Then
            
            txtResult = txtResult & "取得交易識別碼 GetXID : " & .GetXID & vbCrLf
            txtResult = txtResult & "取得授權交易之代碼 GetAuthrrpid : " & .GetAuthrrpid & vbCrLf
            txtResult = txtResult & "取得交易之調閱序號 GetTermseq : " & .GetTermseq & vbCrLf
            txtResult = txtResult & "取得交易之調閱編號 GetRetrref : " & .GetRetrref & vbCrLf
            txtResult = txtResult & "批次關閉狀態提示 GetBatchclose : " & .GetBatchclose & vbCrLf
            txtResult = txtResult & "取得授權交易授權碼 GetAuthcode : " & .GetAuthcode & vbCrLf
            txtResult = txtResult & "取得授權金額 GetAmount : " & .GetAmount & vbCrLf
            txtResult = txtResult & "取得交易幣值代號 GetCurrency : " & .GetCurrency & vbCrLf
            txtResult = txtResult & "取得幣值指數 GetExponent : " & .GetExponent & vbCrLf
            txtMsg = "信用卡授權作業 完成"
            
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
            txtMsg = "信用卡授權作業 失敗"
        
        End If
    
    End With
    
' GetErrStatus : 取得交易的執行狀態。
' GetErrCode : 取得交易的錯誤代碼

End Sub

Private Sub lblAuthRev_Click() ' 取消授權

    txtResult = ""

    ' Function AUTH_REV(strHyPOS, strCurrencyType, intExponentType, _
                                     intECItype, strMerchantID, strXID, _
                                     strAuthrrpid, strAuthcode, strTermseq, lngAmt) As Long
    
    With csPay
    
        If .AUTH_REV(strServerURL, strCurrencyType, intExponentType, .GetECI(strCardNo), _
                                strMerID, strXID, strAuthrrpid, strAuthcode, strTermseq, lngAmt) = 0 Then
                                
            txtResult = txtResult & "取得交易之調閱編號 GetRetrref : " & .GetRetrref & vbCrLf
            txtResult = txtResult & "批次關閉提示 GetBatchclose : " & .GetBatchclose & vbCrLf
            txtResult = txtResult & "取得交易之調閱序號 GetTermseq : " & .GetTermseq & vbCrLf
            txtMsg = "取消 信用卡授權作業 (即取消訂單) 完成"
        
        Else
        
            txtResult = txtResult & "Err Status : " & .GetErrStatus & vbCrLf
            txtResult = txtResult & "Err Code : " & .GetErrCode & vbCrLf
            txtResult = txtResult & "API Return : " & .GetAPIreturn & vbCrLf
            txtResult = txtResult & "API Return Description : " & .GetAPIreturn(True) & vbCrLf
            txtMsg = "取消 信用卡授權作業 (即取消訂單) 失敗"
            
        End If
    
    End With
    
End Sub

Private Sub lblCap_Click() ' 請款

    txtResult = ""
    
    'Function CAP(strHyPOS, strCurrencyType, intExponentType, _
                        intECItype, strMerchantID, strXID, _
                        strAuthrrpid, strAuthcode, strTermseq, _
                        lngAmt, Optional lngOrgAmt = -1) As Long
                        
    With csPay
    
        If .CAP(strServerURL, strCurrencyType, intExponentType, .GetECI(strCardNo), _
                    strMerID, strXID, strAuthrrpid, strAuthcode, strTermseq, lngAmt, lngAmt) = 0 Then
                    
            txtResult = txtResult & "取得交易之調閱編號 GetRetrref : " & .GetRetrref & vbCrLf
            txtResult = txtResult & "批次關閉提示 GetBatchclose : " & .GetBatchclose & vbCrLf
            txtResult = txtResult & "交易批次序號 GetBatchId : " & .GetBatchId & vbCrLf
            txtResult = txtResult & "交易批次序號 GetBatchSeq : " & .GetBatchSeq & vbCrLf
            txtMsg = "信用卡請款作業 完成"
            
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
            txtMsg = "信用卡請款作業 失敗"
            
        End If
        
    End With
    
End Sub

Private Sub lblCapRev_Click() ' 取消請款
    
    txtResult = ""
    
    'Function CAP_REV(strHyPOS, strCurrencyType, intExponentType, _
                                    intECItype, strMerchantID, strXID, _
                                    strAuthrrpid, strAuthcode, strTermseq, _
                                    lngAmt, strBatchID, strBatchSeq) As Long

    With csPay
    
        If .CAP_REV(strServerURL, strCurrencyType, intExponentType, .GetECI(strCardNo), _
                                strMerID, strXID, strAuthrrpid, strAuthcode, strTermseq, lngAmt, _
                                strBatchID, strBatchSeq) = 0 Then
                                
            txtResult = txtResult & "取得交易之調閱編號 GetRetrref : " & .GetRetrref & vbCrLf
            txtResult = txtResult & "批次關閉提示 GetBatchclose : " & .GetBatchclose & vbCrLf
            txtMsg = "取消 信用卡請款作業 完成"
            
        Else
        
            txtResult = txtResult & "Err Status : " & .GetErrStatus & vbCrLf
            txtResult = txtResult & "Err Code : " & .GetErrCode & vbCrLf
            txtResult = txtResult & "API Return : " & .GetAPIreturn & vbCrLf
            txtResult = txtResult & "API Return Description : " & .GetAPIreturn(True) & vbCrLf
            txtMsg = "取消 信用卡請款作業 失敗"
            
        End If
        
    End With
    
End Sub

Private Sub cmdCred_Click() ' 退款

    txtResult = ""
    
'    Function CRED(strHyPOS, strCurrencyType, intExponentType, _
                            intECItype, strMerchantID, strXID, _
                            strAuthrrpid, strAuthcode, lngAmt, _
                            strBatchID, strBatchSeq, Optional lngOrgAmt = -1) As Long

    With csPay
    
        If .CRED(strServerURL, strCurrencyType, intExponentType, .GetECI(strCardNo), _
                        strMerID, strXID, strAuthrrpid, strAuthcode, _
                        lngAmt, strBatchID, strBatchSeq, lngAmt) = 0 Then
                                
            txtResult = txtResult & "取得交易之調閱編號 GetRetrref : " & .GetRetrref & vbCrLf
            txtResult = txtResult & "批次關閉提示 GetBatchclose : " & .GetBatchclose & vbCrLf
            txtResult = txtResult & "交易批次序號 GetBatchId : " & .GetBatchId & vbCrLf
            txtResult = txtResult & "交易批次序號 GetBatchSeq : " & .GetBatchSeq & vbCrLf
            txtMsg = "進行退款動作 完成"
            
        Else
        
            txtResult = txtResult & "Err Status : " & .GetErrStatus & vbCrLf
            txtResult = txtResult & "Err Code : " & .GetErrCode & vbCrLf
            txtResult = txtResult & "API Return : " & .GetAPIreturn & vbCrLf
            txtResult = txtResult & "API Return Description : " & .GetAPIreturn(True) & vbCrLf
            txtMsg = "進行退款動作 失敗"
            
        End If
        
    End With
    
End Sub

Private Sub lblBatchClose_Click() ' 結帳

    txtResult = ""
    
    ' Function BATCHCLOSE(strHyPOS, strMerchantID) As Long
    
    With csPay
    
        If .BATCHCLOSE(strServerURL, strMerID) = 0 Then
        
            txtResult = txtResult & "交易批次編號 GetBatchId : " & .GetBatchId & vbCrLf
            txtResult = txtResult & "批次關閉時間 GetCloseDateTime : " & .GetCloseDateTime & vbCrLf
            txtMsg = "進行批次結帳動作 完成"
            
        Else
        
            txtResult = txtResult & "Err Status : " & .GetErrStatus & vbCrLf
            txtResult = txtResult & "Err Code : " & .GetErrCode & vbCrLf
            txtResult = txtResult & "API Return : " & .GetAPIreturn & vbCrLf
            txtResult = txtResult & "API Return Description : " & .GetAPIreturn(True) & vbCrLf
            txtMsg = "進行批次結帳動作 失敗"
            
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

