VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Begin VB.Form frmUnion 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   Icon            =   "frmUnion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   6450
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   570
      TabIndex        =   4
      Top             =   1515
      Width           =   1410
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4365
      TabIndex        =   5
      Top             =   1515
      Width           =   1410
   End
   Begin prjGiList.GiList gilCMCode 
      Height          =   315
      Left            =   2205
      TabIndex        =   1
      Top             =   2205
      Visible         =   0   'False
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
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
      Left            =   2205
      TabIndex        =   2
      Top             =   675
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilEmpNO 
      Height          =   315
      Left            =   2205
      TabIndex        =   0
      Top             =   270
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F2Corresponding =   -1  'True
   End
   Begin Gi_Date.GiDate gdtRealDate 
      Height          =   330
      Left            =   2205
      TabIndex        =   3
      Top             =   1065
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
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
   Begin VB.Label lblRealDate 
      Caption         =   "入帳日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1140
      TabIndex        =   9
      Top             =   1125
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "收費人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1140
      TabIndex        =   8
      Top             =   360
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收費方式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1140
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label lblClctEn 
      AutoSize        =   -1  'True
      Caption         =   "付款種類"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1140
      TabIndex        =   6
      Top             =   750
      Width           =   780
   End
End
Attribute VB_Name = "frmUnion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnUpdate As Boolean        '更新程式是否成功
Private TmpFilePath As String
Private ErrFilePath As String
Private strFilePath As String       'INI檔案路徑
Private strRealDate As String       '入帳日期
Private strSourcePath As String     '原始檔案路徑
Private strUpdEn As String          '記錄異動人員
Private strPrgName As String        '轉帳程式名稱
Private strCMCode As String         '收費方式代碼
Private strCMName As String         '收費方式名稱
Private strClctEn As String         '收費人員代碼
Private strClctName As String       '收費人員名稱
Private strPTCode As String         '付款種類
Private strPTName As String         '付款種類名稱
Private strServiceType As String    '服務類別
Private strBankName As String       '銀行名稱
Private strBankId As String         '銀行代碼
Private strCompCode As String       '公司別
Private strOption As String          '判斷之前option的狀態
Private strUPDName As String        '4388 增加更新異動人員 By Kin 2009/05/04
Private rsTmp As New ADODB.Recordset
Public IsTcb As Boolean
Public isTcbCard As Boolean

Public TableOwnerName As String
Public frmCreditCardType As mCreditCardType    ''  2003/10/30 這個值代表所要使用的信用卡種類
Public pstrCitemCode As String
Public pstrRealDatefrm As String
Private strErrCode As String
Private strErrDescription As String
Private blnFuBonIntegrate As Boolean

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
            strBankId = .BankId
            strUPDName = .UpdName
             GetOwner = TableOwnerName
   End With
   
   If InitData = False Then
            File.Close
            ErrFile.Close
            Set fso = Nothing
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
    Dim strSQL As String
    Dim strCM As String
    
    Dim strEmpNO  As String
    Dim strEmpName As String
    Dim strUPDTime As String
 ''以下這兩個變數，用以儲存當扣帳失敗時，取得的客戶編號
    Dim lngErrCust   As Long
    Dim strErrCust As String
    strFailUpdSQL = Empty
    strUPDTime = GetDTString(RightNow)
    lngCount = 0
    lngErrCount = 0
    
    TmpFilePath = strFilePath & "\Tmp" & strPrgName & "In.log"
    ErrFilePath = strFilePath & "\" & strPrgName & "Err.log"
    
    Call GetFilePath(strFilePath)
    
    fso.CopyFile strSourcePath, TmpFilePath
    Set File = fso.OpenTextFile(TmpFilePath, ForReading)
    Set ErrFile = fso.CreateTextFile(ErrFilePath)
    
    strNowTime = Now
    lngTime = Timer
        
    If Len(gilEmpNO.GetCodeNo) = 0 Then
        strEmpNO = ""
        strEmpName = ""
    Else
        strEmpNO = gilEmpNO.GetCodeNo
        strEmpName = gilEmpNO.GetDescription
    End If
    
        ''  以下這個判斷式使用於 TCB 他行卡的時候，用以取得最後一行資料行的信用卡處理日
        
    If isTcbCard = False And IsTcb = True Then
        Dim lastFile As TextStream
        Dim lastLine As String
        Dim strNoTcbCardDate As String
                
        Set lastFile = fso.OpenTextFile(TmpFilePath, ForReading)
        While Not lastFile.AtEndOfLine   '讀取轉帳資料
            lastLine = ""
            lastLine = lastFile.ReadLine
            ''  strNoTcbCardDate = Mid(lastLine, 20, 8)
            ''  2003/07/17  改成從 16 碼取得信用卡處理日
            strNoTcbCardDate = Mid(lastLine, 18, 8)
        Wend
         lastFile.Close
    End If
    
    
    gcnGi.BeginTrans
        ''  20040827 加上中國信託的判斷
        '#5494 再增加藍新科技的判斷,主要是不要判斷檔頭的資料 By Kin 2010/02/02
        '#5609 增加聯合信用卡批次 By Kin 2010/06/07
        '#7765 ignore the headline by kin 2018/05/09
    If (frmCreditCardType = lngHSBC) Or (frmCreditCardType = lngChinatrust) Or (frmCreditCardType = lngNewWeb) Or (frmCreditCardType = lngFubon) Then
        strData = File.ReadLine
    End If
    Dim aRealDate As String
    While Not File.AtEndOfLine   '讀取轉帳資料
        strData = File.ReadLine
        aRealDate = ""
        Dim strAuthorizeNo As String  '' 這個變數記錄授權碼
        Dim strNote As String
        strAuthorizeNo = Empty
        strNote = Empty
        '****************************************************************************************
        '#3417 失敗更新UCCode與UCName的SQL基本語法 By Kin 2007/12/13
        '#4388 增加更新異動人員與異動時間 By Kin 2009/05/04
        strFailUpdSQL = Empty
        strFailUpdSQL = "UPDATE " & TableOwnerName & "SO033 Set UCCode=" & strUCCode & _
                    ",UCName='" & strUCName & "',UPDEN='" & strUPDName & "'" & _
                    ",UPDTime='" & strUPDTime & "',Note=Note||"
        '****************************************************************************************
        
        strAuthorizeNo = ""
       ''  2003/10/30   以下新增新視波中國商銀的判斷
        
        Select Case frmCreditCardType
            Case 0
                If IsTcb = False Then     ''  如果是中華商銀的格式，則走另外一個迴路
                                 '' 入帳單號 為回登單號 (B097~B112)
                    strBillNo = Trim(Mid(strData, 97, 16))
                      ''  """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
                      '' 檢查授權碼(B46~B53)  ，若扣帳不成功，則寫入錯誤log檔，錯誤筆數加１
                      '' 授權碼是數字才可以
                       ''---------------------------------------------------------------------------------------------------------
                       ''   2002/11/21   -- 調整這一段 檢查授權碼的規則 --
                       ''  bytes 46~53之授權碼alltrim後,長度 >= 4碼(不論為數字或字串),
                       ''  且bytes 69~84前8碼='APPROVED',就當作成功,並轉入系統收費資料中.
                      ''---------------------------------------------------------------------------------------------------------
                    Dim blnApproved  As Boolean
                    '#5395 APPROVED原本只判斷大寫,現在大小寫不判斷
                    blnApproved = Len(Trim(Mid(strData, 46, 8))) >= 4 And UCase(Mid(strData, 69, 8)) = "APPROVED"
                       ''If Not (Len(Trim(Mid(strData, 46, 8))) >= 4 And Mid(strData, 46, 8) = "APPROVED") Then
                    If blnApproved = False Then
                                        ''  2003/07/10  新增以下程式碼取得扣帳失敗之客戶編號
                        lngErrCust = GetfFailCustID(strBillNo)
                        If lngErrCust = -1 Then
                            strErrCust = "查不到此單據編號"
                        Else
                            '' strErrCust = CStr(strBillNo)
                            strErrCust = lngErrCust
                        End If
                         
                        ErrFile.WriteLine ("單據編號： " & strBillNo & " 信用卡扣帳失敗" & "；客戶編號：" & strErrCust)
                        
                        '*****************************************************************************************************************************
                        '#3417 失敗時要更新UCCode與UCName By Kin 2007/12/13
                        If blnFailUpd Then
                            strFailUpdSQL = strFailUpdSQL & "'單據編號： " & strBillNo & " 信用卡扣帳失敗" & "；客戶編號：" & strErrCust & "'" & _
                                          " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                           " And CancelFlag <> 1 And UCCode Is Not Null"
                            gcnGi.Execute strFailUpdSQL
                        End If
                        '*****************************************************************************************************************************
                        
                        lngErrCount = lngErrCount + 1
                        GoTo Nextloop
                    End If
                    If Not IsNumeric(Mid(strData, 38, 8)) Then
                        strRealAmt = 0
                    Else
                        strRealAmt = CStr(Val(Mid(strData, 38, 8)))
                    End If
                    strRealDate = Mid(strData, 85, 6)
                    '#5902 如果遇到民國100,則前面要補1 By Kin 2011/01/11
                    If Left(strRealDate, 1) = "0" Then
                        strRealDate = "1" & strRealDate
                    Else
                        strRealDate = "0" & strRealDate
                    End If
                    '#5902 原本只取6碼,所以年取2位,現在強迫最前面補1位,所以要取3碼 By Kin 2011/01/11
                    strRealDate = Trim(CStr(2000 + Val(Mid(strRealDate, 1, 3))) & Mid(strRealDate, 3, 4))
                    strAuthorizeNo = Mid(strData, 46, 8)
                Else    ''  以下這個是中華商銀的格式檔
                    If isTcbCard = True Then
                        If blnMultiAccount = True Then
                            strBillNo = Mid(strData, 34, 11)   '' 是多帳號的就直接取便是了，因為當初是直接存進去的
                        Else
                            strBillNo = GetBillNo_New(Trim(Mid(strData, 34, 11))) '' 需轉成  16 碼的格式
                        End If
                       ''  目前
                        If Trim(MidMbcs(strData, 139, 1)) <> "Y" Then
                           ''  2003/07/10  新增以下程式碼取得扣帳失敗之客戶編號
                            lngErrCust = GetfFailCustID(strBillNo)
                            If lngErrCust = -1 Then
                                strErrCust = "查不到此單據編號"
                            Else
                                strErrCust = lngErrCust
                            End If
                            ErrFile.WriteLine ("單據編號： " & strBillNo & " 信用卡扣帳失敗" & "；客戶編號：" & strErrCust)
                            
                            '*****************************************************************************************************************************
                            '#3417 失敗時要更新UCCode與UCName By Kin 2007/12/13
                            If blnFailUpd Then
                                strFailUpdSQL = strFailUpdSQL & "'單據編號： " & strBillNo & " 信用卡扣帳失敗" & "；客戶編號：" & strErrCust & "'" & _
                                              " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                               " And CancelFlag <> 1 And UCCode Is Not Null"
                                gcnGi.Execute strFailUpdSQL
                            End If
                            '*****************************************************************************************************************************
                            
                            lngErrCount = lngErrCount + 1
                            GoTo Nextloop
                        End If
                        If Not IsNumeric(Mid(strData, 66, 7)) Then
                            strRealAmt = 0
                        Else
                            strRealAmt = CStr(Val(Mid(strData, 66, 7)))
                        End If
                        strRealDate = Mid(strData, 45, 6)
                        strRealDate = Trim(CStr(1911 + Val(Mid(strRealDate, 1, 2))) & Mid(strRealDate, 3, 4))
                        strAuthorizeNo = Mid(strData, 51, 6)
                    Else
                        strBillNo = Trim(Mid(strData, 17, 15))
                        If Trim(Mid(strData, 96, 2)) <> "00" Then
                            If Left(strData, 1) = "T" Then GoTo Nextloop
                            ''  2003/07/10  新增以下程式碼取得扣帳失敗之客戶編號
                            lngErrCust = GetfFailCustID(strBillNo)
                            If lngErrCust = -1 Then
                                strErrCust = "查不到此單據編號"
                            Else
                                strErrCust = lngErrCust
                            End If
                            ErrFile.WriteLine ("單據編號： " & strBillNo & " 信用卡扣帳失敗" & "；客戶編號：" & strErrCust)
                            
                            '*****************************************************************************************************************************
                            '#3417 失敗時要更新UCCode與UCName By Kin 2007/12/13
                            If blnFailUpd Then
                                strFailUpdSQL = strFailUpdSQL & "'單據編號： " & strBillNo & " 信用卡扣帳失敗" & "；客戶編號：" & strErrCust & "'" & _
                                              " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                               " And CancelFlag <> 1 And UCCode Is Not Null"
                                gcnGi.Execute strFailUpdSQL
                            End If
                            '*****************************************************************************************************************************
                            
                            lngErrCount = lngErrCount + 1
                            GoTo Nextloop
                        End If
                         ''  --2003/07/17--   配合出單規格調整，這個必需回推兩個位元，原本為  58, 12
                        If Not IsNumeric(Mid(strData, 58, 10)) Then
                            strRealAmt = 0
                        Else
                            strRealAmt = CStr(Val(Mid(strData, 58, 10)))
                        End If
                        '入帳日期
                        strRealDate = strNoTcbCardDate
                        strAuthorizeNo = Mid(strData, 98, 6)
                    End If
                End If     ''  如果是中華商銀的格式，則走另外一個迴路的  if  判斷結束
                                         
            Case 1    ''  2003/10/09  以下跑中國商銀格式
                  '中國信託格式   2007/2/8  By Kin 2007/2/9 註解
        '' *************************************************************************************************************
                If blnMultiAccount = True Then
                    strBillNo = (Mid(strData, 49, 11))
                Else
                    strBillNo = GetBillNo_New((Mid(strData, 50, 11)))
                End If
           ''判斷44~45的交易碼
                Dim intTranCode As Integer
                Dim blnApprovedChina  As Boolean
                intTranCode = Val(Trim(Mid(strData, 44, 2)))
         '' 20040906 改成 取 75 位位的8  碼
                '#5395 APPROVED原本只判斷大寫,現在大小寫不判斷
                blnApprovedChina = UCase(Mid(strData, 75, 8)) = "APPROVED"
           
                If blnApprovedChina = False Then
                    lngErrCust = GetfFailCustID(strBillNo)
                    If lngErrCust = -1 Then
                        strErrCust = "查不到此單據編號"
                    Else
                        strErrCust = lngErrCust
                    End If
                    ErrFile.WriteLine ("單據編號： " & strBillNo & " 信用卡扣帳失敗" & "；客戶編號：" & strErrCust)
                    '*****************************************************************************************************************************
                    '#3417 失敗時要更新UCCode與UCName By Kin 2007/12/13
                    If blnFailUpd Then
                        strFailUpdSQL = strFailUpdSQL & "'單據編號： " & strBillNo & " 信用卡扣帳失敗" & "；客戶編號：" & strErrCust & "'" & _
                                      " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                       " And CancelFlag <> 1 And UCCode Is Not Null"
                        gcnGi.Execute strFailUpdSQL
                    End If
                    '*****************************************************************************************************************************
                    
                    lngErrCount = lngErrCount + 1
                    GoTo Nextloop
                End If
'　　　　　**********************************************************************************************************
            '#3049中國信託原本不判斷回授電子檔的金額，改為判斷回授電子檔的金額 By Kin 2007/2/9
             ''2003/10/13  這一段的不用管他了，因為沒有回入帳金額，jacy回覆如是說(此行功能#3049再將它回復)
                If Not IsNumeric(Mid(strData, 30, 10)) Then
                    strRealAmt = 0
                Else
                    strRealAmt = CStr(Val(Mid(strData, 30, 10)))
                End If
   '　　　　************************************************************************************************************
           '入帳日期
                strRealDate = Mid(strData, 91, 6)
                strRealDate = Trim(CStr(2000 + Val(Mid(strRealDate, 1, 2))) & Mid(strRealDate, 3, 4))
                '' 20040906 改成 取 60位位的6碼
                strAuthorizeNo = Mid(strData, 60, 6)
        '' * ************************************************************************************************************
            Case 2     ''  2003/12/02 以下跑新匯豐銀行格式
                
                Dim blnApprovedHSBC  As Boolean
                If blnMultiAccount = True Then   ''  so043.para24=1則取媒體單號11碼
                    strBillNo = (Mid(strData, 53, 11))
                Else                                             ''  so043.para24=0 則取媒體單號15碼
                    strBillNo = GetBillNo_New((Mid(strData, 53, 11)))
                End If
              
                If Trim(Mid(strData, 50, 2)) <> "00" Then
                    blnApprovedHSBC = False
                    ErrFile.WriteLine (GetHSBCResponseMessage(Trim(Mid(strData, 50, 2))))
                    
                    '*****************************************************************************************************************************
                    '#3417 失敗時要更新UCCode與UCName By Kin 2007/12/13
                    If blnFailUpd Then
                        strFailUpdSQL = strFailUpdSQL & "';回應碼:" & Trim(Mid(strData, 50, 2)) & ";訊息:" & _
                                        GetHSBCResponseMessage(Trim(Mid(strData, 50, 2))) & "'" & _
                                      " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                       " And CancelFlag <> 1 And UCCode Is Not Null"
                        gcnGi.Execute strFailUpdSQL
                    End If
                    '*****************************************************************************************************************************
                Else
                    blnApprovedHSBC = True
                End If
    
                If blnApprovedHSBC = False Then
                    lngErrCust = GetfFailCustID(strBillNo)
                    If lngErrCust = -1 Then
                        strErrCust = "查不到此單據編號"
                    Else
                        strErrCust = lngErrCust
                    End If
                    ErrFile.WriteLine ("單據編號： " & strBillNo & " 信用卡扣帳失敗" & "；客戶編號：" & strErrCust)
                    
                    '*****************************************************************************************************************************
                    '#3417 失敗時要更新UCCode與UCName By Kin 2007/12/13
                    If blnFailUpd Then
                        strFailUpdSQL = strFailUpdSQL & "'單據編號： " & strBillNo & " 信用卡扣帳失敗" & "；客戶編號：" & strErrCust & "'" & _
                                      " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                       " And CancelFlag <> 1 And UCCode Is Not Null"
                        gcnGi.Execute strFailUpdSQL
                    End If
                    '*****************************************************************************************************************************
                    
                    lngErrCount = lngErrCount + 1
                    GoTo Nextloop
                End If
                If Not IsNumeric(Mid(strData, 30, 10)) Then
                    strRealAmt = 0
                Else
                    strRealAmt = CStr(Val(Mid(strData, 30, 10)))  ''後兩碼是角分，因此不取只取 10 碼
                End If
             '' 204/02/12 新增一個入帳日期欄位，如果這個欄位有填值，以這個欄位為準
                If Len(gdtRealDate.Text) > 0 Then
                    strRealDate = gdtRealDate.ClipText
                Else
                    strRealDate = Mid(strData, 1, 6)
                    strRealDate = Trim(CStr(1911 + Val(Mid(strRealDate, 1, 2))) & Mid(strRealDate, 3, 4))
                End If
            ''  以下為授權碼 ，必需填入 SO033   AuthorizeNo  欄位 .....
                strAuthorizeNo = Trim(Mid(strData, 43, 6))
                
                
            '#3697 增加一組中國信託新格式 By Kin 2007/12/31
            Case 3
                strErrCode = ""
                If Mid(strData, 1, 1) = "H" Then GoTo Nextloop
                If Mid(strData, 1, 1) = "R" Then GoTo Nextloop
                If Mid(strData, 1, 1) = "" Then GoTo Nextloop
                If Mid(strData, 1, 1) = "T" Then
                    strErrCode = MidMbcs(strData, 43, 2)
                    strBillNo = Trim(Mid(strData, 45, 15))
                    If strErrCode <> "00" Then
                        '#3947 錯誤訊息增加單據編號與客戶編號 By Kin 2008/05/26
                        lngErrCust = GetfFailCustID(strBillNo)
                        If lngErrCust = -1 Then
                            strErrCust = "查不到此單據編號"
                        Else
                            strErrCust = lngErrCust
                        End If
                        ErrFile.WriteLine ("單據編號: " & strBillNo & " 客戶編號: " & strErrCust & " 入帳失敗！ 錯誤代碼：" & strErrCode & " 錯誤原因：" & GetChinatrusNewMessage(strErrCode))
                        '*****************************************************************************************************************************
                        '#3417 失敗時要更新UCCode與UCName By Kin 2007/12/31
                        If blnFailUpd Then
                            strFailUpdSQL = strFailUpdSQL & "'入帳失敗！ 錯誤代碼：" & strErrCode & " 錯誤原因：" & GetChinatrusNewMessage(strErrCode) & "'" & _
                                          " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                           " And CancelFlag <> 1 And UCCode Is Not Null"
                            gcnGi.Execute strFailUpdSQL
                        End If
                        '*****************************************************************************************************************************
    
                        lngErrCount = lngErrCount + 1
                        GoTo Nextloop
                    Else
                        lngErrCust = GetfFailCustID(strBillNo)
                        If lngErrCust = -1 Then
                            strErrCust = "單據：" & strBillNo & "　不存在！請查明"
                            ErrFile.WriteLine (strErrCust)
                            '*****************************************************************************************************************************
                            '#3417 失敗時要更新UCCode與UCName By Kin 2007/12/31
                            If blnFailUpd Then
                                strFailUpdSQL = strFailUpdSQL & "'" & strErrCust & "'" & _
                                              " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                               " And CancelFlag <> 1 And UCCode Is Not Null"
                                gcnGi.Execute strFailUpdSQL
                            End If
                            '*****************************************************************************************************************************
                            lngErrCount = lngErrCount + 1
                            GoTo Nextloop
                        Else
                            If Not IsNumeric(Mid(strData, 30, 13)) Then
                                strRealAmt = 0
                            Else
                                strRealAmt = CStr(Val(Mid(strData, 30, 11)))  ''後兩碼是角分，因此不取只取 10 碼
                            End If
                            If Len(gdtRealDate.Text) > 0 Then
                                strRealDate = gdtRealDate.GetValue
                            Else
                                
                            End If
                            
                            strAuthorizeNo = Trim(MidMbcs(strData, 67, 6))

                        End If
    
                    End If
                End If
            '#5494增加一組藍新科技 By Kin 2010/02/02
            Case 4
                strBillNo = Trim(Replace((Mid(strData, 2, 30)), " ", ""))
                If Mid(strData, 98, 2) <> "00" Then
                    ErrFile.WriteLine ("單據編號: " & strBillNo & _
                        " ;回應碼:" & Mid(strData, 98, 2) & " ;回應訊息:" & GetNewWebMessage(Trim(Mid(strData, 98, 2))))
                    
                    If blnFailUpd Then
                        strFailUpdSQL = strFailUpdSQL & "';回應碼:" & Trim(Mid(strData, 98, 2)) & ";訊息:" & _
                                        GetNewWebMessage(Trim(Mid(strData, 98, 2))) & "'" & _
                                      " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                       " And CancelFlag <> 1 And UCCode Is Not Null"
                        gcnGi.Execute strFailUpdSQL
                    End If
                    lngErrCount = lngErrCount + 1
                    GoTo Nextloop
                Else
                    lngErrCust = GetfFailCustID(strBillNo)
                    If lngErrCust = -1 Then
                        strErrCust = "單據：" & strBillNo & "　不存在！請查明"
                        ErrFile.WriteLine (strErrCust)
                        
                        If blnFailUpd Then
                            strFailUpdSQL = strFailUpdSQL & "'" & strErrCust & "'" & _
                                          " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                           " And CancelFlag <> 1 And UCCode Is Not Null"
                            gcnGi.Execute strFailUpdSQL
                        End If
                        
                        lngErrCount = lngErrCount + 1
                        GoTo Nextloop
                    Else
                        strRealAmt = CStr(Val(GetAmt(Mid(strData, 44, 13))))
                        If Len(gdtRealDate.Text) > 0 Then
                            strRealDate = gdtRealDate.GetValue
                        End If
                        '回覆碼要回至SO033 By Kin 2010/02/03
                        strAuthorizeNo = Trim(MidMbcs(strData, 82, 6))
                    
                        If Trim(MidMbcs(strData, 32, 9)) <> "" Then
                            strNote = "藍新訂單編號:" & Trim(MidMbcs(strData, 32, 9))
                        End If
                    End If
                End If
            Case 5
                strBillNo = Trim(Replace((Mid(strData, 97, 16)), " ", ""))
                If Val(Mid(strData, 66, 3) & "") > 9 Then
                    ErrFile.WriteLine ("單據編號: " & strBillNo & _
                        " ;回應碼:" & Mid(strData, 66, 3) & " ;回應訊息:" & GetUnionBathMessage(Mid(strData, 66, 3), strData))
                        
                        'GetUnionBathMessage(Trim(Mid(strData, 66, 3))))
                    
                    If blnFailUpd Then
                        strFailUpdSQL = strFailUpdSQL & "';回應碼:" & Trim(Mid(strData, 66, 3)) & ";訊息:" & _
                                        GetUnionBathMessage(Mid(strData, 66, 3), strData) & "'" & _
                                      " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                       " And CancelFlag <> 1 And UCCode Is Not Null"
                        gcnGi.Execute strFailUpdSQL
                    End If
                    lngErrCount = lngErrCount + 1
                    GoTo Nextloop
                Else
                    lngErrCust = GetfFailCustID(strBillNo)
                    If lngErrCust = -1 Then
                        strErrCust = "單據：" & strBillNo & "　不存在！請查明"
                        ErrFile.WriteLine (strErrCust)
                        
                        If blnFailUpd Then
                            strFailUpdSQL = strFailUpdSQL & "'" & strErrCust & "'" & _
                                          " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                           " And CancelFlag <> 1 And UCCode Is Not Null"
                            gcnGi.Execute strFailUpdSQL
                        End If
                        lngErrCount = lngErrCount + 1
                        GoTo Nextloop
                    Else
                        strRealAmt = CStr(Val(Mid(strData, 38, 8) & ""))
                        If Len(gdtRealDate.Text) > 0 Then
                            strRealDate = gdtRealDate.GetValue
                        Else
                            
                            aRealDate = Mid(strData, 113, 6)
                            If Left(aRealDate, 1) = "0" Then
                                aRealDate = "1" & aRealDate
                            End If
                            strRealDate = CStr(Val(aRealDate) + 19110000)
                        End If
                        strAuthorizeNo = Replace(Mid(strData, 46, 8), " ", "")
                    End If
                End If
            '#7765 By Kin 2018/05/09
            Case 6
                If Mid(strData, 1, 1) = "T" Then GoTo Nextloop
                If LenB(strData) < 300 Then
                    ErrFile.WriteLine ("回覆檔格式不正確")
                    lngErrCount = lngErrCount + 1
                    GoTo Nextloop
                End If
                strBillNo = Trim(Replace((Mid(strData, 18, 25)), " ", ""))
                '7916 using table to write the description of the response by kin 2018/10/05
                Dim StrResponCode As String
                Dim strResponDes As String
                Dim intReplyLen As Byte
                Dim strSuccess1 As String
                Dim strSuccess2 As String
                '#8528 By Kin 2019/12/03
                intReplyLen = 3
                If blnFuBonIntegrate Then intReplyLen = 4
                '8575 stop checking the bank by kin 2020/02/21
                intReplyLen = 3
                strSuccess1 = Right("0000", intReplyLen)
                strSuccess2 = Right("0001", intReplyLen)
                StrResponCode = Replace(Mid(strData, 199, intReplyLen), " ", "")
                strResponDes = GetBankRespons("CREDITCARDFUBON", StrResponCode)
                'CREDITCARDFUBON
                '8575 stop checking the bank by kin 2020/02/21
                If Right("0000" & StrResponCode, 4) = "0006" And blnFuBonIntegrate And 1 = 0 Then
                    
                End If
                If StrResponCode <> strSuccess1 And StrResponCode <> strSuccess2 Then
                '8575 stop checking the bank by kin 2020/02/21
                    If Not blnFuBonIntegrate And Right("0000" & StrResponCode, 4) = "0006" And 1 = 0 Then
                    
                    Else
                        ErrFile.WriteLine ("單據編號: " & strBillNo & _
                        " ;回應碼:" & StrResponCode & ";訊息:" & strResponDes)
                        
                        If blnFailUpd Then
                            strFailUpdSQL = strFailUpdSQL & "'入帳失敗！回應碼:" & StrResponCode & ";訊息:" & _
                                           strResponDes & "'" & _
                                          " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                           " And CancelFlag <> 1 And UCCode Is Not Null"
                            gcnGi.Execute strFailUpdSQL
                        End If
                        lngErrCount = lngErrCount + 1
                        GoTo Nextloop
                    End If
                    
                End If
                
                If Len(Replace(Mid(strData, 216, 8), " ", "")) = 0 Then
                    ErrFile.WriteLine ("單據編號: " & strBillNo & _
                        " ;回應碼:" & StrResponCode & " ;訊息:授權碼無值")
                                                
                    
                    If blnFailUpd Then
                        strFailUpdSQL = strFailUpdSQL & "'入帳失敗！回應碼:" & StrResponCode & ";訊息:" & _
                                        "授權碼無值" & "'" & _
                                      " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                       " And CancelFlag <> 1 And UCCode Is Not Null"
                        gcnGi.Execute strFailUpdSQL
                    End If
                    lngErrCount = lngErrCount + 1
                    GoTo Nextloop
                Else
                    lngErrCust = GetfFailCustID(strBillNo)
                    If lngErrCust = -1 Then
                        strErrCust = "單據：" & strBillNo & "　不存在！請查明"
                        ErrFile.WriteLine (strErrCust)
                        
                        If blnFailUpd Then
                            strFailUpdSQL = strFailUpdSQL & "'" & strErrCust & "'" & _
                                          " Where " & IIf(blnMultiAccount, "MediaBillNo='", "BillNo='") & strBillNo & "'" & _
                                           " And CancelFlag <> 1 And UCCode Is Not Null"
                            gcnGi.Execute strFailUpdSQL
                        End If
                        lngErrCount = lngErrCount + 1
                        GoTo Nextloop
                    Else
                        strRealAmt = CStr(Val(Mid(strData, 46, 11) & ""))
                        If Len(gdtRealDate.Text) > 0 Then
                            strRealDate = gdtRealDate.GetValue
                        Else
                            aRealDate = Mid(strData, 145, 8)
                            If Len(Replace(aRealDate, " ", "")) = 0 Then
                                aRealDate = Format(Now, "yyyyMMdd")
                            End If
                        End If
                        If Len(aRealDate) > 0 Then
                            strRealDate = aRealDate
                        End If
                        strAuthorizeNo = StrResponCode
                        'strAuthorizeNo = Replace(Mid(strData, 46, 8), " ", "")
                    End If
                End If
        End Select
       
        strClctEn = gilClctEn.GetCodeNo
        strClctName = gilClctEn.GetDescription
        If Len(gilCMCode.GetCodeNo & "") > 0 Then
            strCMCode = gilCMCode.GetCodeNo
            strCMName = gilCMCode.GetDescription
        End If
        '' 20040107 新增若是畫面上收款入帳日期有值，則以其傳進來的為準
        If Len(pstrRealDatefrm) > 0 Then strRealDate = pstrRealDatefrm
        '*****************************************************************************************
        '#3049原本檢核的Function為ChkData，現改為ChkDataTo3313，將Function獨立，以免影嚮其它程式
        '#3417 多傳銀行代碼,如果收費人員有值的話,SO033.NOTE要填入銀行代碼 By Kin 2008/01/22
        '#5275 更新異動人員是Name,不是Code,原本參數為 strUpdEn,現在改為strUPDName By Kin 2009/09/07
        '#5494 藍新科技32~40字元要填入備註欄,所以增加vNote的參數 By Kin 2010/02/03
        If ChkDataTo3313(strBillNo, strRealAmt, strRealDate, strUPDName, _
                        strCMCode, strCMName, strClctEn, strClctName, _
                        strEmpNO, strEmpName, strAuthorizeNo, frmCreditCardType, _
                        pstrCitemCode, strServiceType, strBankId, strNote) = False Then
            IntoAccount = False
            GoTo Roolback
        End If
        '******************************************************************************************
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
    Me.MousePointer = vbDefault
    MsgBox "更新失敗!", vbExclamation, "警告..."
    File.Close
    ErrFile.Close
    Set fso = Nothing
    Exit Function
ChkErr:
    ErrSub "frmUnion", "IntoAccount," & strBillNo & "," & strData
End Function
Private Function GetBankRespons(ByVal PrgName As String, ByVal ErrCode As String) As String
  On Error Resume Next
    Dim ret  As String
    Dim aSQL As String
    Dim aWhere As String
    '8575 stop checking the bank by kin 2020/02/21
    If blnFuBonIntegrate And 1 = 0 Then
        aWhere = " And  nvl(PlatForm,0) = 1 "
    Else
        aWhere = " And  nvl(PlatForm,0) = 0 "
    End If
    aSQL = "Select ErrDescription From " & GetOwner & "SO033G Where upper(PrgName)=upper('" & PrgName & "') " & _
            " And ErrorCode = '" & ErrCode & "' " & aWhere
    ret = GetRsValue(aSQL, gcnGi) & ""
    If Len(ret) = 0 Then
        ret = "銀行回覆失敗(找不到回覆檔)"
    End If
    GetBankRespons = ret
End Function
'#5494 藍新格式取金額,最後2碼為小數點 By Kin 2010/02/03
Private Function GetAmt(ByVal aData As String) As String
  On Error GoTo ChkErr
    Dim retAmt As String
    retAmt = Empty
    retAmt = CStr(Val(Mid(aData, 1, Len(aData) - 2))) & "." & Right(aData, 2)
    GetAmt = retAmt
    Exit Function
ChkErr:
  Call ErrSub("frmUnion", "GetAmt")
End Function
Private Sub cmdCancel_Click()
On Error Resume Next
   Unload Me
   
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    Dim rsCD013 As New ADODB.Recordset
    blnFailUpd = False
    strUCCode = Empty
    strUCName = Empty
    If Not IsDataOk Then Exit Sub
    lngErrCount = 0
    lngAmtTotal = 0
    lngStatusCount = 0
    Me.MousePointer = vbHourglass
    '********************************************************************************************************************************************************
    '#3417 電子檔匯出時,要填入未收原因(RefNo=4) By Kin 2007/12/13
    If Not GetRS(rsCD013, "Select * From " & TableOwnerName & "CD013 Where RefNo=6 And StopFlag<>1 Order By CodeNo Desc", gcnGi, adUseClient, adOpenKeyset) Then Exit Sub
    If rsCD013.EOF Then
        blnFailUpd = False
    Else
        blnFailUpd = True
        strUCCode = rsCD013("CodeNo") & ""
        strUCName = rsCD013("Description") & ""
    End If
    '*********************************************************************************************************************************************************
     '7179
    blnCrossCustCombine = Val(GetRsValue("Select Nvl(CrossCustCombine,0) From " & frmUnion.TableOwnerName & "SO041", gcnGi) & "")
    blnUpdate = IntoAccount '開始入帳
    If blnUpdate = False Then  '更新失敗，將相關檔案關閉
        objStorePara.UpdState = False
        Exit Sub
    Else
       objStorePara.UpdState = True
    End If
    
    Call DefineRS
    Call ScrToRcd
    
    Call MsgResult       '顯示結果訊息
    '若錯誤筆數大於0則用NotePad來顯示檔案內容
    If lngErrCount > 0 Then Call Shell("NotePad " & ErrFilePath, vbNormalFocus)
    If lngStatusCount > 0 Then Call Shell("NotePad " & ErrFilePath, vbNormalFocus)
    Me.MousePointer = vbDefault
    File.Close
    ErrFile.Close
    Set fso = Nothing
    
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
     Dim strCMCodeRefNo5 As String
     Dim strCMCodeRefNo4 As String
     Dim strPTCodeRefNo4 As String
     Me.Caption = objStorePara.BankName & ""
     Call IntoData
     Call subGil
     Call RcdToScr
    '#4022 如果啟動多媒体不判斷服務別 By Kin 2008/07/22
    If strServiceType = Empty Then
        blnMultiAccount = IIf(Val(RPxx(gcnGi.Execute("SELECT para24 FROM " & TableOwnerName & "SO043 WHERE CompCode =" & strCompCode)(0))) = 1, True, False)
    Else
        blnMultiAccount = IIf(Val(RPxx(gcnGi.Execute("SELECT para24 FROM " & TableOwnerName & "SO043 WHERE CompCode =" & strCompCode & " AND ServiceType = '" & strServiceType & "'").GetString)) = 1, True, False)
    End If
    blnFuBonIntegrate = objStorePara.uFubonIntegrate
    
    If blnFuBonIntegrate And frmCreditCardType = lngFubon Then
        '#8555 Cancel to set the value of default and disable  the control by Kin 2020/01/13
        'strCMCodeRefNo5 = gcnGi.Execute("Select CodeNo From " & TableOwnerName & "CD031 Where Nvl(StopFlag,0) = 0 And RefNo= 5 And RowNum = 1 ")(0) & ""
        gilCMCode.Enabled = False
'        If Len(strCMCodeRefNo5) & "" <> "" Then
'            gilCMCode.SetCodeNo strCMCodeRefNo5
'            gilCMCode.Query_Description
'            'gilCMCode.Enabled = False
'        End If
    Else
     '#8555 Setting the value of default when not being fubonintegetrate by kin 2020/01/13
     '#8555 hidde the control
  '      strCMCodeRefNo4 = gcnGi.Execute("Select CodeNo From " & TableOwnerName & "CD031 Where Nvl(StopFlag,0) = 0 And RefNo= 4 And RowNum = 1  order by CodeNo")(0) & ""
  '      If Len(strCMCodeRefNo4 & "") > 0 Then
  '          gilCMCode.SetCodeNo strCMCodeRefNo4
  '          gilCMCode.Query_Description
  '      End If
    End If
    '#8555 Setting the value of default By Kin 2020/01/13
    strPTCodeRefNo4 = gcnGi.Execute("Select CodeNo From " & TableOwnerName & "CD032 Where Nvl(StopFlag,0) = 0 And RefNo= 4 And RowNum = 1 ")(0) & ""
    If Len(strPTCodeRefNo4 & "") > 0 Then
        gilClctEn.SetCodeNo strPTCodeRefNo4
        gilClctEn.Query_Description
    End If
'   gilEmpNO.SetCodeNo objStorePara.UpdEn
'   gilEmpNO.Query_Description
'   If Len(gilEmpNO.GetDescription & "") = 0 Then
'        gilEmpNO.Clear
'   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
      File.Close
      ErrFile.Close
      Set fso = Nothing
End Sub

Private Sub OpenRec()
On Error GoTo ChkErr
  Dim strSQL As String
  Dim rsData As New ADODB.Recordset
  
     If rsData.State = adStateOpen Then rsData.Close
     strSQL = _
          "Select " & _
          "A.CodeNo,A.Description,A.ModeNo,B.Description ModeName " & _
          "From " & TableOwnerName & _
          "CD031A A," & TableOwnerName & "CD031 B " & _
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
    '#4022 如果啟動多媒体不判斷服務別 By Kin 2008/07/22
    If strServiceType <> Empty Then
        SetLst gilClctEn, "CodeNo", "Description", 10, 20, , , "CD032", " Where (ServiceType Is Null Or ServiceType = '" & strServiceType & "'" & ")  AND Nvl(StopFlag,0) = 0   "
        SetLst gilCMCode, "CodeNo", "Description", 10, 20, , , "CD031", " Where (ServiceType Is Null Or ServiceType = '" & strServiceType & "'" & ")  AND Nvl(StopFlag,0) = 0  And RefNo = 4 "
    Else
        SetLst gilClctEn, "CodeNo", "Description", 10, 20, , , "CD032", " Where Nvl(StopFlag,0) = 0   "
        SetLst gilCMCode, "CodeNo", "Description", 10, 20, , , "CD031", " Where Nvl(StopFlag,0) = 0 And RefNo = 4   "
    End If
    SetLst gilEmpNO, "EmpNo", "EmpName", 10, 20, , , "CM003", " Where CompCode = " & strCompCode & "  AND Nvl(StopFlag,0)  = 0  "
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
        If blnFuBonIntegrate And frmCreditCardType = lngFubon Then
        Else
         '   gilCMCode.SetFocus
         '   MsgBox "請選擇收費方式!", vbExclamation, "警告..."
        End If
    End If
    If Len(gilClctEn.GetCodeNo & "") = 0 Then
       gilClctEn.SetFocus
       MsgBox "請選擇付款種類!", vbExclamation, "警告..."
       Exit Function
    End If
    IsDataOk = True
   Exit Function
   
ChkErr:
   ErrSub Me.Name, "IsDataOk"
End Function

Private Function GetfFailCustID(ByVal strID As String) As Long
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim lngCustID As Long
    Dim strBillField As String
  
  On Error GoTo ChkErr
  
  '''以下的 BillNo 欄位 改成 MediaBillNo
  
    If blnMultiAccount = True Then
        strBillField = "MediaBillno"
    Else
        strBillField = "Billno"
    End If

    strSQL = "SELECT CUSTID FROM " & frmUnion.TableOwnerName & "SO033 WHERE " & strBillField & " ='" & strID & "'"
    '7179
    If blnCrossCustCombine Then
        strSQL = "Select A.*,(Select MainCustId From " & frmUnion.TableOwnerName & "SO202 Where MduId=SO001.AMduId) MainCustId " & _
                " From (" & strSQL & ") A," & frmUnion.TableOwnerName & "SO001 Where A.CustId=SO001.CustId "
    End If
    rs.Open strSQL, gcnGi, adOpenKeyset, adLockReadOnly
    
    If rs.EOF Or rs.BOF Then
        lngCustID = -1
    Else
        '7179
        If blnCrossCustCombine Then
            If Len(rs("MainCustId") & "") > 0 Then
                lngCustID = rs("MainCustId")
            Else
                lngCustID = rs("CustID")
            End If
        Else
            lngCustID = rs("CustID")
        End If
    End If
    GetfFailCustID = lngCustID
    rs.Close
    Set rs = Nothing
    Exit Function
ChkErr:
   ErrSub Me.Name, "GetfFailCustID"
End Function

'#3697 中國信託新格式傳回來的錯誤 By Kin 2007/12/31
Private Function GetChinatrusNewMessage(ByVal pstrMessage As String) As String
    Dim strMessage As String
    '#5052 回覆代碼有做調整 By Kin 2009/05/14
    strMessage = ""
    Select Case UCase(pstrMessage)
        Case ""
            strMessage = "無回覆代碼"
            
        Case "12"
            strMessage = "未開卡"
            strErrCode = "12"
        Case "14"
            'strMessage = "交易拒絕-卡號錯"         '#5052
            strMessage = "卡號錯誤"
            strErrCode = "14"
        Case "30"
            strMessage = "交易拒絕-format error"    '#5052
            strErrCode = "30"
        Case "41"
            'strMessage = "沒收卡"                  '#5052
            strMessage = "沒收卡交易拒絕"
            strErrCode = "41"
        Case "43"
            'strMessage = "沒收卡"                  '#5052
            strMessage = "沒收卡交易拒絕"
            strErrCode = "43"
        Case "51"
            strMessage = "額度不足/超額"
            strErrCode = "51"
        Case "54"
            strMessage = "卡片有效期限錯誤"
            strErrCode = "54"
        Case "56"
            'strMessage = "交易拒絕-CAF"
            strMessage = "交易拒絕-卡號異常"
            strErrCode = "56"
        Case "57"
            strMessage = "交易拒絕-未開卡"
            strErrCode = "57"
        Case "61"
            strMessage = "交易拒絕-洽發卡行"
            strErrCode = "61"
        Case "65"
            strMessage = "交易拒絕-洽發卡行"    '#5052
            strErrCode = "65"
        Case "75"
            strMessage = "交易拒絕-密碼錯誤"    '#5052
            strErrCode = "75"
        Case "83"
            strMessage = "交易拒絕- 無此卡號"   '#5052
            strErrCode = "83"
        Case "86"
            strMessage = "交易拒絕-類型有誤"
            strErrCode = "86"
        Case "87"
            'strMessage = "交易拒絕-CVV"
            strMessage = "CVV檢核錯誤"          '#5052
            strErrCode = "87"
        Case "94"
            strMessage = "交易重覆"             '#5052
            strErrCode = "94"
        Case "02"
            strMessage = "照會發卡行"
            strErrCode = "02"
        Case "03"
            strMessage = "商店代號有誤/管制商店"
            strErrCode = "03"
        Case "06"
            strMessage = "系統異常"
            strErrCode = "06"
        Case "07"
            strMessage = "沒收卡"
            strErrCode = "07"
        Case "N8"
            strMessage = "交易拒絕-超過限額"
            strErrCode = "N8"
        Case "O6"
            strMessage = "卡號錯誤-卡號誤"
            strErrCode = "O6"
        Case "N8"
            strMessage = "交易拒絕-超過限額"    '#5052
            strErrCode = "N8"
        Case "P9"
            strMessage = "逾代行日限額/次"      '#5052
            strErrCode = "P9"
        Case "Q0"
            strMessage = "交易日期有誤"         '#5052
            strErrCode = "Q0"
        Case "T1"
            strMessage = "交易拒絕-金額有誤"
            strErrCode = "T1"
        Case "T3"
            strMessage = "商店不接受此卡類信用卡"
            strErrCode = "T3"
        Case "T8"
            strMessage = "卡號錯誤"
            strErrCode = "T8"
        Case Else
            strMessage = "交易拒絕"
            strErrCode = pstrMessage
    End Select
        
    GetChinatrusNewMessage = strMessage

End Function
Private Function GetUnionBathMessage(ByVal pstrMessage As String, ByVal aData As String) As String
  On Error Resume Next
    Dim strMessage As String
    strErrCode = UCase(pstrMessage)
    strMessage = Trim(Mid(aData, 69, 16))
    If strMessage <> "" Then
        GetUnionBathMessage = strMessage
        Exit Function
    End If
    Select Case UCase(pstrMessage)
        Case "086"
            strMessage = "VISA LINE DOWN"
        Case "087"
            strMessage = "HOST DOWN"
        Case "088"
            strMessage = "TIME OUT"
        Case "555"
            strMessage = "SYSTEM TIME_OUT"
        Case "653"
            strMessage = "TC ERROR"
        Case "650"
            strMessage = "REJECT CRD_ERR"
        Case "651"
            strMessage = "SYSTEM ERROR"
        Case "088", "101"
            strMessage = "CALL BANK"
        Case "050", "076"
            strMessage = "REJECT"
        Case "051", "651", "652"
            strMessage = "EXPIRED CARD"
        Case "675"
            strMessage = "SYSTEM REJECT"
        Case Else
            strMessage = "其它訊息"
    End Select
    If Val(UCase(pstrMessage)) >= 900 And Val(UCase(pstrMessage)) < 910 Then
        strMessage = "PICK UP CARD"
    End If
    GetUnionBathMessage = strMessage
End Function
'#5494 藍新科技銀行回覆的代碼 By Kin 2010/02/02
Private Function GetNewWebMessage(ByVal pstrMessage As String) As String
  On Error Resume Next
  Dim strMessage As String
  strMessage = Empty
  strErrCode = UCase(pstrMessage)
  Select Case UCase(pstrMessage)
    Case "00"
        strMessage = "OK"
    Case "01", "02"
        strMessage = "請查詢銀行"
    Case "03"
        strMessage = "無效之特店"
    Case "04", "07"
        strMessage = "沒收卡"
    Case "05"
        strMessage = "拒絕交易"
    Case "06"
        strMessage = "錯誤"
    Case "12"
        strMessage = "無效交易"
    Case "13"
        strMessage = "不符發卡行設定之限額"
    Case "14"
        strMessage = "卡號不存在"
    Case "15"
        strMessage = "發卡行不存在"
    Case "19"
        strMessage = "再輸入之交易型態"
    Case "21"
        strMessage = "無法將原先之交易取消"
    Case "25"
        strMessage = "登錄於VISA 停掛系統之資料無法接受"
    Case "28"
        strMessage = "查詢延n錄於VISA停掛系統之資料無法接受"
    Case "30"
        strMessage = "格式錯誤"
    Case "33", "54"
        strMessage = "過期卡"
    Case "41"
        strMessage = "遺失卡"
    Case "43"
        strMessage = "失竊卡，予以沒收"
    Case "51"
        strMessage = "可用餘額不足"
    Case "55"
        strMessage = "不正確"
    Case "56", "57"
        strMessage = "發卡行未提供持卡人該項交易之功能"
    Case "61"
        strMessage = "超過發卡行設定於VISA系統活動檔之限額"
    Case "62"
        strMessage = "該卡片發卡行已限制"
    Case "65"
        strMessage = "超過發卡行設定於VISA系統活動檔之限次"
    Case "75"
        strMessage = "PIN 輸入錯誤超過次數"
    Case "80"
        strMessage = "交易日期不符"
    Case "81"
        strMessage = "PIN 密碼或CVV 值不正確"
    Case "82"
        strMessage = "CVV 不正確"
    Case "85"
        strMessage = "經查證後,卡片並無錯"
    Case "91"
        strMessage = "發卡行無法接受或處理該訊息"
    Case "93"
        strMessage = "交易無法完成,不允許之交易"
    Case "94"
        strMessage = "重覆交易"
    Case "96"
        strMessage = "系統功能異常"
    Case "99"
        strMessage = "線路繁忙"
    Case "IE"
        strMessage = "ID錯誤"
    Case Else
        strMessage = "不正確的回應碼 !!"
  End Select
  GetNewWebMessage = strMessage
End Function


Private Function GetHSBCResponseMessage(ByVal pstrMessage As String) As String

    Dim strMessage As String

    strMessage = ""
    Select Case pstrMessage
        Case "00"
            strMessage = "APPROVAL xxxxxx 00  APPROVED or COMPLETE SUCCESSFULLY(核准)"
            strErrCode = "00"
        Case "01"
            strMessage = "CALL    01  REFER TO CARD ISSUER(請查詢銀行)"
            strErrCode = "01"
        Case "03"
            strMessage = "INVALID MERCHANT    03  INVALID MERCHANT(商店代號錯誤)"
            strErrCode = "03"
        Case "04"
            strMessage = "'*HOLD CARD  04  PICK UP CARD(沒收卡．請與銀行聯絡)"
            strErrCode = "04"
        Case "05"
            strMessage = "DO NOT HONOR    05  DO NOT HONOR(超額或商店暫時停卡)"
            strErrCode = "05"
        Case "12"
            strMessage = "SERV NOT ALLOWED    12  INVALID TRANSACTION(發卡行斷線或系統設定錯誤．請電話授權)"
            strErrCode = "12"
        Case "13"
            strMessage = "AMOUNT ERROR    13  INVALID AMOUNT(金額錯誤)"
            strErrCode = "13"
        Case "14"
            strMessage = "CARD NO. ERROR  14  INVALID CARD NUMBER(NO SUCH NUMBER) (卡號錯誤)"
            strErrCode = "14"
        Case "19"
            strMessage = "RE ENTER    19  RE-ENTER TRANSACTION(作業逾時)"
            strErrCode = "19"
        Case "41"
            strMessage = "*HOLD-CALL  41  PICK UP CARD(LOSE CARD) (遺失卡)"
            strErrCode = "41"
        Case "43"
            strMessage = "*HOLD CALL  43  PICK UP CARD(STOLEN CARD) (遺失卡或失竊卡)"
            strErrCode = "43"
        Case "51"
            strMessage = "DECLINE 51  NOT SUFFICIFENT FUNDS(超額)"
            strErrCode = "51"
        Case "91"
            strMessage = "NO REPLY    91  SWITCH OR ISSUER SYSTEM INOPERATIVE(未開卡)"
            strErrCode = "91"
        Case "54"
            strMessage = "*EXPIRED CARD   54  EXPIRED CARD(過期卡)"
            strErrCode = "54"
        Case Else
            strMessage = "不正確的回應碼 !!"
        End Select

    GetHSBCResponseMessage = strMessage

End Function

