VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Begin VB.Form frmSO1156A 
   Caption         =   "PPV費用結算作業[SO1156A]"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   Icon            =   "SO1156A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5190
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox txtAuthenticId 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2160
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtSrNO 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2160
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtCustid 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2160
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2. 確定"
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
      Height          =   375
      Left            =   135
      TabIndex        =   9
      Top             =   6435
      Width           =   1365
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
      Height          =   375
      Left            =   1875
      TabIndex        =   10
      Top             =   6435
      Width           =   1365
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3615
      TabIndex        =   11
      Top             =   6435
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   240
      TabIndex        =   19
      Top             =   4680
      Width           =   2745
      Begin VB.CheckBox chkRun 
         Caption         =   "執行結算作業"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         TabIndex        =   8
         Top             =   840
         Width           =   1635
      End
      Begin VB.CheckBox chkTran2 
         Caption         =   "產生未執行結算明細報表"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         TabIndex        =   7
         Top             =   480
         Width           =   2475
      End
      Begin VB.CheckBox chkTran1 
         Caption         =   "產生預計結算明細報表"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         TabIndex        =   6
         Top             =   120
         Value           =   1  '核取
         Width           =   2385
      End
   End
   Begin VB.Frame fraLast 
      Caption         =   "上次結算記錄"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   4950
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "結算截止日:"
         Height          =   180
         Left            =   825
         TabIndex        =   18
         Top             =   300
         Width           =   945
      End
      Begin VB.Label lblTranDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "YY/MM/DD"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1875
         TabIndex        =   17
         Top             =   285
         Width           =   1050
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         Caption         =   "上次執行結算時間:"
         Height          =   180
         Left            =   240
         TabIndex        =   16
         Top             =   630
         Width           =   1485
      End
      Begin VB.Label lblUpdTime 
         BackColor       =   &H00E0E0E0&
         Caption         =   "YY/MM/DD HH:MM:SS"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1875
         TabIndex        =   15
         Top             =   645
         Width           =   1980
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "上次執行結算人員: "
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1530
      End
      Begin VB.Label lblUpdEn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "XXXXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1875
         TabIndex        =   13
         Top             =   975
         Width           =   2760
      End
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   1980
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth1       =   900
      FldWidth2       =   1855
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   300
      Left            =   1920
      TabIndex        =   0
      Top             =   1575
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FldWidth1       =   900
      FldWidth2       =   1855
      F2Corresponding =   -1  'True
   End
   Begin Gi_Date.GiDate gdaEndDate 
      Height          =   270
      Left            =   2160
      TabIndex        =   5
      Top             =   4320
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   476
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ReplaceText     =   -1  'True
   End
   Begin VB.Label lblAuthenticId 
      AutoSize        =   -1  'True
      Caption         =   "認證編號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   30
      Top             =   3600
      Width           =   780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "(含此日之前的資料將被結算)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   3480
      TabIndex        =   29
      Top             =   4320
      Width           =   1665
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblEndDate 
      AutoSize        =   -1  'True
      Caption         =   "結算日期期限"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   840
      TabIndex        =   28
      Top             =   4320
      Width           =   1200
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCustid 
      AutoSize        =   -1  'True
      Caption         =   "客編"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   27
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblPara36 
      BackColor       =   &H00E0E0E0&
      Caption         =   "XXXXXXXXX"
      Height          =   255
      Left            =   2160
      TabIndex        =   26
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblPara35 
      BackColor       =   &H00E0E0E0&
      Caption         =   "XXXXXXXXX"
      Height          =   195
      Left            =   2160
      TabIndex        =   25
      Top             =   2520
      Width           =   1200
   End
   Begin VB.Label lbl5 
      AutoSize        =   -1  'True
      Caption         =   "PPV結算作業月數限額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   2880
      Width           =   1905
   End
   Begin VB.Label lbl4 
      AutoSize        =   -1  'True
      Caption         =   "PPV結算作業金額限額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   2520
      Width           =   1905
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1260
      TabIndex        =   22
      Top             =   1605
      Width           =   585
   End
   Begin VB.Label lblSrNO 
      AutoSize        =   -1  'True
      Caption         =   "STB設備序號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   975
      TabIndex        =   21
      Top             =   3960
      Width           =   1110
   End
   Begin VB.Label lblServicetype 
      AutoSize        =   -1  'True
      Caption         =   "服務類別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   20
      Top             =   2070
      Width           =   780
   End
End
Attribute VB_Name = "frmSO1156A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngTranDate As Long             ' 結算截止日
Dim rsCD039 As New ADODB.Recordset  ' 公司別代碼檔
Dim intPara35 As Integer, intPara36 As Integer
Dim intPara35In As Integer, intPara36In As Integer
Dim rsSO105Amt As New ADODB.Recordset
Dim strSQL As String
Dim strChooseString As String
Dim strFormula As String
Dim strCompCode As String
Dim strServiceType As String
Dim strCustid As String
Dim strSrNO As String
Dim strAuth As String   '認證編號
Dim strEndDate As String
Dim strMsg As String
Dim blnDO As Boolean
Dim blnGive As Boolean

Private Sub cmdCancel_Click()
  On Error Resume Next
    Unload Me
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      Call PreviousRpt(GetPrinterName(5), RptName("SO1156"), Me.Caption)
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    strServiceType = "": strCompCode = "": strEndDate = "": strCustid = "": strSrNO = ""
    blnGive = False
   Call ReleaseCOM(Me)
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    If gdaEndDate.GetValue = "" Then MsgBox "請填入結算日期期限!", vbInformation, "訊息!": Exit Sub
        strServiceType = gilServiceType.GetCodeNo
        strCompCode = gilCompCode.GetCodeNo
        strEndDate = gdaEndDate.GetValue
        strCustid = txtCustID
        strSrNO = txtSrNO
        strAuth = txtAuthenticId
    If chkTran1.Value = 1 Then
        If Not subChoose(1, strCompCode, strServiceType, strEndDate, , strCustid, strSrNO, strAuth) Then Exit Sub
        subPrint
    End If
    If chkRun.Value = 1 Then
        If Not subChoose(3, strCompCode, strServiceType, strEndDate, , strCustid, strSrNO, strAuth) Then Exit Sub
    End If
    If chkTran2.Value = 1 Then
        If lblUpdTime.Caption = "" Then
            MsgBox "無結算截止日!", vbInformation, "訊息": strMsg = "無結算截止日!": If blnGive = True Then Unload Me: Exit Sub Else Exit Sub
        Else
            If Not subChoose(2, strCompCode, strServiceType, strEndDate, lblUpdTime.Caption, strCustid, strSrNO, strAuth) Then Exit Sub
            subPrint
        End If
    End If
     If blnGive = True And chkRun.Value = 1 Then Unload Me
    Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdOk_Click")
End Sub
Private Sub subPrint()
  On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strMsg As String
      Screen.MousePointer = vbHourglass
        cmdCancel.SetFocus
        Me.Enabled = False
          ReadyGoPrint
          subCreateView
          Set rsTmp = gcnGi.Execute("SELECT Count(*) as intCount From " & strViewName & " Where RowNum=1")
          If rsTmp("intCount") = 0 Then
              If InStr(1, strFormula, "1") > 0 Then
                strMsg = "無預計執行結算資料"
              Else
                strMsg = "無未執行結算資料"
              End If
              MsgBox strMsg, vbExclamation, "提示"

'              MsgNoRcd
              SendSQL , , True
          Else
              strSQL = "SELECT * From " & strViewName & " V"
              strFormula = strFormula & ";Sort0={V.CustId}"
             Call PrintRpt(GetPrinterName(5), "SO1156A.rpt", , "預計結算明細報表 [SO1156A]", strSQL, strChooseString, , True, , , strFormula)

          End If
        Me.Enabled = True
      Screen.MousePointer = vbDefault
      CloseRecordset rsTmp
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subPrint")
End Sub

Private Function subCreateView() As Boolean
  Dim strView As String
    On Error Resume Next
      gcnGi.Execute "Drop View " & strViewName
    On Error GoTo ChkErr
        strViewName = GetTmpViewName
        strView = "Create View " & strViewName & " as ( " & strSQL & ")"
        gcnGi.Execute strView
        SendSQL strView, True
    Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subCreateView")
End Function

Private Function NewSO062(Optional ByVal strCompCode As String, _
                           Optional ByVal strServiceType As String, Optional ByVal strEndDate As String, _
                           Optional ByVal strCustid As String, Optional ByVal strSrNO As String, _
                           Optional ByVal strAuth As String, Optional ByRef strMsg As String) As Boolean
On Error GoTo ChkErr
    Dim strCH As String
    
    If rsSO105Amt.RecordCount <> 0 Then
        strCH = "公司別=" & strCompCode & " ;服務類別=" & strServiceType & " ; 結算日期期限=" & strEndDate & "; 客編=" & strCustid & "; STB設備序號=" & strSrNO & "; 認證編號=" & strCustid
        gcnGi.Execute "Update " & GetOwner & "SO062 Set TranDate=TO_DATE('" & strEndDate & "','YYYY/MM/DD') " & _
                      ",UpdEn='" & garyGi(1) & "',UpdTime='" & GetDTString(RightNow) & "',Para='" & strCH & _
                      "' Where Type=3 And CompCode='" & strCompCode & "' AND ServiceType='" & strServiceType & "'"
        strMsg = "結算作業完成！"
        blnDO = True
    Else
        strMsg = "無結算資料！"
        blnDO = True
    End If
    NewSO062 = True
Exit Function
ChkErr:
    Call ErrSub(Me.Name, "NewSO062")
End Function

'抓取欲結算之PPV訂購資料
Private Function subChoose(intType As Integer, Optional ByVal strCompCode As String, _
                           Optional ByVal strServiceType As String, Optional ByVal strEndDate As String, _
                           Optional ByVal strUpdTime As String, Optional ByVal strCustid As String, _
                           Optional ByVal strSrNO As String, Optional ByVal strAuth As String) As Boolean
  On Error GoTo ChkErr
    strChoose = ""
    If intType = 1 Then
        strSQL = "Select A.Custid,A.Authenticid,A.Name,A.OrderNo,A.AcceptTime,B.Amount,B.CreditAmt,C.Playstatus FROM " & _
                 GetOwner & "SO105E A," & GetOwner & "SO105G B," & GetOwner & "SO155 C " & _
                 " Where A.COMPCODE = " & strCompCode & _
                 " AND A.accepttime <= To_Date('" & strEndDate & "','YYYYMMDD')+1 AND B.Canceltime Is Null " & _
                 " AND B.ProductId=C.ProductId AND B.PrePay = 0 AND C.PlayStatus= 1 "

         strFormula = "Type=1"
    ElseIf intType = 2 Then

        strSQL = "Select A.Custid,A.Authenticid,A.Name,A.OrderNo,A.AcceptTime,B.Amount,B.CreditAmt,C.Playstatus FROM " & _
                 GetOwner & "SO105E A," & GetOwner & "SO105G B," & GetOwner & "SO155 C " & _
                 " Where A.COMPCODE = " & strCompCode & _
                 " AND A.accepttime <= To_Date('" & Trim(Format(strUpdTime, "YYMMDD") + 19110000) & "','YYYYMMDD')+1 And B.Canceltime Is Null" & _
                 " AND B.ProductId=C.ProductId"
        strFormula = "Type=2"
    ElseIf intType = 3 Then
        strSQL = "Select A.Custid,A.CompCode,A.AuthenticId,sum(B.Amount) Amount FROM " & _
                 GetOwner & "SO105E A," & GetOwner & "SO105G B," & GetOwner & "SO155 C " & _
                 " Where A.COMPCODE = " & strCompCode & _
                 " AND B.ProductId=C.ProductId AND C.PlayStatus= 1" & _
                 " AND A.accepttime <= To_Date('" & strEndDate & "','YYYYMMDD')+1 AND B.Canceltime Is Null "
    End If
    
    If strCustid <> "" Then Call subAnd("A.Custid ='" & strCustid & "'")
    If strSrNO <> "" Then Call subAnd("B.Facisno ='" & strSrNO & "'")
    If strAuth <> "" Then Call subAnd("A.AuthenticId='" & strAuth & "'")
    
    strSQL = strSQL & " AND A.ServiceType='" & strServiceType & "'" & _
                " AND A.OrderNo=B.OrderNo AND B.Accounting=0" & _
                " AND B.Prepay=0" & IIf(strChoose <> "", " And ", "") & strChoose
                
    If intType = 3 Then
        strSQL = strSQL & " Group by Custid,CompCode,AuthenticId ORDER BY CUSTID,AuthenticId"
        If Not GetRS(rsSO105Amt, strSQL, gcnGi, , adOpenForwardOnly, adLockReadOnly, adUseClient, False) Then Exit Function
        If Not BringAccount(strCompCode, strServiceType, strEndDate) Then Exit Function
        If Not DisCountTest(strCompCode, strServiceType, strEndDate) Then Exit Function
        If Not NewSO062(strCompCode, strServiceType, strEndDate, strCustid, strSrNO, strAuth, strMsg) Then Exit Function
        If strMsg <> "" Then
           MsgBox strMsg, vbInformation, "訊息！"
        End If
        subChoose = True
    End If
    strChooseString = "公司別　: " & subSpace(strCompCode) & ";" & _
                      "服務類別: " & subSpace(strServiceType) & ";" & _
                      "結算日期期限　: " & subSpace(strEndDate) & ";" & _
                      "PPV結算作業金額限額: " & intPara35 & ";" & _
                      "PPV結算作業月數限額: " & intPara36
                      
    subChoose = True
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Function

Private Sub CreateTable()
  On Error Resume Next
    cnn.Execute "Drop Table SO1156"
  On Error GoTo ChkErr
    cnn.Execute "Create Table SO1156 (CompCode text(3),Custid Text(8),AuthenticId text(10),Amount Long,48CodeNo text(10))"
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "CreateTable")
End Sub

Private Function BringAccount(ByVal strCompCode As String, ByVal strServiceType As String, ByVal strEndDate As String) As Boolean
 On Error GoTo ChkErr
    Dim rsSO105Date As New ADODB.Recordset
    Dim rsCD048 As New ADODB.Recordset
    Dim strSO105Date As String, strCD048 As String
    Dim strCD04CodeNO As String
    Dim dblMounth As Double
    Set cnn = GetTmpMdbCn
    '建立暫存檔
    CreateTable
    'rsSO105Amt是找出該客戶及認證編號的總金額
    cnn.BeginTrans
    While Not rsSO105Amt.EOF
        strSO105Date = "select A.Custid,A.CompCode,A.AuthenticId,B.Amount,A.AcceptTime,C.ProductId,C.PlayStatus FROM " & _
                    GetOwner & "SO105E A," & GetOwner & "SO105G B ," & GetOwner & "SO155 C " & _
                    "Where A.COMPCODE = " & strCompCode & _
                    " AND A.accepttime <= To_Date('" & strEndDate & "','YYYYMMDD')+1" & _
                    " AND A.ServiceType='" & strServiceType & "'" & _
                    " AND A.OrderNo=B.OrderNo AND B.Canceltime Is Null AND B.Accounting=0" & _
                    " AND A.Custid='" & rsSO105Amt("Custid") & "'" & _
                    " AND B.ProductId=C.ProductId" & _
                    " AND B.PrePay = 0 AND C.PlayStatus= 1"
        If rsSO105Amt("Authenticid") & "" <> "" Then
            strSO105Date = strSO105Date & " And A.Authenticid='" & rsSO105Amt("Authenticid") & "" & "'"
        End If
        
        If Not GetRS(rsSO105Date, strSO105Date, gcnGi, , adOpenForwardOnly, adLockReadOnly, adUseClient, False) Then Exit Function
        '再由其客編級公司別認證編號去查詢其受理日期
        Do While Not rsSO105Date.EOF

            '其總金額少於Pare35金額限額且受理日期到結算日期其間小於超過Para36的月數限額時不做計算
            dblMounth = CInt((CDate(Format(strEndDate, "####/##/##")) - CDate(Format(rsSO105Date("AcceptTime") & "", "YYYY/MM/DD"))) / 30)
            
            If rsSO105Amt("Amount") < intPara35 Or dblMounth > intPara36 Then
                rsSO105Date.MoveNext: Exit Do
'                If rsSO105Date.EOF Then
            End If
            '查出日期落在CD048的優惠代碼CodeNo
            strCD048 = "select * From " & GetOwner & "CD048 WHERE PeriodType=3 And to_date('" & CDate(Format(rsSO105Date("AcceptTime"), "YYYY/MM/DD")) & "','YYYY/MM/DD') >= StartDate AND to_date('" & CDate(Format(rsSO105Date("AcceptTime"), "YYYY/MM/DD")) & "','YYYY/MM/DD')<= StopDate +1"
            If Not GetRS(rsCD048, strCD048, gcnGi, , adOpenForwardOnly, adLockReadOnly, adUseClient, False) Then Exit Function
            If Not rsCD048.EOF Then
                strCD04CodeNO = rsCD048("CodeNo") & ""
            Else
                strCD04CodeNO = ""
            End If
            '符合資料丟入暫存
           cnn.Execute Replace("INSERT INTO SO1156 (Custid,CompCode,AuthenticId,Amount,48CodeNo) VALUES (" & _
                       GetNullString(rsSO105Date("Custid").Value) & "," & _
                       GetNullString(rsSO105Date("CompCode").Value) & "," & _
                       GetNullString(rsSO105Date("AuthenticId").Value) & "," & _
                       GetNullString(rsSO105Date("Amount").Value) & "," & _
                       GetNullString(strCD04CodeNO) & ")", Chr(0), "")
            
            rsSO105Date.MoveNext
            DoEvents
        Loop
        
        rsSO105Amt.MoveNext
        DoEvents
    Wend
    cnn.CommitTrans
    BringAccount = True
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "BringAccount")
End Function

'折扣率計算
Private Function DisCountTest(ByVal strCompCode As String, ByVal strServiceType As String, ByVal strEndDate As String) As Boolean
 On Error GoTo ChkErr
 Dim rsTmp As New ADODB.Recordset
 Dim rsCD048 As New ADODB.Recordset
 Dim rsSumB As New ADODB.Recordset
 Dim rsCountT As New ADODB.Recordset    '該客戶so105g 在起迄日期內(包含結算)的筆數=片數
 Dim rsCount As New ADODB.Recordset
 Dim rsSO105 As New ADODB.Recordset
 Dim rsSO130 As New ADODB.Recordset
 Dim strCD048 As String, strSumB As String
 Dim dblSum As Double, dblCNT As Double
 Dim strCount As String, strSQL As String
 Dim strCodeNo As String
 Dim StrSO105 As String
 Dim dblDisCAmt As String
 Dim strID As String, strDateNow As String
 Dim dblTotalAmt As Double, strAuthenticID As String
 Set rsTmp = cnn.Execute("SELECT Custid,CompCode,AuthenticId,[48CodeNo],sum(Amount) as AMT From SO1156 Group By Custid,CompCode,AuthenticId,[48CodeNo]")
 
 While Not rsTmp.EOF
    '從優惠代碼找出CD048及48A的資料
    strCD048 = "Select A.*,B.* from " & GetOwner & "CD048 A," & GetOwner & "CD048A B Where A.CodeNo=B.CodeNo And A.CodeNo='" & rsTmp("48CodeNo") & "'"
    If Not GetRS(rsCD048, strCD048, gcnGi, , adOpenForwardOnly, adLockReadOnly, adUseClient, False) Then Exit Function
    
        '找出該客戶so105g 在結算期限內未結算的資料
        '本期片數
        strCount = "Select count(*) Count FROM " & GetOwner & "SO105E A," & GetOwner & "SO105G B ," & GetOwner & "SO155 C" & _
                   " Where A.COMPCODE = '" & strCompCode & "'" & _
                   " AND A.OrderNo=B.OrderNo AND B.CancelTime is null AND B.Accounting=0 AND B.Prepay=0" & _
                   " AND B.ProductId=C.ProductId AND C.PlayStatus= 1" & _
                   " AND A.ServiceType='" & strServiceType & _
                   "' And A.CustId='" & rsTmp("CustId") & _
                   "' AND A.AcceptTime <= To_Date('" & strEndDate & "','YYYYMMDD')+1"
        If rsTmp("Authenticid") & "" <> "" Then
            strCount = strCount & " And A.AuthenticId='" & rsTmp("AuthenticId") & "'"
        End If
    
        If Not GetRS(rsCount, strCount, gcnGi, , adOpenForwardOnly, adLockReadOnly, adUseClient, False) Then Exit Function
        
    If rsTmp("48CodeNo") & "" <> "" Then
        '判斷此筆資料是否累計前期
        If rsCD048("GrandTotalFlag") = 1 Then '累計前期
        '累計前期的總片數
            strSumB = "select sum(B.Amount) as Amount from " & GetOwner & "SO105E A ," & GetOwner & "SO105G B," & GetOwner & "SO155 C Where A.OrderNo=B.OrderNo" & _
                      " AND B.CancelTime is null AND B.Prepay=0" & _
                      " AND B.ProductId=C.ProductId AND C.PlayStatus= 1" & _
                      " And A.CompCode='" & rsTmp("CompCode") & "' and A.CustId='" & rsTmp("CustId") & _
                      "' And A.AcceptTime >= to_date('" & rsCD048("STARTDATE") & "','yyyy/mm/dd')" & _
                      " And A.AcceptTime < to_date('" & rsCD048("STOPDATE") & "','yyyy/mm/dd')+1"
            If rsTmp("Authenticid") & "" <> "" Then
                strSumB = strSumB & " And A.AuthenticId='" & rsTmp("AuthenticId") & "'"
            End If
            If Not GetRS(rsSumB, strSumB, gcnGi, , adOpenForwardOnly, adLockReadOnly, adUseClient, False) Then Exit Function
        '找出該客戶so105g 在起迄日期內(包含結算)的筆數=片數
        '累計前期的總片數
            strCount = "Select count(*) Count FROM " & GetOwner & "SO105E A," & GetOwner & "SO105G B," & GetOwner & "SO155 C" & _
                       " Where A.COMPCODE = '" & strCompCode & "'" & _
                       " AND A.OrderNo=B.OrderNo AND B.CancelTime is null AND B.Prepay=0" & _
                       " AND B.ProductId=C.ProductId AND C.PlayStatus= 1" & _
                       " AND A.ServiceType='" & strServiceType & _
                       "' And A.CustId='" & rsTmp("CustId") & _
                       "' And A.AcceptTime >= to_date('" & rsCD048("STARTDATE") & "','yyyy/mm/dd')" & _
                       " And A.AcceptTime < to_date('" & rsCD048("STOPDATE") & "','yyyy/mm/dd')+1"
            If rsTmp("Authenticid") & "" <> "" Then
                 strCount = strCount & " And A.AuthenticId='" & rsTmp("AuthenticId") & "'"
            End If
            If Not GetRS(rsCountT, strCount, gcnGi, , adOpenForwardOnly, adLockReadOnly, adUseClient, False) Then Exit Function
            
            dblSum = rsSumB("Amount")
            dblCNT = rsCountT("Count")
    
        Else
            dblSum = rsTmp("AMT")        '本期金額
            dblCNT = rsCount("Count")    '本期片數
            
        End If
        '判斷優惠種類 2=折扣率
        If rsCD048("DiscountType") = 2 Then
            '判斷定義種類 1=消費金額
            If rsCD048("DefineType") = 1 Then
                '判斷是否超出消費金額
                If dblSum > rsCD048("TotalAMT") Then
                    If rsCD048("GrandTotalFlag") = 1 Then '累計前期時,超過總金額後以當期金額做折扣
                            dblDisCAmt = CInt(rsTmp("AMT") * (1 - (rsCD048("Rate") / 100)))
                    Else
                    '判斷適用原則總類 1=合計金額
                        If rsCD048("UseType") = 1 Then
                            dblDisCAmt = CInt(rsTmp("AMT") * (1 - (rsCD048("Rate") / 100)))
                        ElseIf rsCD048("UseType") = 2 Then '2=超過金額
                            If (rsTmp("AMT") - rsCD048("TotalAMT")) > 0 Then dblDisCAmt = CInt((rsTmp("AMT") - rsCD048("TotalAMT")) * (1 - (rsCD048("Rate") / 100))) Else dblDisCAmt = 0
                        End If
                    End If
                Else
                    dblDisCAmt = 0
                End If
            '判斷定義種類 2=消費片數
            ElseIf rsCD048("DefineType") = 2 Then
                If dblCNT >= rsCD048("TotalCNT") Then '判斷是否超出消費片數
                    If rsCD048("GrandTotalFlag") = 1 Then '累計前期時,超過片數後以當期金額做折扣
                            dblDisCAmt = CInt(rsTmp("AMT") * (1 - (rsCD048("Rate") / 100)))
                    Else
                        If rsCD048("UseType") = 1 Then
                            dblDisCAmt = CInt(rsTmp("AMT") * (1 - (rsCD048("Rate") / 100)))
                        ElseIf rsCD048("UseType") = 2 Then '2=超過金額
                            If rsCD048("GrandTotalFlag") = 1 Then '累計前期時,超過總金額後以當期金額做折扣
                                dblDisCAmt = CInt(rsTmp("AMT") * (1 - (rsCD048("Rate") / 100)))
                            Else
                                If (rsTmp("AMT") - rsCD048("TotalAMT")) > 0 Then dblDisCAmt = CInt((rsTmp("AMT") - rsCD048("TotalAMT")) * (1 - (rsCD048("Rate") / 100))) Else dblDisCAmt = 0
                            End If
                        End If
                    End If
                Else
                    dblDisCAmt = 0
                End If
            End If
        '判斷優惠種類 1=固定金額
        ElseIf rsCD048("DiscountType") = 1 Then
            If rsCD048("DefineType") = 1 Then
                If dblSum >= rsCD048("totalamt") Then
                    dblDisCAmt = rsCD048("FixAmount")
                Else
                    dblDisCAmt = 0
                End If
            ElseIf rsCD048("DefineType") = 2 Then
                If dblCNT >= rsCD048("TotalCNT") Then
                    dblDisCAmt = rsCD048("FixAmount")
                Else
                    dblDisCAmt = 0
                End If
            End If
        End If
    Else
        dblDisCAmt = 0
    End If
    
    '依條件找出需insert SO130 PPV後付制結算檔的資料
        StrSO105 = "Select A.CompCode,A.Custid,A.Authenticid,A.Name,A.AccountNo,A.CardExpDate,Sum(B.Amount) Amount FROM " & _
                    GetOwner & "SO105E A," & GetOwner & "SO105G B ," & GetOwner & "SO155 C" & _
                    " Where A.COMPCODE = '" & strCompCode & "'" & _
                    " AND A.OrderNo=B.OrderNo AND B.CancelTime is null AND B.Accounting=0 AND B.Prepay=0" & _
                    " AND B.ProductId=C.ProductId AND C.PlayStatus= 1" & _
                    " AND A.ServiceType='" & strServiceType & _
                    "' And A.CustId='" & rsTmp("CustId") & _
                    "' AND A.AcceptTime <= To_Date('" & strEndDate & "','YYYYMMDD')+1"

        If rsTmp("Authenticid") & "" <> "" Then
              StrSO105 = StrSO105 & " And A.AuthenticId='" & rsTmp("AuthenticId") & "'"
        End If
        StrSO105 = StrSO105 & " GROUP BY A.CompCode,A.Custid,A.Authenticid,A.Name,A.AccountNo,A.CardExpDate"
        If Not GetRS(rsSO105, StrSO105, gcnGi, , adOpenForwardOnly, adLockReadOnly, adUseClient, False) Then Exit Function
        
        If Not GetRS(rsSO130, "SELECT * FROM " & GetOwner & "SO130", gcnGi, , adOpenForwardOnly, adLockOptimistic, adUseClient, False) Then Exit Function
'         Set rsSO130 = gcnGi.Execute("SELECT * FROM " & GetOwner & "SO130")
        With rsSO130
            .AddNew
            strID = GetAccId(strCompCode)
            rsSO130("AccountingId").Value = RPxx(strID)
            rsSO130("CompCode").Value = strCompCode
            rsSO130("Custid").Value = rsSO105("Custid")
            rsSO130("AuthenticId").Value = rsSO105("AuthenticId") & ""
            rsSO130("Name").Value = rsSO105("Name")
            rsSO130("AccountNo").Value = rsSO105("AccountNo") & ""
            rsSO130("CardExpDate").Value = rsSO105("CardExpDate")
            rsSO130("AccountingAMT").Value = rsTmp("AMT")
            rsSO130("DiscountAMT").Value = dblDisCAmt
            dblTotalAmt = rsTmp("AMT") - dblDisCAmt
            rsSO130("TotalAMT").Value = dblTotalAmt
            rsSO130("AccountingCNT").Value = rsCount("Count") & ""
            If dblDisCAmt > 0 Then
                rsSO130("Discountcode").Value = rsCD048("CodeNo") & ""
            End If
'            strDateNow = GetDTString(RightNow)
            rsSO130("AccountingTime").Value = RightNow
            rsSO130("AccountingEn").Value = garyGi(0)
            rsSO130("AccountingName").Value = garyGi(1)
           .Update
        End With
        If rsSO105("AUTHENTICID") & "" <> "" Then
            strAuthenticID = rsSO105("Authenticid") & ""
        Else
            strAuthenticID = ""
        End If
        '產生T單
'         Call CreaSO033(rsSO105("Custid"), rsSO105("CitemCode"))
        Call CreaSO033(rsSO105("Custid"), strAuthenticID, dblTotalAmt, strID, strDateNow, strCompCode, strServiceType)
    rsTmp.MoveNext
 Wend
 DisCountTest = True
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "DisCountTest")
End Function

'產生T單
'Private Sub CreaSO033(intCustid As Long, intCitemCode As Integer)
Private Sub CreaSO033(intCustid As Long, strAuthID As String, dblTotalAmt As Double, _
                        strID As String, strDateNow As String, strCompCode As String, strServiceType As String)
  On Error GoTo ChkErr
  Dim rsSO033 As New ADODB.Recordset
  Dim rsCode As New ADODB.Recordset
  Dim rsSO002A As New ADODB.Recordset
  Dim rsSO014 As New ADODB.Recordset
  Dim rsOrdNo As New ADODB.Recordset
  Dim rsTemp As New ADODB.Recordset
  Dim rs105G As New ADODB.Recordset
  Dim strBillNo As String
  Dim strSQL As String, strWhere As String
  Dim strCitemName As String
    If Not GetRS(rsSO033, "SELECT * FROM " & GetOwner & "SO033 Where 1=0", gcnGi, , adOpenForwardOnly, adLockOptimistic, adUseClient, False) Then Exit Sub
    With rsSO033
        .AddNew
        rsSO033("CustID").Value = intCustid
        rsSO033("CompCode").Value = strCompCode
        strBillNo = GetInvoiceNo("T", strServiceType)
        rsSO033("BillNo").Value = strBillNo
        rsSO033("Item").Value = 1
        Set rsTemp = gcnGi.Execute("Select CodeNo,Description FROM " & GetOwner & "CD019 Where RefNo=3")
        rsSO033("CitemCode").Value = rsTemp("CodeNo")
        rsSO033("CitemName").Value = rsTemp("Description")
        rsSO033("AuthenticId").Value = strAuthID
        rsSO033("ShouldDate").Value = GetADdate(RightNow)
        rsSO033("OldAmt").Value = dblTotalAmt
        rsSO033("ShouldAmt").Value = dblTotalAmt
        rsSO033("OldPeriod").Value = 0
        rsSO033("RealPeriod").Value = 0
        strSQL = "SELECT B.CodeNo CD31CODE,B.Description CD31DES,C.CODENO CD13CODE,C.Description CD13DES FROM " & GetOwner & "SO044 A," & GetOwner & "CD031 B," & GetOwner & "CD013 C " & _
                 "WHERE A.CMCode=B.CodeNo AND A.UCCode=C.CODENO AND A.ServiceType='" & strServiceType & "' AND B.StopFlag=0"
        If Not GetRS(rsCode, strSQL, gcnGi, , adOpenForwardOnly, adLockOptimistic, adUseClient, False) Then Exit Sub
        rsSO033("CMCode").Value = rsCode("CD31CODE")
        rsSO033("CMName").Value = rsCode("CD31DES")
        rsSO033("UCCode").Value = rsCode("CD13CODE")
        rsSO033("UCName").Value = rsCode("CD13DES")
        rsSO033("PTCode").Value = 1
        rsSO033("PTName").Value = "現金"
        rsSO033("ClassCode").Value = CInt(gcnGi.Execute("Select ClassCode1 FROM " & GetOwner & "SO001 Where CustId=" & intCustid).GetString)
        rsSO033("FaciSeqNo").Value = intCustid
        rsSO033("ServiceType").Value = strServiceType
        rsSO033("Note").Value = "PPV 費用結算產生的收費資料"
        strSQL = "select * from " & GetOwner & "SO002A WHERE AuthenticId='" & strAuthID & "'"
        If Not GetRS(rsSO002A, strSQL, gcnGi, , adOpenForwardOnly, adLockOptimistic, adUseClient, False) Then Exit Sub
        If Not rsSO002A.EOF Then
            rsSO033("BankCode").Value = rsSO002A("BankCode")
            rsSO033("BankName").Value = rsSO002A("BankName")
            rsSO033("AccountNo").Value = rsSO002A("AccountNo")
        End If
        rsSO033("CreateTime").Value = RightNow
        rsSO033("UpdTime").Value = GetDTString(RightNow)
        rsSO033("CreateEn").Value = garyGi(0)
        rsSO033("UpdEn").Value = garyGi(1)
        '如果地址依劇為收費地址
        If gcnGi.Execute("Select PARA14 FROM " & GetOwner & "SO043 Where COMPCODE='" & strCompCode & "' And ServiceType='" & strServiceType & "'").GetString = 1 Then
            strSQL = "SELECT B.* FROM " & GetOwner & "SO001 A," & GetOwner & "SO014 B WHERE A.ChargeAddrNo=B.AddrNo AND A.CUSTID=" & intCustid
        Else    '裝機地址
            strSQL = "SELECT B.* FROM " & GetOwner & "SO001 A," & GetOwner & "SO014 B WHERE A.InstAddrNo=B.AddrNo AND A.CUSTID=" & intCustid
        End If
        If Not GetRS(rsSO014, strSQL, gcnGi, , adOpenForwardOnly, adLockOptimistic, adUseClient, False) Then Exit Sub
        rsSO033("AddrNo").Value = rsSO014("AddrNo")
        rsSO033("StrtCode").Value = rsSO014("StrtCode")
        rsSO033("MduId").Value = rsSO014("MduId")
        rsSO033("ServCode").Value = rsSO014("ServCode")
        rsSO033("ClctAreaCode").Value = rsSO014("ClctAreaCode")
        rsSO033("OldClctEn").Value = rsSO014("ClctEn")
        rsSO033("OldClctName").Value = rsSO014("ClctName")
        rsSO033("AreaCode").Value = rsSO014("AreaCode")
        rsSO033("ClctEn").Value = rsSO014("ClctEn")
        rsSO033("ClctName").Value = rsSO014("ClctName")
        .Update
        strSQL = "SELECT A.OrderNo FROM " & GetOwner & "SO105E A," & GetOwner & "SO105G B Where A.OrderNo = b.OrderNo AND COMPCODE='" & strCompCode & _
                 "' AND CustId=" & intCustid & " AND B.CancelTime IS NULL"
        If strAuthID <> "" Then
            strSQL = strSQL & " AND A.AuthenticId='" & strAuthID & "'"
        End If
        If Not GetRS(rsOrdNo, strSQL, gcnGi, , adOpenForwardOnly, adLockOptimistic, adUseClient, False) Then Exit Sub
        strWhere = ""
        While Not rsOrdNo.EOF
            strWhere = strWhere & ",'" & rsOrdNo("orderno") & "'"
            rsOrdNo.MoveNext
        Wend
'        strSQL = "Update " & GetOwner & "SO105G SET AccountingTime=TO_DATE('" & Format(RightNow, "YYYY/MM/DD HH:NN:SS") & "','YYYY/MM/DD HH24:MI:SS'), AccountingEn='" & garyGi(0) & "'" & _
                ",AccountingName='" & garyGi(1) & "',AccountingId='" & strID & "',Accounting=1," & _
                "BillNO='" & strBillNo & "',Item=1,CitemCode='" & rsTemp("CodeNo") & "',CitemName='" & rsTemp("Description") & "' Where CancelTime IS NULL AND OrderNo IN (" & Mid(strWhere, 2) & ")"
'        gcnGi.Execute strSQL
        strSQL = "SELECT B.AccountingTime,B.AccountingEn,B.AccountingName,B.AccountingId,B.Accounting,B.BillNO,B.Item,B.CitemCode,B.CitemName" & _
                 " FROM " & GetOwner & "SO105G B, " & GetOwner & "SO155 C" & _
                 " Where CancelTime Is Null " & _
                 " AND B.ProductId=C.ProductId AND C.PlayStatus= 1" & _
                 " AND OrderNo IN (" & Mid(strWhere, 2) & ")"
        If GetRS(rs105G, strSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) = False Then Exit Sub
        While Not rs105G.EOF
            With rs105G
'            .AddNew
            .Fields("AccountingTime") = Format(RightNow, "YYYY/MM/DD HH:NN:SS")
            .Fields("AccountingEn") = garyGi(0)
            .Fields("AccountingName") = garyGi(1)
            .Fields("AccountingId") = strID
            .Fields("Accounting") = 1
            .Fields("BillNO") = strBillNo
            .Fields("Item") = 1
            .Fields("CitemCode") = rsTemp("CodeNo")
            .Fields("CitemName") = rsTemp("Description")
            .Update
            End With
            rs105G.MoveNext
        Wend
    End With
  Exit Sub
ChkErr:
    ErrSub Me.Name, "CreaSO033"
End Sub

Private Function GetAccId(strCompCode As String)
    Dim strAccID As String
    Dim strSQL As String
    strSQL = "Select Max(Accountingid) FROM " & GetOwner & "SO130 WHERE AccountingTime > TO_DATE('" & Format(RightNow, "YYYY/MM/DD") & "','YYYY/MM/DD') AND AccountingTime <= TO_DATE('" & Format(RightNow, "YYYY/MM/DD") & "','YYYY/MM/DD')+1"
    strAccID = CStr(Trim(CInt(Right(RPxx("0" + gcnGi.Execute(strSQL).GetString), 4)) + 1))
    GetAccId = Right("000" & strCompCode, 3) & Format(RightNow, "yyyymmdd") & Right("0000" & strAccID, 4)
End Function

Public Sub subAnd(strCH As String)
  On Error GoTo ChkErr
    If strChoose <> "" Then
        strChoose = strChoose & " And " & strCH
    Else
        strChoose = strCH
    End If
  Exit Sub
ChkErr:
    Call ErrSub("Utility5", "subAnd")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call SubGil

  Exit Sub
ChkErr:
    ErrSub Me.Name, "Form_Load"
End Sub
Private Sub Form_Activate()
    On Error Resume Next
        Screen.MousePointer = vbDefault
End Sub
Private Sub SubGil()
  On Error GoTo ChkErr
    SetgiList gilServiceType, "CodeNo", "Description", "CD046"
    SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode)
    If strCompCode = "" Then gilCompCode.SetCodeNo garyGi(9) Else gilCompCode.SetCodeNo (strCompCode)
    gilCompCode.Query_Description
    If strServiceType = "" Then Call SelectServType(gilCompCode.GetCodeNo, gilServiceType) Else gilServiceType.SetCodeNo (strServiceType)
     gilServiceType.GetDescription
  Exit Sub
ChkErr:
    ErrSub Me.Name, "subGil"
End Sub
Private Function DefaultVal()
  On Error GoTo ChkErr
    Dim rsSO062 As New ADODB.Recordset
    Dim rsSO043 As New ADODB.Recordset  ' 收費參數
    Dim strSQL As String, strSO043 As String
    

        ' 3.初始時，至最近日結參數記錄檔(SO062)，取Type=1的資料
        strSQL = "SELECT * FROM " & GetOwner & "SO062 WHERE TYPE = 3 And CompCode=" & gilCompCode.GetCodeNo & " And ServiceType='" & gilServiceType.GetCodeNo & "'"
        If OpenRecordset(rsSO062, strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False) <> giOK Then
            ErrHandle "開啟 [最近日結參數記錄檔](SO062) 時發生錯誤: " & Err.Description & " 這個程式即將關閉。", False, 2, "Form_Load"
            Unload Me
            Exit Function
        End If
        '收費參數資料
        strSO043 = "Select para35,para36 from " & GetOwner & "SO043 Where CompCode=" & gilCompCode.GetCodeNo & " And ServiceType='" & gilServiceType.GetCodeNo & "'"
        If Not GetRS(rsSO043, strSO043, gcnGi, , adOpenForwardOnly, adLockReadOnly, adUseClient, False) Then Exit Function
        
        ' 若SO062無資料 或無傳入值
        If rsSO062.EOF Then
            ' 則各欄位為空白或初值
            lblTranDate.Caption = ""
            lblUpdTime.Caption = ""
            lblUpdEn.Caption = ""
            lblPara35.Caption = "0"
            lblPara36.Caption = "0"
            gdaEndDate.SetValue ""
        ' 若有資料
        Else
            '結算截止日
'            lblTranDate.Caption = Format(rsSO062("TranDate").Value & "", "ee/MM/DD")
            lblTranDate.Caption = GetDTString(rsSO062("TranDate").Value & "")
            lngTranDate = Val(rsSO062("TranDate").Value & "")
            '上次執行結算時間
            lblUpdTime.Caption = rsSO062("UpdTime").Value & ""
            '上次執行結算人員
            lblUpdEn.Caption = rsSO062("UpdEn").Value & ""
            'PPV結算作業金額限額 PPV結算作業月數限額
            intPara35 = rsSO043("Para35").Value & ""
            lblPara35.Caption = rsSO043("Para35").Value & ""
            intPara36 = rsSO043("Para36").Value & ""
            lblPara36.Caption = rsSO043("Para36").Value & ""
            gdaEndDate.SetValue DateAdd("d", 1, rsSO062("TranDate").Value & "")
            
        End If
        rsSO062.Close
        Set rsSO062 = Nothing
        
        If blnGive Then
            txtCustID = strCustid
            If strCustid <> "" Then txtCustID.Enabled = False
            txtSrNO = strSrNO
            If strSrNO <> "" Then txtSrNO.Enabled = False
            gdaEndDate.SetValue strEndDate
            If strEndDate <> "" Then gdaEndDate.Enabled = False
            If intPara35In <> 0 Then lblPara35.Caption = intPara35In Else lblPara35.Caption = intPara35
            If intPara36In <> 0 Then lblPara36.Caption = intPara36In Else lblPara36.Caption = intPara36
            chkTran1.Value = 1
            chkTran2.Value = 1
            chkRun.Value = 1
            If strCompCode <> "" Then gilCompCode.Enabled = False
            If strServiceType <> "" Then gilServiceType.Enabled = False
        End If
   Exit Function
ChkErr:
   ErrSub Me.Name, "DefaultVal"
End Function

Private Sub gilServiceType_Change()
  On Error GoTo ChkErr
    DefaultVal
  Exit Sub
ChkErr:
   ErrSub Me.Name, "gilServiceType_Change"
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Dim strSQL As String

    If Len(gilCompCode.GetCodeNo & "") = 0 Then Exit Sub
    ' 公司別：若有值且離開此欄位時
    If rsCD039.State = adStateOpen Then rsCD039.Close
    strSQL = "SELECT VouPath,CODENO,EmcCompCode,AccountVer FROM " & GetOwner & "CD039 Where CODENO = " & gilCompCode.GetCodeNo & " ORDER BY CODENO"
    If OpenRecordset(rsCD039, strSQL, gcnGi, adOpenStatic, adLockReadOnly, adUseClient, False) <> giOK Then
        ErrHandle "在開啟 [公司別代碼檔](CD039) 時發生錯誤: " & Err.Description & " 這個程式即將關閉!", False, 2, "gilCompCode_Validate"
        End
        Exit Sub
    End If
     Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
     DefaultVal
  Exit Sub
ChkErr:
    ErrSub Me.Name, "gilCompCode_Change"
End Sub

Private Sub gilCompCode_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then MsgMustBe ("公司別"): Cancel = True
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Public Property Let uCompCode(ByVal vData As String)
    strCompCode = vData
    blnGive = True
End Property

Public Property Let uServiceType(ByVal vData As String)
    strServiceType = vData
    blnGive = True
End Property

Public Property Let uCustId(ByVal vData As String)
    strCustid = vData
    blnGive = True
End Property

'STB設備序號
Public Property Let uSrNO(ByVal vData As String)
    strSrNO = vData
    blnGive = True
End Property

'認證編號
Public Property Let uAuth(ByVal vData As String)
    strAuth = vData
End Property

'結算日期期限
Public Property Let uEndDate(ByVal vData As String)
    strEndDate = vData
End Property

'PPV結算作業金額限額
Public Property Let uPara35(ByVal vData As Integer)
    intPara35In = vData
End Property

'PPV結算作業月數限額
Public Property Let uPara36(ByVal vData As Integer)
    intPara36In = vData
End Property

Public Property Get uMsg() As Variant
    uMsg = strMsg
End Property

Public Property Get uDo() As Boolean
    uDo = blnDO
End Property
