VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Begin VB.Form frmSO1156A 
   Caption         =   "PPV�O�ε���@�~[SO1156A]"
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
   StartUpPosition =   2  '�ù�����
   Begin VB.TextBox txtAuthenticId 
      BeginProperty Font 
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
      Caption         =   "F2. �T�w"
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
      Height          =   375
      Left            =   135
      TabIndex        =   9
      Top             =   6435
      Width           =   1365
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
      Height          =   375
      Left            =   1875
      TabIndex        =   10
      Top             =   6435
      Width           =   1365
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "�W���έp���G"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
         Caption         =   "���浲��@�~"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Caption         =   "���ͥ����浲����ӳ���"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Caption         =   "���͹w�p������ӳ���"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Value           =   1  '�֨�
         Width           =   2385
      End
   End
   Begin VB.Frame fraLast 
      Caption         =   "�W������O��"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
         Caption         =   "����I���:"
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
            Name            =   "�s�ө���"
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
         Caption         =   "�W�����浲��ɶ�:"
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
            Name            =   "�s�ө���"
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
         Caption         =   "�W�����浲��H��: "
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
            Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
      Caption         =   "�{�ҽs��"
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
      Left            =   1320
      TabIndex        =   30
      Top             =   3600
      Width           =   780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "(�t���餧�e����ƱN�Q����)"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Caption         =   "�Ƚs"
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
      Caption         =   "PPV����@�~��ƭ��B"
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
      Left            =   120
      TabIndex        =   24
      Top             =   2880
      Width           =   1905
   End
   Begin VB.Label lbl4 
      AutoSize        =   -1  'True
      Caption         =   "PPV����@�~���B���B"
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
      Left            =   120
      TabIndex        =   23
      Top             =   2520
      Width           =   1905
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "���q�O"
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
      Left            =   1260
      TabIndex        =   22
      Top             =   1605
      Width           =   585
   End
   Begin VB.Label lblSrNO 
      AutoSize        =   -1  'True
      Caption         =   "STB�]�ƧǸ�"
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
      Left            =   975
      TabIndex        =   21
      Top             =   3960
      Width           =   1110
   End
   Begin VB.Label lblServicetype 
      AutoSize        =   -1  'True
      Caption         =   "�A�����O"
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
Dim lngTranDate As Long             ' ����I���
Dim rsCD039 As New ADODB.Recordset  ' ���q�O�N�X��
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
Dim strAuth As String   '�{�ҽs��
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
    If gdaEndDate.GetValue = "" Then MsgBox "�ж�J����������!", vbInformation, "�T��!": Exit Sub
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
            MsgBox "�L����I���!", vbInformation, "�T��": strMsg = "�L����I���!": If blnGive = True Then Unload Me: Exit Sub Else Exit Sub
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
                strMsg = "�L�w�p���浲����"
              Else
                strMsg = "�L�����浲����"
              End If
              MsgBox strMsg, vbExclamation, "����"

'              MsgNoRcd
              SendSQL , , True
          Else
              strSQL = "SELECT * From " & strViewName & " V"
              strFormula = strFormula & ";Sort0={V.CustId}"
             Call PrintRpt(GetPrinterName(5), "SO1156A.rpt", , "�w�p������ӳ��� [SO1156A]", strSQL, strChooseString, , True, , , strFormula)

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
        strCH = "���q�O=" & strCompCode & " ;�A�����O=" & strServiceType & " ; ����������=" & strEndDate & "; �Ƚs=" & strCustid & "; STB�]�ƧǸ�=" & strSrNO & "; �{�ҽs��=" & strCustid
        gcnGi.Execute "Update " & GetOwner & "SO062 Set TranDate=TO_DATE('" & strEndDate & "','YYYY/MM/DD') " & _
                      ",UpdEn='" & garyGi(1) & "',UpdTime='" & GetDTString(RightNow) & "',Para='" & strCH & _
                      "' Where Type=3 And CompCode='" & strCompCode & "' AND ServiceType='" & strServiceType & "'"
        strMsg = "����@�~�����I"
        blnDO = True
    Else
        strMsg = "�L�����ơI"
        blnDO = True
    End If
    NewSO062 = True
Exit Function
ChkErr:
    Call ErrSub(Me.Name, "NewSO062")
End Function

'��������⤧PPV�q�ʸ��
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
           MsgBox strMsg, vbInformation, "�T���I"
        End If
        subChoose = True
    End If
    strChooseString = "���q�O�@: " & subSpace(strCompCode) & ";" & _
                      "�A�����O: " & subSpace(strServiceType) & ";" & _
                      "�����������@: " & subSpace(strEndDate) & ";" & _
                      "PPV����@�~���B���B: " & intPara35 & ";" & _
                      "PPV����@�~��ƭ��B: " & intPara36
                      
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
    '�إ߼Ȧs��
    CreateTable
    'rsSO105Amt�O��X�ӫȤ�λ{�ҽs�����`���B
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
        '�A�Ѩ�Ƚs�Ť��q�O�{�ҽs���h�d�ߨ���z���
        Do While Not rsSO105Date.EOF

            '���`���B�֩�Pare35���B���B�B���z����쵲�����䶡�p��W�LPara36����ƭ��B�ɤ����p��
            dblMounth = CInt((CDate(Format(strEndDate, "####/##/##")) - CDate(Format(rsSO105Date("AcceptTime") & "", "YYYY/MM/DD"))) / 30)
            
            If rsSO105Amt("Amount") < intPara35 Or dblMounth > intPara36 Then
                rsSO105Date.MoveNext: Exit Do
'                If rsSO105Date.EOF Then
            End If
            '�d�X������bCD048���u�f�N�XCodeNo
            strCD048 = "select * From " & GetOwner & "CD048 WHERE PeriodType=3 And to_date('" & CDate(Format(rsSO105Date("AcceptTime"), "YYYY/MM/DD")) & "','YYYY/MM/DD') >= StartDate AND to_date('" & CDate(Format(rsSO105Date("AcceptTime"), "YYYY/MM/DD")) & "','YYYY/MM/DD')<= StopDate +1"
            If Not GetRS(rsCD048, strCD048, gcnGi, , adOpenForwardOnly, adLockReadOnly, adUseClient, False) Then Exit Function
            If Not rsCD048.EOF Then
                strCD04CodeNO = rsCD048("CodeNo") & ""
            Else
                strCD04CodeNO = ""
            End If
            '�ŦX��ƥ�J�Ȧs
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

'�馩�v�p��
Private Function DisCountTest(ByVal strCompCode As String, ByVal strServiceType As String, ByVal strEndDate As String) As Boolean
 On Error GoTo ChkErr
 Dim rsTmp As New ADODB.Recordset
 Dim rsCD048 As New ADODB.Recordset
 Dim rsSumB As New ADODB.Recordset
 Dim rsCountT As New ADODB.Recordset    '�ӫȤ�so105g �b�_�������(�]�t����)������=����
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
    '�q�u�f�N�X��XCD048��48A�����
    strCD048 = "Select A.*,B.* from " & GetOwner & "CD048 A," & GetOwner & "CD048A B Where A.CodeNo=B.CodeNo And A.CodeNo='" & rsTmp("48CodeNo") & "'"
    If Not GetRS(rsCD048, strCD048, gcnGi, , adOpenForwardOnly, adLockReadOnly, adUseClient, False) Then Exit Function
    
        '��X�ӫȤ�so105g �b��������������⪺���
        '��������
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
        '�P�_������ƬO�_�֭p�e��
        If rsCD048("GrandTotalFlag") = 1 Then '�֭p�e��
        '�֭p�e�����`����
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
        '��X�ӫȤ�so105g �b�_�������(�]�t����)������=����
        '�֭p�e�����`����
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
            dblSum = rsTmp("AMT")        '�������B
            dblCNT = rsCount("Count")    '��������
            
        End If
        '�P�_�u�f���� 2=�馩�v
        If rsCD048("DiscountType") = 2 Then
            '�P�_�w�q���� 1=���O���B
            If rsCD048("DefineType") = 1 Then
                '�P�_�O�_�W�X���O���B
                If dblSum > rsCD048("TotalAMT") Then
                    If rsCD048("GrandTotalFlag") = 1 Then '�֭p�e����,�W�L�`���B��H������B���馩
                            dblDisCAmt = CInt(rsTmp("AMT") * (1 - (rsCD048("Rate") / 100)))
                    Else
                    '�P�_�A�έ�h�`�� 1=�X�p���B
                        If rsCD048("UseType") = 1 Then
                            dblDisCAmt = CInt(rsTmp("AMT") * (1 - (rsCD048("Rate") / 100)))
                        ElseIf rsCD048("UseType") = 2 Then '2=�W�L���B
                            If (rsTmp("AMT") - rsCD048("TotalAMT")) > 0 Then dblDisCAmt = CInt((rsTmp("AMT") - rsCD048("TotalAMT")) * (1 - (rsCD048("Rate") / 100))) Else dblDisCAmt = 0
                        End If
                    End If
                Else
                    dblDisCAmt = 0
                End If
            '�P�_�w�q���� 2=���O����
            ElseIf rsCD048("DefineType") = 2 Then
                If dblCNT >= rsCD048("TotalCNT") Then '�P�_�O�_�W�X���O����
                    If rsCD048("GrandTotalFlag") = 1 Then '�֭p�e����,�W�L���ƫ�H������B���馩
                            dblDisCAmt = CInt(rsTmp("AMT") * (1 - (rsCD048("Rate") / 100)))
                    Else
                        If rsCD048("UseType") = 1 Then
                            dblDisCAmt = CInt(rsTmp("AMT") * (1 - (rsCD048("Rate") / 100)))
                        ElseIf rsCD048("UseType") = 2 Then '2=�W�L���B
                            If rsCD048("GrandTotalFlag") = 1 Then '�֭p�e����,�W�L�`���B��H������B���馩
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
        '�P�_�u�f���� 1=�T�w���B
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
    
    '�̱����X��insert SO130 PPV��I����ɪ����
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
        '����T��
'         Call CreaSO033(rsSO105("Custid"), rsSO105("CitemCode"))
        Call CreaSO033(rsSO105("Custid"), strAuthenticID, dblTotalAmt, strID, strDateNow, strCompCode, strServiceType)
    rsTmp.MoveNext
 Wend
 DisCountTest = True
  Exit Function
ChkErr:
    Call ErrSub(Me.Name, "DisCountTest")
End Function

'����T��
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
        rsSO033("PTName").Value = "�{��"
        rsSO033("ClassCode").Value = CInt(gcnGi.Execute("Select ClassCode1 FROM " & GetOwner & "SO001 Where CustId=" & intCustid).GetString)
        rsSO033("FaciSeqNo").Value = intCustid
        rsSO033("ServiceType").Value = strServiceType
        rsSO033("Note").Value = "PPV �O�ε��ⲣ�ͪ����O���"
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
        '�p�G�a�}�̼@�����O�a�}
        If gcnGi.Execute("Select PARA14 FROM " & GetOwner & "SO043 Where COMPCODE='" & strCompCode & "' And ServiceType='" & strServiceType & "'").GetString = 1 Then
            strSQL = "SELECT B.* FROM " & GetOwner & "SO001 A," & GetOwner & "SO014 B WHERE A.ChargeAddrNo=B.AddrNo AND A.CUSTID=" & intCustid
        Else    '�˾��a�}
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
    Dim rsSO043 As New ADODB.Recordset  ' ���O�Ѽ�
    Dim strSQL As String, strSO043 As String
    

        ' 3.��l�ɡA�̪ܳ�鵲�ѼưO����(SO062)�A��Type=1�����
        strSQL = "SELECT * FROM " & GetOwner & "SO062 WHERE TYPE = 3 And CompCode=" & gilCompCode.GetCodeNo & " And ServiceType='" & gilServiceType.GetCodeNo & "'"
        If OpenRecordset(rsSO062, strSQL, gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False) <> giOK Then
            ErrHandle "�}�� [�̪�鵲�ѼưO����](SO062) �ɵo�Ϳ��~: " & Err.Description & " �o�ӵ{���Y�N�����C", False, 2, "Form_Load"
            Unload Me
            Exit Function
        End If
        '���O�ѼƸ��
        strSO043 = "Select para35,para36 from " & GetOwner & "SO043 Where CompCode=" & gilCompCode.GetCodeNo & " And ServiceType='" & gilServiceType.GetCodeNo & "'"
        If Not GetRS(rsSO043, strSO043, gcnGi, , adOpenForwardOnly, adLockReadOnly, adUseClient, False) Then Exit Function
        
        ' �YSO062�L��� �εL�ǤJ��
        If rsSO062.EOF Then
            ' �h�U��쬰�ťթΪ��
            lblTranDate.Caption = ""
            lblUpdTime.Caption = ""
            lblUpdEn.Caption = ""
            lblPara35.Caption = "0"
            lblPara36.Caption = "0"
            gdaEndDate.SetValue ""
        ' �Y�����
        Else
            '����I���
'            lblTranDate.Caption = Format(rsSO062("TranDate").Value & "", "ee/MM/DD")
            lblTranDate.Caption = GetDTString(rsSO062("TranDate").Value & "")
            lngTranDate = Val(rsSO062("TranDate").Value & "")
            '�W�����浲��ɶ�
            lblUpdTime.Caption = rsSO062("UpdTime").Value & ""
            '�W�����浲��H��
            lblUpdEn.Caption = rsSO062("UpdEn").Value & ""
            'PPV����@�~���B���B PPV����@�~��ƭ��B
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
    ' ���q�O�G�Y���ȥB���}������
    If rsCD039.State = adStateOpen Then rsCD039.Close
    strSQL = "SELECT VouPath,CODENO,EmcCompCode,AccountVer FROM " & GetOwner & "CD039 Where CODENO = " & gilCompCode.GetCodeNo & " ORDER BY CODENO"
    If OpenRecordset(rsCD039, strSQL, gcnGi, adOpenStatic, adLockReadOnly, adUseClient, False) <> giOK Then
        ErrHandle "�b�}�� [���q�O�N�X��](CD039) �ɵo�Ϳ��~: " & Err.Description & " �o�ӵ{���Y�N����!", False, 2, "gilCompCode_Validate"
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
    If gilCompCode.GetCodeNo = "" Then MsgMustBe ("���q�O"): Cancel = True
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

'STB�]�ƧǸ�
Public Property Let uSrNO(ByVal vData As String)
    strSrNO = vData
    blnGive = True
End Property

'�{�ҽs��
Public Property Let uAuth(ByVal vData As String)
    strAuth = vData
End Property

'����������
Public Property Let uEndDate(ByVal vData As String)
    strEndDate = vData
End Property

'PPV����@�~���B���B
Public Property Let uPara35(ByVal vData As Integer)
    intPara35In = vData
End Property

'PPV����@�~��ƭ��B
Public Property Let uPara36(ByVal vData As Integer)
    intPara36In = vData
End Property

Public Property Get uMsg() As Variant
    uMsg = strMsg
End Property

Public Property Get uDo() As Boolean
    uDo = blnDO
End Property
