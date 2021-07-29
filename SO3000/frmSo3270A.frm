VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.2#0"; "GiDate.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSo3270A 
   Caption         =   "臨櫃代收資料過帳作業 [SO327..A]"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   6705
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   345
      Left            =   3870
      TabIndex        =   6
      Top             =   1530
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.確定"
      Height          =   375
      Left            =   1650
      TabIndex        =   5
      Top             =   1530
      Width           =   1245
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   2790
      TabIndex        =   4
      Top             =   870
      Width           =   2925
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5760
      TabIndex        =   3
      Top             =   870
      Width           =   435
   End
   Begin Gi_Date.GiDate gdaRealDate 
      Height          =   345
      Left            =   2100
      TabIndex        =   2
      Top             =   300
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
   Begin MSComDlg.CommonDialog comdPath 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblAccount 
      AutoSize        =   -1  'True
      Caption         =   "（若空白，則以扣帳日為入帳日）"
      Height          =   180
      Left            =   3330
      TabIndex        =   7
      Top             =   390
      Width           =   2700
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      Caption         =   "資料來源（路徑+檔名）："
      Height          =   180
      Left            =   660
      TabIndex        =   1
      Top             =   960
      Width           =   2070
   End
   Begin VB.Label lblDate1 
      AutoSize        =   -1  'True
      Caption         =   "收款入帳日期："
      Height          =   180
      Left            =   720
      TabIndex        =   0
      Top             =   390
      Width           =   1260
   End
End
Attribute VB_Name = "frmSo3270A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private intPara5 As Integer
Private intPara8 As Integer
Private rs As New ADODB.Recordset
Private strItemName As String
Private strnewtime As String
Private fso As New FileSystemObject
Private file As TextStream

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsDataOk = False Then Exit Sub

On Error GoTo ChkErr
Set file = fso.OpenTextFile(Trim(txtPath.Text), ForReading)
Call IntoAccount
Exit Sub
ChkErr:
    Resume 0
    Call ErrSub(Me.Name, "cmdOK_Click()")
End Sub

Private Sub IntoAccount()
Dim strRealAmt As String, strReceiver As String
Dim strRealDate As String, strBillNO As String
Dim strmsg As String, rsTmp As New ADODB.Recordset
Dim intErrCount As Integer, intCount As Integer
Dim rs As New ADODB.Recordset
Dim strNowTime As String
strnewtime = Date
Dim strSQL1 As String, strSQL2 As String

Do While Not file.AtEndOfLine
strmsg = file.ReadLine
Debug.Print strmsg
strRealAmt = "": strReceiver = "": strRealDate = "": strBillNO = ""
strRealAmt = Val(Mid(strmsg, 41, 14))
strReceiver = Mid(strmsg, 10, 8)
strRealDate = Mid(strmsg, 21, 6)
If strRealDate <> "" Then
    strRealDate = str(1911 + Val(Mid(strRealDate, 1, 2))) & Mid(strRealDate, 3, 4)
End If
strBillNO = Mid(strmsg, 65, 15)
If Mid(strmsg, 1, 1) <> 2 Then GoTo CLoop
If rsTmp.State = adStateOpen Then rsTmp.Close
 Call OpenRecordset(rsTmp, "Select sum(RealAmt) as TotalRealAmt,Sum(ShouldAmt) as TotalShouldAmt From so033 Where billno='" & strBillNO & "'", gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False)
 If rsTmp("TotalRealAmt").Value > 0 Then
    
    
 Else
    If rsTmp("TotalShouldAmt").Value <> strRealAmt Then
        
    
    Else
        Call OpenRecordset(rs, "Select Rowid,CustId,ShouldAmt,CitemCode,realStartDate,RealStopDate,RealPeriod From So033 Where Billno='" & Trim(strBillNO) & "'", gcnGi, adOpenKeyset, adLockOptimistic, adUseClient, False)
        If rs.RecordCount > 0 Then rs.MoveFirst
        While Not rs.EOF
            strSQL1 = "Update So033 Set RealAmt=" & strRealAmt & ",Clcten='" & strReceiver & "',ClctName='臨櫃代收',RealDate=To_Date('" & strRealDate & "','YYYYMMDD'),UpdTime='" & strnewtime & "', UpdEn ='" & garyGi(1) & "',UCCode=Null,Cmcode=5,CMName='" & strItemName & "' Where Rowid='" & rs("RowId").Value & "'"
            'strSQL2
            gcnGi.Execute strSQL1
            rs.MoveNext
        Wend
    End If
 End If
CLoop:
Loop
End Sub


Private Sub cmdPath_Click()
    comdPath.Filter = "Text (*.txt)|*.txt|所有檔案 (*.*)|*.*"
    comdPath.ShowOpen
    txtPath.Text = comdPath.FileName
End Sub

Private Function IsDataOk() As Boolean
On Error GoTo ChkErr
IsDataOk = False
    If Dir(Trim(txtPath.Text)) = "" Then MsgBox "代收資料不存在，請重新選取！", vbExclamation, "訊息！": Exit Function
IsDataOk = True
Exit Function
ChkErr:
    Call ErrSub(Me.Name, "IsdataOK")
End Function


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    If Alfa2 = True Then
        Call GetGlobal
    End If
    Call GetInitData
End Sub

Private Sub DefineRs()
    With rs
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "RealDate", adBSTR, 8
        .Fields.Append "FilePath", adBSTR, 40
        .Open
    End With
End Sub

Private Sub GetInitData()
    Dim rsTmp  As New ADODB.Recordset
    If OpenRecordset(rsTmp, "Select Para5,Para8 from So043", gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False) <> giOK Then MsgBox ConnectErr, vbExclamation, "訊息！": Exit Sub
    If rsTmp.EOF = False Then
        intPara5 = rsTmp("Para5").Value
        intPara8 = rsTmp("Para8").Value
    End If
    Dim strFilePath As String
    strFilePath = ReadGICMIS1("ErrLogPath")
    If Dir(strFilePath & "\AgencyIn.log") = "" Then
        Call DefineRs
    Else
        rs.Open strFilePath & "\AgencyIn.log"
        gdaRealDate.SetValue rs("RealDate").Value & ""
        txtPath.Text = rs("FilePath").Value & ""
    End If
    If rsTmp.State = adStateOpen Then rsTmp.Close
    If OpenRecordset(rsTmp, "Select Description From CD031 where CodeNo=5", gcnGi, adOpenForwardOnly, adLockReadOnly, adUseClient, False) <> giOK Then MsgBox ConnectErr, vbExclamation, "訊息！": Exit Sub
    If rsTmp.EOF = False Then
        strItemName = rsTmp("Description").Value & ""
    End If

End Sub

