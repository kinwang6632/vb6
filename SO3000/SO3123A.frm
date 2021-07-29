VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Begin VB.Form frmSO3123A 
   BorderStyle     =   1  '單線固定
   Caption         =   "應收區域及道路分析 [SO3123A]"
   ClientHeight    =   3990
   ClientLeft      =   2790
   ClientTop       =   2940
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3123A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CheckBox chkAddCMCode 
      Caption         =   "分組加收費方式"
      Height          =   195
      Left            =   3900
      TabIndex        =   16
      Top             =   600
      Width           =   2265
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   525
      Left            =   2550
      TabIndex        =   5
      Top             =   3330
      Width           =   1515
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   525
      Left            =   5220
      TabIndex        =   6
      Top             =   3330
      Width           =   1305
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      Height          =   525
      Left            =   150
      TabIndex        =   4
      Top             =   3330
      Width           =   1275
   End
   Begin prjGiList.GiList giloldClctEn1 
      Height          =   315
      Left            =   990
      TabIndex        =   0
      Top             =   970
      Width           =   2625
      _ExtentX        =   4630
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
      FldWidth1       =   800
      FldWidth2       =   1500
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList giloldClctEn2 
      Height          =   315
      Left            =   3900
      TabIndex        =   1
      Top             =   970
      Width           =   2625
      _ExtentX        =   4630
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
      FldWidth1       =   800
      FldWidth2       =   1500
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   990
      TabIndex        =   12
      Top             =   90
      Width           =   2625
      _ExtentX        =   4630
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
      FldWidth1       =   600
      FldWidth2       =   1700
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   990
      TabIndex        =   13
      Top             =   530
      Width           =   2625
      _ExtentX        =   4630
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
      FldWidth1       =   600
      FldWidth2       =   1700
      F2Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilClctEn1 
      Height          =   315
      Left            =   990
      TabIndex        =   2
      Top             =   1365
      Width           =   2625
      _ExtentX        =   4630
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
      FldWidth1       =   800
      FldWidth2       =   1500
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilClctEn2 
      Height          =   315
      Left            =   3900
      TabIndex        =   3
      Top             =   1365
      Width           =   2625
      _ExtentX        =   4630
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
      FldWidth1       =   800
      FldWidth2       =   1500
      F5Corresponding =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '透明
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3660
      TabIndex        =   18
      Top             =   1400
      Width           =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '透明
      Caption         =   "收費人員"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1365
      Width           =   915
   End
   Begin VB.Label lblServiceType 
      AutoSize        =   -1  'True
      Caption         =   "服務類別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   180
      TabIndex        =   15
      Top             =   560
      Width           =   780
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      Caption         =   "公司別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   330
      TabIndex        =   14
      Top             =   150
      Width           =   585
   End
   Begin VB.Label lblClctEn 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '透明
      Caption         =   "原收費人員"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   990
      Width           =   1035
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  '透明
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3660
      TabIndex        =   10
      Top             =   1000
      Width           =   120
   End
   Begin VB.Line Line1 
      X1              =   30
      X2              =   6810
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "說明:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   90
      TabIndex        =   9
      Top             =   1920
      Width           =   510
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "1.當收費員有輸入時,表分析此收費員之資料;反之,則分析全部收費員"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   210
      TabIndex        =   8
      Top             =   2220
      Width           =   5730
   End
   Begin VB.Label lbl4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "2.以收費員為主要索引;行政區及街道為輔做資料統計."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   210
      TabIndex        =   7
      Top             =   2640
      Width           =   4515
   End
End
Attribute VB_Name = "frmSO3123A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cnn As New ADODB.Connection
Private strChoose As String
Private strChooseString As String
Private RptObj As Object
Private strForm As String


Private Sub cmdExit_Click()
  On Error GoTo ChkErr
   Unload Me
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub cmdPrevRpt_Click()
    On Error GoTo ChkErr
        Screen.MousePointer = vbHourglass
        If chkAddCMCode = 1 Then
            Call PreviousRpt(GetPrinterName(5), "SO3123A.Rpt", Me.Caption)
        Else
            Call PreviousRpt(GetPrinterName(5), "SO3123A2.Rpt", Me.Caption)
        End If
        Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ChkErr
    Dim strSQL As String
    Dim lngAffected As Long
        Screen.MousePointer = vbHourglass
        cmdExit.SetFocus
        If Not ChkDTok Then Exit Sub
        ReadyGoPrint
        Call subChoose
        If Not subInsertMDB(lngAffected) Then Exit Sub
        If lngAffected = 0 Then
            MsgNoRcd
        Else
            If chkAddCMCode = 1 Then
                strSQL = "SELECT * FROM SO3123A2"
                Call PrintRpt(GetPrinterName(5), "SO3123A2.rpt", "SO3123", Me.Caption, strSQL, strChooseString, "原收費員", True, "Tmp0000.MDB")
            Else
                strSQL = "SELECT * FROM SO3123A"
                Call PrintRpt(GetPrinterName(5), "SO3123A.rpt", "SO3123", Me.Caption, strSQL, strChooseString, "原收費員", True, "Tmp0000.MDB")
            End If
        End If
        Screen.MousePointer = vbDefault
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function subInsertMDB(ByRef lngAffected As Long) As Boolean
    On Error GoTo ChkErr
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim strField As String, strGroup As String
    Dim strTable As String
        strField = "SELECT CLCTEN,CLCTNAME,CD001.Description AreaName,CD017.Description StrtName"
        If chkAddCMCode = 1 Then
            strField = strField & ",CMCode,CMName"
            strTable = "SO3123A2"
            strGroup = ",CMCode,CMName"
        Else
            strTable = "SO3123A"
        End If
        strSQL = strField & ",Count(*) as intCount From " & GetOwner & "CD017 CD017," & GetOwner & IIf(strForm = "SO3123", "SO032 ", "SO033 ") & " SO032 ," & GetOwner & "CD001 CD001 Where " & _
        " So032.StrtCode=CD017.CodeNo And So032.AreaCode = CD001.CodeNo " & IIf(strChoose <> "", " And ", "") & strChoose & "   Group By CLCTEN,CLCTNAME,CD001.Description,CD017.Description" & strGroup
        
        If Not GetRS(rsTmp, strSQL) Then Exit Function
        If Not CreateMDBTable(rsTmp, strTable, cnn) Then Exit Function
        SendSQL rsTmp.Source, True
        cnn.BeginTrans
        On Error GoTo ErTrans
        While Not (rsTmp.BOF Or rsTmp.EOF)
            cnn.Execute GetTmpMDBExecuteStr(rsTmp, strTable)
            rsTmp.MoveNext
            DoEvents
        Wend
        lngAffected = rsTmp.RecordCount
        cnn.CommitTrans
        Call Close3Recordset(rsTmp)
        subInsertMDB = True
    Exit Function
ErTrans:
    cnn.RollbackTrans
ChkErr:
    ErrSub Me.Name, "subInsertMDB"
End Function


Private Sub subChoose()
    On Error GoTo ChkErr
        strChoose = " ClctYM is Not Null "
        strChooseString = ""
        If gilCompCode.GetCodeNo <> "" Then Call subAnd2(strChoose, "SO032.CompCode =" & gilCompCode.GetCodeNo)
        If gilServiceType.GetCodeNo <> "" Then Call subAnd2(strChoose, "SO032.ServiceType ='" & gilServiceType.GetCodeNo & "'")
        '#1659  950517原"收費人員"改為"原收費人員"，另新增一"收費人員"
        If Not giloldClctEn1.GetCodeNo = "" Then
           strChoose = strChoose & " And OldClctEn>='" & giloldClctEn1.GetCodeNo & "'"
        End If
        If Not giloldClctEn2.GetCodeNo = "" Then
           strChoose = strChoose & " And OldClctEn<='" & giloldClctEn2.GetCodeNo & "'"
        End If
        If Not gilClctEn1.GetCodeNo = "" Then
           strChoose = strChoose & " And ClctEn>='" & gilClctEn1.GetCodeNo & "'"
        End If
        If Not gilClctEn2.GetCodeNo = "" Then
           strChoose = strChoose & " And ClctEn<='" & gilClctEn2.GetCodeNo & "'"
        End If
        
        strChooseString = "原收費人員: " & subSpace(giloldClctEn1.GetDescription) & "~" & subSpace(giloldClctEn2.GetDescription) & _
                          "收費人員: " & subSpace(gilClctEn1.GetDescription) & "~" & subSpace(gilClctEn2.GetDescription)
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
     Set cnn = GetTmpMdbCn
     Call subGil
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ChkErr
        If Not ChkGiList(KeyCode) Then Exit Sub
        Call FunctionKey(KeyCode, Shift)
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub

Public Sub FunctionKey(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Select Case KeyCode
           '----------------------------------------------------
           Case vbKeyF5 '   F5: 列印 相當於按下cmdPrint
                        If cmdPrint.Enabled = True Then cmdPrint.Value = True
           Case vbKeyEscape
                        If cmdExit.Enabled = True Then cmdExit.Value = True
           '----------------------------------------------------
    End Select
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "FunctionKey")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    SetgiList gilClctEn1, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True
    SetgiList gilClctEn2, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True
    SetgiList giloldClctEn1, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True
    SetgiList giloldClctEn2, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True
    SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
    SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGil")
End Sub
  
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        Set cnn = Nothing
End Sub

Public Property Get uForm() As String
    On Error GoTo ChkErr
        uForm = strForm
    Exit Property
ChkErr:
    ErrSub Me.Name, "Get uForm"
End Property

Public Property Let uForm(ByVal vForm As String)
    On Error GoTo ChkErr
        strForm = vForm
    Exit Property
ChkErr:
    ErrSub Me.Name, "Let uForm"
End Property

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)

End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3120", "SO3123") Then Exit Sub
        Call subGil
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.SetCodeNo ""
        gilServiceType.Query_Description
        gilServiceType.ListIndex = 1
        
End Sub

