VERSION 5.00
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.8#0"; "GiMulti.ocx"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{0CE6E66E-92AE-11D3-8311-0080C8453DDF}#3.0#0"; "GiYear.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO5220A 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '單線固定
   Caption         =   "每月各項收費加總統計 [SO5220A]"
   ClientHeight    =   4830
   ClientLeft      =   1950
   ClientTop       =   1530
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO5220A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form16"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdExcel 
      Caption         =   "匯成Excel"
      Height          =   525
      Left            =   3450
      TabIndex        =   11
      Top             =   4200
      Width           =   1245
   End
   Begin Gi_Year.GiYear gyrRealDate 
      Height          =   345
      Left            =   1020
      TabIndex        =   2
      Top             =   870
      Visible         =   0   'False
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   609
      ForeColor       =   255
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
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   525
      Left            =   1710
      TabIndex        =   10
      Top             =   4200
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   525
      Left            =   5670
      TabIndex        =   12
      Top             =   4200
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "F5:列印"
      Default         =   -1  'True
      Height          =   525
      Left            =   90
      TabIndex        =   9
      Top             =   4200
      Width           =   1245
   End
   Begin Gi_Multi.GiMulti gimBillType 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1710
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      ButtonCaption   =   "單  據  類  別"
      DataType        =   2
      DIY             =   -1  'True
   End
   Begin Gi_Multi.GiMulti gimCMCode 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      ButtonCaption   =   "收  費  方  式"
   End
   Begin Gi_YM.GiYM gymRealDate 
      Height          =   345
      Left            =   1020
      TabIndex        =   3
      Top             =   870
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   609
      ForeColor       =   255
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
   Begin VB.Frame fraReportType 
      BackColor       =   &H00E0E0E0&
      Caption         =   "報表種類"
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   90
      TabIndex        =   13
      Top             =   120
      Width           =   2595
      Begin VB.OptionButton optReach 
         BackColor       =   &H00E0E0E0&
         Caption         =   "年度累計"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1350
         TabIndex        =   1
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton optInstant 
         BackColor       =   &H00E0E0E0&
         Caption         =   "當月金額"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   3900
      TabIndex        =   4
      Top             =   270
      Width           =   2940
      _ExtentX        =   5186
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
      F5Corresponding =   -1  'True
   End
   Begin prjGiList.GiList gilServiceType 
      Height          =   315
      Left            =   3900
      TabIndex        =   5
      Top             =   720
      Width           =   2940
      _ExtentX        =   5186
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
      F5Corresponding =   -1  'True
   End
   Begin CS_Multi.CSmulti gimClassCode 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2070
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      ButtonCaption   =   "客   戶   類   別"
   End
   Begin VB.Label lblCompCode 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "公  司  別"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3060
      TabIndex        =   18
      Top             =   330
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "服務類別"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3060
      TabIndex        =   17
      Top             =   780
      Width           =   780
   End
   Begin VB.Label lblRealYear 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收費年份"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label lblHelp 
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
      TabIndex        =   15
      Top             =   2580
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblRealDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "收費年月"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   780
   End
End
Attribute VB_Name = "frmSO5220A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'使用Table: SO033 or SO034 A,CD002 B
Option Explicit
Dim blnExcel As Boolean

Private Sub cmdExcel_Click()
  On Error GoTo ChkErr
    blnExcel = True
    cmdPrint_Click
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExcel_Click")
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
      If optInstant.Value Then
          Call PreviousRpt(GetPrinterName(5), RptName("SO5220", 1), "每月各項收費加總統計(當月金額) [SO5220A1]")
      Else
          Call PreviousRpt(GetPrinterName(5), RptName("SO5220", 2), "每月各項收費加總統計(年度總計) [SO5220A2]")
      End If
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo ChkErr
    If Not ChkDTok Then Exit Sub
    If Not IsDataOk Then Exit Sub
    Screen.MousePointer = vbHourglass
      cmdExit.SetFocus
      Me.Enabled = False
        ReadyGoPrint
        Call subChoose
        Call subPrint
        'Call subInsertMDB
      Me.Enabled = True
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub

Private Function IsDataOk() As Boolean
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then gilCompCode.SetFocus: strErrFile = "公司別": GoTo 66
    If optInstant.Value And gymRealDate.GetValue = "" Then gymRealDate.SetFocus: strErrFile = "收費年月": GoTo 66
    If optReach.Value And gyrRealDate.GetValue = "" Then gyrRealDate.SetFocus: strErrFile = "收費年份": GoTo 66
    IsDataOk = True
    Exit Function
66:
    MsgMustBe (strErrFile)
  Exit Function
ChkErr:
  Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Sub subChoose()
  On Error GoTo ChkErr
    strChoose = " And A.cancelFlag=0 "
    strChooseString = ""
   
    If optInstant.Value Then
      If gymRealDate.GetValue <> "" Then Call subAnd("A.RealDate >= To_Date(" & gymRealDate.GetValue & "01,'yyyymmdd') And A.RealDate < Last_Day(To_Date(" & gymRealDate.GetValue & "01,'yyyymmdd')) +1")
    Else
      If gyrRealDate.GetValue <> "" Then Call subAnd("A.RealDate >= To_Date(" & gyrRealDate.GetValue & "0101,'yyyymmdd') And A.RealDate < To_Date(" & gyrRealDate.GetValue & "1231,'yyyymmdd')+1  ")
    End If
    
  'GiMulti
    If gimBillType.GetQryStr <> "" Then Call subAnd("SubStr(A.BillNo,7,1) " & gimBillType.GetQryStr)
    If gimCMCode.GetQryStr <> "" Then Call subAnd("A.CMCode " & gimCMCode.GetQryStr)
    If gimClassCode.GetQryStr <> "" Then Call subAnd("A.ClassCode " & gimClassCode.GetQryStr)
      
  'GiList
    If gilCompCode.GetCodeNo <> "" Then Call subAnd("A.CompCode=" & gilCompCode.GetCodeNo)
    If gilServiceType.GetCodeNo <> "" Then Call subAnd("A.ServiceType='" & gilServiceType.GetCodeNo & "'")
      
    'strChoose = " And " & strChoose

    strChooseString = IIf(optInstant.Value, "收費年月: " & subSpace(gymRealDate.GetValue(True)), "收費年度: " & subSpace(gyrRealDate.GetValue)) & ";" & _
                      "公司別　: " & subSpace(gilCompCode.GetDescription) & ";" & _
                      "服務類別: " & subSpace(gilServiceType.GetDescription) & ";" & _
                      "收費方式: " & subSpace(gimCMCode.GetDispStr) & ";" & _
                      "單據類別: " & subSpace(gimBillType.GetDispStr) & ";" & _
                      "客戶類別: " & subSpace(gimClassCode.GetDispStr)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subChoose")
End Sub

'Private Sub subInsertMDB()
'  On Error GoTo ChkErr
'  Dim rsTmp As New adodb.Recordset
'  Dim StrTableName As String
'    rsTmp.CursorLocation = adUseClient
'    If rsTmp.State = 1 Then rsTmp.Close
'    cnn.BeginTrans
'    If optInstant.Value Then
'       cnn.Execute "Delete From SO5220A1"
'       strSql = "Select B.Description as ServName,A.CitemName as CitemName,to_char(A.RealDate,'DD') as RealDate,Sum(A.RealAmt) as SumAmt From " & _
'                "CD002 B,SO033 A Where A.ServCode=B.CodeNo " & strChoose & " Group By B.Description,A.CitemName,To_Char(A.RealDate,'DD') UNION All " & _
'                "Select B.Description as ServName,A.CitemName as CitemName,to_char(A.RealDate,'DD') as RealDate,Sum(A.RealAmt) as SumAmt From " & _
'                "CD002 B,SO034 A Where A.ServCode=B.CodeNo " & strChoose & " Group By B.Description,A.CitemName,To_Char(A.RealDate,'DD')"
'
'       Set rsTmp = gcnGi.Execute(strSql)
'       SendSQL strSql, True
'       While Not rsTmp.EOF
'            cnn.Execute "Insert Into SO5220A1 (ServName,CitemName,RealDate,SumAmt) Values ('" & _
'                        rsTmp("ServName") & "','" & _
'                        rsTmp("CitemName") & "','" & _
'                        rsTmp("RealDate") & "'," & _
'                        IIf(IsNull(rsTmp("SumAmt")), 0, rsTmp("SumAmt")) & ")"
'            rsTmp.MoveNext
'            DoEvents
'       Wend
'       StrTableName = "SO5220A1"
'    Else
'       cnn.Execute "Delete From SO5220A2"
'       strSql = "Select B.Description as ServName,A.CitemName as CitemName,To_char(A.RealDate,'YYYYMM') as RealDate,Sum(A.RealAmt) As SumAmt " & _
'                "From SO033 A,CD002 B Where A.ServCode=B.CodeNo " & strChoose & _
'                " Group By B.Description,A.CitemName,To_char(A.RealDate,'YYYYMM') UNION All " & _
'                "Select B.Description as ServName,A.CitemName as CitemName,To_char(A.RealDate,'YYYYMM') as RealDate,Sum(A.RealAmt) As SumAmt " & _
'                "From SO034 A,CD002 B Where A.ServCode=B.CodeNo " & strChoose & _
'                " Group By B.Description,A.CitemName,To_char(A.RealDate,'YYYYMM') "
'
'       Set rsTmp = gcnGi.Execute(strSql)
'       SendSQL strSql, True
'       While Not rsTmp.EOF
'            cnn.Execute "Insert Into SO5220A2 (ServName,CitemName,RealDate,SumAmt) Values ('" & _
'                        rsTmp("ServName") & "','" & _
'                        rsTmp("CitemName") & "','" & _
'                        Right(rsTmp("RealDate"), 2) & "'," & _
'                        IIf(IsNull(rsTmp("SumAmt")), 0, rsTmp("SumAmt")) & ")"
'
'            rsTmp.MoveNext
'            DoEvents
'       Wend
'       StrTableName = "SO5220A2"
'    End If
'    cnn.CommitTrans
'
'    If rsTmp.State = 1 Then rsTmp.Close
'    rsTmp.Open "SELECT Count(*) as intCount FROM " & StrTableName, cnn, adOpenForwardOnly, adLockReadOnly
'    If rsTmp("intCount") = 0 Then
'       MsgNoRcd
'    Else
'       strSql = "SELECT * FROM " & StrTableName
'       If optInstant.Value Then
'          Call PrintRpt(GetPrinterName(5), RptName("SO5220", 1), , "每月各項收費加總統計[當月金額]", strSql, strChooseString, , True, "Tmp0000.Mdb", , , GiPaperLandscape)
'       Else
'          Call PrintRpt(GetPrinterName(5), RptName("SO5220", 2), , "每月各項收費加總統計[年度總計]", strSql, strChooseString, , True, "Tmp0000.Mdb", , , GiPaperLandscape)
'       End If
'    End If
'    Set rsTmp = Nothing
'  Exit Sub
'ChkErr:
'    cnn.RollbackTrans
'    Call ErrSub(Me.Name, "subInsertMDB")
'End Sub

Private Sub subPrint()
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
  Dim rsExcel As New ADODB.Recordset
    If optInstant.Value Then
       strSQL = "Select ServName,CitemName,Sum(RealDate01) Date1,Sum(RealDate02) Date2,Sum(RealDate03) Date3," & _
                "Sum(RealDate04) Date4,Sum(RealDate05) Date5,Sum(RealDate06) Date6,Sum(RealDate07) Date7,Sum(RealDate08) Date8,Sum(RealDate09) Date9," & _
                "Sum(RealDate10) Date10,Sum(RealDate11) Date11,Sum(RealDate12) Date12,Sum(RealDate13) Date13,Sum(RealDate14) Date14,Sum(RealDate15) Date15,Sum(RealDate16) Date16,Sum(RealDate17) Date17," & _
                "Sum(RealDate18) Date18,Sum(RealDate19) Date19,Sum(RealDate20) Date20,Sum(RealDate21) Date21,Sum(RealDate22) Date22,Sum(RealDate23) Date23,Sum(RealDate24) Date24,Sum(RealDate25) Date25," & _
                "Sum(RealDate26) Date26,Sum(RealDate27) Date27,Sum(RealDate28) Date28,Sum(RealDate29) Date29,Sum(RealDate30) Date30,Sum(RealDate31) Date31,Sum(SumAmt) SumAmt From (Select ServName,CitemName," & _
                "Decode(RealDate,'01',Sum(SumAmt),0) RealDate01,Decode(RealDate,'02',Sum(SumAmt),0) RealDate02,Decode(RealDate,'03',Sum(SumAmt),0) RealDate03,Decode(RealDate,'04',Sum(SumAmt),0) RealDate04," & _
                "Decode(RealDate,'05',Sum(SumAmt),0) RealDate05,Decode(RealDate,'06',Sum(SumAmt),0) RealDate06,Decode(RealDate,'07',Sum(SumAmt),0) RealDate07,Decode(RealDate,'08',Sum(SumAmt),0) RealDate08," & _
                "Decode(RealDate,'09',Sum(SumAmt),0) RealDate09,Decode(RealDate,'10',Sum(SumAmt),0) RealDate10,Decode(RealDate,'11',Sum(SumAmt),0) RealDate11,Decode(RealDate,'12',Sum(SumAmt),0) RealDate12," & _
                "Decode(RealDate,'13',Sum(SumAmt),0) RealDate13,Decode(RealDate,'14',Sum(SumAmt),0) RealDate14,Decode(RealDate,'15',Sum(SumAmt),0) RealDate15,Decode(RealDate,'16',Sum(SumAmt),0) RealDate16," & _
                "Decode(RealDate,'17',Sum(SumAmt),0) RealDate17,Decode(RealDate,'18',Sum(SumAmt),0) RealDate18,Decode(RealDate,'19',Sum(SumAmt),0) RealDate19,Decode(RealDate,'20',Sum(SumAmt),0) RealDate20," & _
                "Decode(RealDate,'21',Sum(SumAmt),0) RealDate21,Decode(RealDate,'22',Sum(SumAmt),0) RealDate22,Decode(RealDate,'23',Sum(SumAmt),0) RealDate23,Decode(RealDate,'24',Sum(SumAmt),0) RealDate24," & _
                "Decode(RealDate,'25',Sum(SumAmt),0) RealDate25,Decode(RealDate,'26',Sum(SumAmt),0) RealDate26,Decode(RealDate,'27',Sum(SumAmt),0) RealDate27,Decode(RealDate,'28',Sum(SumAmt),0) RealDate28," & _
                "Decode(RealDate,'29',Sum(SumAmt),0) RealDate29,Decode(RealDate,'30',Sum(SumAmt),0) RealDate30,Decode(RealDate,'31',Sum(SumAmt),0) RealDate31,Sum(SumAmt) SumAmt From (" & _
                "Select B.Description as ServName,A.CitemName as CitemName,to_char(A.RealDate,'DD') as RealDate,Sum(A.RealAmt) as SumAmt " & _
                "FROM CD002 B,SO033 A Where A.ServCode=B.CodeNo " & strChoose & " GROUP BY B.Description,A.CitemName,to_char(A.RealDate,'DD') Union All " & _
                "Select B.Description as ServName,A.CitemName as CitemName,to_char(A.RealDate,'DD') as RealDate,Sum(A.RealAmt) as SumAmt " & _
                "FROM CD002 B,SO034 A Where A.ServCode=B.CodeNo " & strChoose & " GROUP BY B.Description,A.CitemName,to_char(A.RealDate,'DD')) A Group by ServName,CitemName,RealDate) A Group By ServName,CitemName"

       subInsertMDB1 (strSQL)
    Else
       strSQL = "Select ServName,CitemName,Sum(RealDate01) Date1,Sum(RealDate02) Date2,Sum(RealDate03) Date3," & _
                "Sum(RealDate04) Date4,Sum(RealDate05) Date5,Sum(RealDate06) Date6,Sum(RealDate07) Date7,Sum(RealDate08) Date8,Sum(RealDate09) Date9,Sum(RealDate10) Date10,Sum (RealDate11) Date11, Sum(RealDate12) Date12,Sum(SumAmt) SumAmt " & _
                "From (Select ServName,CitemName,Decode(RealDate,'01',Sum(SumAmt),0) RealDate01,Decode(RealDate,'02',Sum(SumAmt),0) RealDate02,Decode(RealDate,'03',Sum(SumAmt),0) RealDate03," & _
                "Decode(RealDate,'04',Sum(SumAmt),0) RealDate04,Decode(RealDate,'05',Sum(SumAmt),0) RealDate05,Decode(RealDate,'06',Sum(SumAmt),0) RealDate06,Decode(RealDate,'07',Sum(SumAmt),0) RealDate07," & _
                "Decode(RealDate,'08',Sum(SumAmt),0) RealDate08,Decode(RealDate,'09',Sum(SumAmt),0) RealDate09,Decode(RealDate,'10',Sum(SumAmt),0) RealDate10,Decode(RealDate,'11',Sum(SumAmt),0) RealDate11," & _
                "Decode(RealDate,'12',Sum(SumAmt),0) RealDate12,Sum(SumAmt) SumAmt From (" & _
                "Select B.Description as ServName,A.CitemName as CitemName,to_char(A.RealDate,'MM') as RealDate,Sum(A.RealAmt) as SumAmt FROM CD002 B,SO033 A " & _
                "Where A.ServCode = B.CodeNo " & strChoose & " GROUP BY B.Description,A.CitemName,to_char(A.RealDate,'MM') UNION All " & _
                "Select B.Description as ServName,A.CitemName as CitemName,to_char(A.RealDate,'MM') as RealDate,Sum(A.RealAmt) as SumAmt FROM CD002 B,SO034 A " & _
                "Where A.ServCode = B.CodeNo " & strChoose & " GROUP BY B.Description,A.CitemName,to_char(A.RealDate,'MM')) A Group by ServName,CitemName,RealDate) A Group By ServName,CitemName"

                
       subInsertMDB2 (strSQL)
    End If
    
    rsTmp.CursorLocation = adUseClient
    If rsTmp.State = 1 Then rsTmp.Close
    Set rsTmp = cnn.Execute("SELECT Count(*) as intCount FROM SO5220A1")
    'rsTmp.Open "SELECT Count(*) as intCount FROM SO5220A1", cnn, adOpenForwardOnly, adLockReadOnly
    If rsTmp("intCount") = 0 Then
         MsgNoRcd
         SendSQL , , True
    Else
        strSQL = "SELECT * FROM SO5220A1"
        If optInstant.Value Then
            If blnExcel Then
                Set rsExcel = cnn.Execute(strSQL)
                '#7074 To Add a choice-condition parameter into excel print function by Kin 2015/09/09
                Call UseProperty(rsExcel, "每月各項收費加總統計[當月金額]", "第一頁", _
                                 "服務區,收費項目,1日,2日,3日,4日,5日,6日,7日,8日,9日,10日,11日,12日,13日,14日,15日,16日,17日,18日,19日,20日,21日,22日,23日,24日,25日,26日,27日,28日,29日,30日,31日,合計", _
                                 , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , strChooseString)
                blnExcel = False
            Else
                Call PrintRpt(GetPrinterName(5), RptName("SO5220", 1), , "每月各項收費加總統計[當月金額]", strSQL, strChooseString, , True, "Tmp0000.Mdb", , , GiPaperLandscape)
            End If
        Else
            If blnExcel Then
                strSQL = "Select ServName,CitemName,Day1,Day2,Day3,Day4,Day5,Day6,Day7,Day8,Day9,Day10,Day11,Day12,SumAmt From SO5220A1"
                Set rsExcel = cnn.Execute(strSQL)
                Call UseProperty(rsExcel, "每月各項收費加總統計[年度總計]", "第一頁", _
                                 "服務區,收費項目,1月,2月,3月,4月,5月,6月,7月,8月,9月,10月,11月,12月,合計")
                blnExcel = False
            Else
                Call PrintRpt(GetPrinterName(5), RptName("SO5220", 2), , "每月各項收費加總統計[年度總計]", strSQL, strChooseString, , True, "Tmp0000.Mdb", , , GiPaperLandscape)
            End If
        End If
    End If
    CloseRecordset rsTmp
    CloseRecordset rsExcel
    blnExcel = False
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "subPrint")
End Sub

Private Sub subInsertMDB1(strSQL As String)
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    If rsTmp.State = 1 Then rsTmp.Close
    Set rsTmp = gcnGi.Execute(strSQL)
    SendSQL strSQL, True
    
     cnn.BeginTrans
      cnn.Execute "Delete From SO5220A1"
      While Not rsTmp.EOF
           cnn.Execute "Insert Into SO5220A1 (ServName,CitemName,Day1,Day2,Day3,Day4,Day5,Day6,Day7,Day8,Day9,Day10,Day11,Day12,Day13," & _
                       "Day14,Day15,Day16,Day17,Day18,Day19,Day20,Day21,Day22,Day23,Day24,Day25,Day26,Day27,Day28,Day29,Day30,Day31,SumAmt) Values ('" & _
                       rsTmp("ServName") & "','" & rsTmp("CitemName") & "'," & _
                       rsTmp("Date1") & "," & rsTmp("Date2") & "," & rsTmp("Date3") & "," & rsTmp("Date4") & "," & _
                       rsTmp("Date5") & "," & rsTmp("Date6") & "," & rsTmp("Date7") & "," & rsTmp("Date8") & "," & _
                       rsTmp("Date9") & "," & rsTmp("Date10") & "," & rsTmp("Date11") & "," & rsTmp("Date12") & "," & _
                       rsTmp("Date13") & "," & rsTmp("Date14") & "," & rsTmp("Date15") & "," & rsTmp("Date16") & "," & _
                       rsTmp("Date17") & "," & rsTmp("Date18") & "," & rsTmp("Date19") & "," & rsTmp("Date20") & "," & _
                       rsTmp("Date21") & "," & rsTmp("Date22") & "," & rsTmp("Date23") & "," & rsTmp("Date24") & "," & _
                       rsTmp("Date25") & "," & rsTmp("Date26") & "," & rsTmp("Date27") & "," & rsTmp("Date28") & "," & _
                       rsTmp("Date29") & "," & rsTmp("Date30") & "," & rsTmp("Date31") & "," & rsTmp("SumAmt") & ")"
           rsTmp.MoveNext
           DoEvents
      Wend
    cnn.CommitTrans
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "subInsertMDB1")
End Sub

Private Sub subInsertMDB2(strSQL As String)
  On Error GoTo ChkErr
  Dim rsTmp As New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    If rsTmp.State = 1 Then rsTmp.Close
    Set rsTmp = gcnGi.Execute(strSQL)
    SendSQL strSQL, True

     cnn.Execute "Delete From SO5220A1"
     cnn.BeginTrans
      While Not rsTmp.EOF
           cnn.Execute "Insert Into SO5220A1 (ServName,CitemName,Day1,Day2,Day3,Day4,Day5,Day6,Day7,Day8,Day9,Day10,Day11,Day12,SumAmt) Values ('" & _
                       rsTmp("ServName") & "','" & rsTmp("CitemName") & "'," & _
                       rsTmp("Date1") & "," & rsTmp("Date2") & "," & rsTmp("Date3") & "," & rsTmp("Date4") & "," & _
                       rsTmp("Date5") & "," & rsTmp("Date6") & "," & rsTmp("Date7") & "," & rsTmp("Date8") & "," & _
                       rsTmp("Date9") & "," & rsTmp("Date10") & "," & rsTmp("Date11") & "," & rsTmp("Date12") & "," & _
                       rsTmp("SumAmt") & ")"
           rsTmp.MoveNext
           DoEvents
      Wend
    cnn.CommitTrans
    CloseRecordset rsTmp
  Exit Sub
ChkErr:
  cnn.RollbackTrans
  Call ErrSub(Me.Name, "subInsertMDB2")
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    Call subGim
    Call subGil
    gilCompCode.SetCodeNo garyGi(9)
    gilCompCode.Query_Description
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub cmdExit_Click()
  On Error GoTo ChkErr
    Unload Me
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdExit_Click")
End Sub

Private Sub subGim()
  On Error GoTo ChkErr
    Call SetgiMulti(gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱")
    Call SetgiMulti(gimClassCode, "CodeNo", "Description", "CD004", "客戶類別代碼", "客戶類別名稱")
    Call SetgiMultiAddItem(gimBillType, "B,I,P,M,T", "收費單,裝機單,停拆移機單,維修單,臨時收費單")
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub SetgiMultiAddItem(objgiMulti As GiMulti, strCodeNo As String, strDescription As String)
  On Error GoTo ChkErr
    Dim varCodeNo As Variant
    Dim varDescription As Variant
    Dim varValue As Variant
    Dim intValue As Integer
    intValue = 0
    varCodeNo = Split(strCodeNo, ",")
    varDescription = Split(strDescription, ",")
    For Each varValue In varCodeNo
        objgiMulti.AddItem varValue, varDescription(intValue)
        intValue = intValue + 1
    Next
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "SetgiMultiAddItem")
End Sub

Private Sub subGil()
  On Error GoTo ChkErr
    'Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039")
    Call SetgiList(gilCompCode, "CodeNo", "Description", "CD039", , , , , , , GetCompCodeFilter(gilCompCode))
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub
  
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ChkErr
    Call FunctionKey(KeyCode, Shift, frmSO5220A)
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  ReleaseCOM frmSO5220A
End Sub

Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
    gilServiceType.ListIndex = 1
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Change")
End Sub

Private Sub gilCompCode_Validate(Cancel As Boolean)
  On Error GoTo ChkErr
    If gilCompCode.GetCodeNo = "" Then
        MsgMustBe ("公司別")
        Cancel = True
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Validate")
End Sub

Private Sub gilServiceType_Change()
  On Error GoTo ChkErr
    GiMultiFilter gimCMCode, gilServiceType.GetCodeNo
    GiMultiFilter gimClassCode, gilServiceType.GetCodeNo
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilServiceType_Change")
End Sub

Private Sub optInstant_Click()
  On Error GoTo ChkErr
    If optInstant.Value Then
        gyrRealDate.Visible = False
        Call gymRealDate.SetValue("")
        gymRealDate.Visible = True
        gymRealDate.SetFocus
        lblRealDate.Visible = True
        lblRealYear.Visible = False
    Else
        gymRealDate.Visible = False
        Call gyrRealDate.SetValue("")
        gyrRealDate.Visible = True
        gyrRealDate.SetFocus
        lblRealDate.Visible = False
        lblRealYear.Visible = True
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "optInstant_Click")
End Sub

Private Sub optReach_Click()
  On Error GoTo ChkErr
    If optInstant.Value Then
        gyrRealDate.Visible = False
        Call gymRealDate.SetValue("")
        gymRealDate.Visible = True
        gymRealDate.SetFocus
        lblRealDate.Visible = True
        lblRealYear.Visible = False
    Else
        gymRealDate.Visible = False
        Call gyrRealDate.SetValue("")
        gyrRealDate.Visible = True
        gyrRealDate.SetFocus
        lblRealDate.Visible = False
        lblRealYear.Visible = True
    End If
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "optReach_Click")
End Sub
