VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO3110A 
   BorderStyle     =   1  '單線固定
   Caption         =   "本期收費資料產生[SO3110A]"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3100A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdClearTmp 
      Caption         =   "清空所有暫存資料"
      Height          =   345
      Left            =   3660
      TabIndex        =   26
      Top             =   7080
      Width           =   2055
   End
   Begin VB.TextBox txtShowMessage 
      Height          =   2595
      Left            =   2430
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   38
      Top             =   2010
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "未產生一覽表"
      Height          =   345
      Left            =   1800
      TabIndex        =   25
      Top             =   7080
      Width           =   1515
   End
   Begin VB.CommandButton cmdPara 
      Caption         =   "上次執行參數"
      Height          =   345
      Left            =   120
      TabIndex        =   24
      Top             =   7080
      Width           =   1425
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2. 確定"
      Default         =   -1  'True
      Height          =   375
      Left            =   8550
      TabIndex        =   27
      Top             =   7050
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   10020
      TabIndex        =   28
      Top             =   7050
      Width           =   1365
   End
   Begin VB.Frame fraData 
      Height          =   6855
      Left            =   90
      TabIndex        =   31
      Top             =   120
      Width           =   11655
      Begin VB.ComboBox cboBillHeadFmt 
         Height          =   315
         Left            =   2250
         Style           =   2  '單純下拉式
         TabIndex        =   42
         Top             =   690
         Visible         =   0   'False
         Width           =   4995
      End
      Begin CS_Multi.CSmulti gimSuite 
         Height          =   405
         Left            =   180
         TabIndex        =   40
         Top             =   5160
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   714
         ButtonCaption   =   "套房名稱"
      End
      Begin VB.CheckBox chkStopPRCust 
         Caption         =   "&4.產生停機客戶於本期之收費資料"
         Height          =   405
         Left            =   7200
         TabIndex        =   22
         Top             =   5940
         Width           =   3225
      End
      Begin CS_Multi.CSmulti gimServiceType2 
         Height          =   285
         Left            =   9150
         TabIndex        =   39
         Top             =   480
         Visible         =   0   'False
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   503
      End
      Begin CS_Multi.CSmulti gimCitemCode 
         Height          =   375
         Left            =   180
         TabIndex        =   8
         Top             =   2070
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   661
         ButtonCaption   =   "收費項目"
         IsReadOnly      =   -1  'True
      End
      Begin Gi_Multi.GiMulti gimClassCode 
         Height          =   375
         Left            =   180
         TabIndex        =   9
         Top             =   2415
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   661
         ButtonCaption   =   "客戶類別"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin prjGiList.GiList gilCompCode 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
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
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilBankCode 
         Height          =   315
         Left            =   8130
         TabIndex        =   6
         Top             =   990
         Visible         =   0   'False
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
         F2Corresponding =   -1  'True
      End
      Begin VB.CheckBox chkATM 
         Caption         =   "是否產生虛擬帳號"
         Height          =   195
         Left            =   5250
         TabIndex        =   5
         Top             =   1050
         Visible         =   0   'False
         Width           =   1905
      End
      Begin Gi_Date.GiDate gdaShouldDate2 
         Height          =   345
         Left            =   3840
         TabIndex        =   4
         Top             =   960
         Width           =   1125
         _ExtentX        =   1984
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
      Begin Gi_Date.GiDate gdaShouldDate1 
         Height          =   345
         Left            =   2340
         TabIndex        =   3
         Top             =   960
         Width           =   1125
         _ExtentX        =   1984
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
      Begin Gi_Multi.GiMulti gimClctArea 
         Height          =   375
         Left            =   180
         TabIndex        =   12
         Top             =   3435
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   661
         ButtonCaption   =   "收  費  區"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Multi.GiMulti gimServCode 
         Height          =   375
         Left            =   180
         TabIndex        =   11
         Top             =   3105
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   661
         ButtonCaption   =   "服  務  區"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Multi.GiMulti gimAreaCode 
         Height          =   375
         Left            =   180
         TabIndex        =   10
         Top             =   2760
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   661
         ButtonCaption   =   "行  政  區"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin VB.CheckBox chkToEndDate 
         Caption         =   "&2.大樓個收戶收費至合約到期日"
         Height          =   405
         Left            =   7200
         TabIndex        =   20
         Top             =   5520
         Width           =   3225
      End
      Begin VB.CheckBox chkGenPRCust 
         Caption         =   "&3.產生待拆客戶於本期之收費資料"
         Height          =   195
         Left            =   3780
         TabIndex        =   21
         Top             =   6030
         Width           =   3225
      End
      Begin VB.CheckBox chkGenOverdue 
         Caption         =   "&1.產生欠費客戶於本期之收費資料"
         Height          =   405
         Left            =   3780
         TabIndex        =   19
         Top             =   5550
         Value           =   1  '核取
         Width           =   3225
      End
      Begin VB.Frame fraAddrType 
         Caption         =   "收費單上地址依據"
         Enabled         =   0   'False
         Height          =   735
         Left            =   240
         TabIndex        =   30
         Top             =   5550
         Width           =   3225
         Begin VB.OptionButton optInstAddress 
            Caption         =   "&B.裝機地址"
            Height          =   195
            Left            =   1710
            TabIndex        =   18
            Top             =   330
            Width           =   1485
         End
         Begin VB.OptionButton optChargeAddress 
            Caption         =   "&A.收費地址"
            Height          =   195
            Left            =   150
            TabIndex        =   17
            Top             =   330
            Value           =   -1  'True
            Width           =   1485
         End
      End
      Begin VB.TextBox txtCustId 
         Height          =   375
         Left            =   1980
         TabIndex        =   23
         Top             =   6330
         Width           =   9315
      End
      Begin Gi_YM.GiYM gymClctYM 
         Height          =   345
         Left            =   1260
         TabIndex        =   2
         Top             =   960
         Width           =   945
         _ExtentX        =   1667
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
      Begin prjGiList.GiList gilServiceType 
         Height          =   315
         Left            =   6300
         TabIndex        =   1
         Top             =   270
         Visible         =   0   'False
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   556
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         F2Corresponding =   -1  'True
      End
      Begin CS_Multi.CSmulti gimStrtCode 
         Height          =   375
         Left            =   180
         TabIndex        =   15
         Top             =   4440
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   661
         ButtonCaption   =   "街道名稱"
      End
      Begin CS_Multi.CSmulti gimMduId 
         Height          =   375
         Left            =   180
         TabIndex        =   14
         Top             =   4110
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   661
         ButtonCaption   =   "大樓名稱"
      End
      Begin Gi_Multi.GiMulti gimCMCode 
         Height          =   375
         Left            =   180
         TabIndex        =   13
         Top             =   3780
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   661
         ButtonCaption   =   "收費方式"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin CS_Multi.CSmulti gimClassCode2 
         Height          =   375
         Left            =   10110
         TabIndex        =   29
         Top             =   5700
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         ButtonCaption   =   "客戶類別"
      End
      Begin Gi_Multi.GiMulti gimServiceType 
         Height          =   375
         Left            =   180
         TabIndex        =   7
         Top             =   1350
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   661
         ButtonCaption   =   "服務類別"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Multi.GiMulti gmdPayType 
         Height          =   390
         Left            =   180
         TabIndex        =   16
         Top             =   4800
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   688
         ButtonCaption   =   "繳  付  類  別"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Multi.GiMulti gimBillHeadFmt 
         Height          =   375
         Left            =   180
         TabIndex        =   43
         Top             =   1710
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   661
         ButtonCaption   =   "多帳戶產生依據"
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "多帳戶產生依據設定"
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   780
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblServiceType 
         AutoSize        =   -1  'True
         Caption         =   "服務類別"
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5280
         TabIndex        =   37
         Top             =   330
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblCompCode 
         AutoSize        =   -1  'True
         Caption         =   "公司別"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   420
         TabIndex        =   36
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lblBankCode 
         AutoSize        =   -1  'True
         Caption         =   "銀行名稱"
         Height          =   195
         Left            =   7230
         TabIndex        =   35
         Top             =   1050
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblNote4 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號(以,相隔)"
         Height          =   195
         Left            =   300
         TabIndex        =   34
         Top             =   6390
         Width           =   1530
      End
      Begin VB.Label lblDay2 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   3570
         TabIndex        =   33
         Top             =   1050
         Width           =   195
      End
      Begin VB.Label lblYM 
         AutoSize        =   -1  'True
         Caption         =   "收集年月"
         Height          =   195
         Left            =   390
         TabIndex        =   32
         Top             =   1020
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmSO3110A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cnTmp As New ADODB.Connection
Private intCombineBill As Integer
Private blnPaynowFlag As Boolean
Private blnGetPara As Boolean
Private blnIsOK As Boolean
Public rsM As New ADODB.Recordset
Public rsD As New ADODB.Recordset
Private blnAutoExec As Boolean
Private strReturnMsg As String
Dim rsBillHeadFmt As New ADODB.Recordset
Public Property Let uAutoExec(vData As Boolean)
    blnAutoExec = vData
End Property

Public Property Let uGetPara(vData As Boolean)
    blnGetPara = vData
End Property

Public Property Get uIsOk() As Boolean
    uIsOk = blnIsOK
End Property

Public Property Get uReturnMsg() As String
    uReturnMsg = strReturnMsg
End Property

Private Sub cboBillHeadFmt1_Click()
On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
    Dim strCD019SQL As String
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
        '#7049 改用CD068A.CitemCode By Kin 2015/07/09
        strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                                    " And CD068.BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1 " & _
                                                    " And PeriodFlag = 1 And Nvl(NoReplaceFlag,0) = 0"
                                                    
                                                    'PeriodFlag = 1 And Nvl(NoReplaceFlag,0) = 0
'        If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        If Not GetRS(rsCD019, strCD019SQL) Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        gimCitemCode.Filter = gimCitemCode.Filter
        gimCitemCode.SetQueryCode strCitemCode
End Sub

Private Sub chkATM_Click()
    On Error Resume Next
        Exit Sub
        If chkATM = 1 Then
            lblBankCode.Visible = True
            gilBankCode.Visible = True
        Else
            lblBankCode.Visible = False
            gilBankCode.Visible = False
            gilBankCode.SetCodeNo ""
            gilBankCode.SetDescription ""
        End If
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
        Unload Me
End Sub

Private Function GetPara() As Boolean
    If Not BatchCharge.GetBatchChargePara(Me, rsM, rsD) Then Exit Function
    blnIsOK = True
    GetPara = True
    Unload Me
End Function

Private Function SetPara() As Boolean
    If Not BatchCharge.SetBatchChargePara(Me, rsM, rsD) Then Exit Function
    SetPara = True
End Function

Private Sub cmdClearTmp_Click()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    DeleteTmp True, True
    Screen.MousePointer = vbDefault
        'MsgBox "清除完成", vbInformation, gimsgPrompt
End Sub

Private Sub cmdOk_Click()
    On Error GoTo chkErr
    Dim strCompCode As String, strCitem As String
    Dim strCustCode As String, strAreaName As String
    Dim strServArea As String, strClctArea As String
    Dim strMduId As String, strStrtName As String
    Dim strAMduId As String
    Dim strCMCode As String
    Dim intAddrType As Integer, intGenOverdue As Integer
    Dim intGenPRCust As Integer, intToEndDate As Integer
    Dim strServiceType As String
    Dim strRetMsg As String
    Dim strClassDispStr As String
    Dim strPayKinds As String
        txtShowMessage.Visible = False
        If IsDataOk = False Then Exit Sub
        Screen.MousePointer = vbHourglass
        strPayKinds = ""
        strCitem = gimCitemCode.GetQueryCode & ""
        strCustCode = gimClassCode.GetQryStr & ""
        strAreaName = gimAreaCode.GetQryStr
        strServArea = gimServCode.GetQryStr
        strClctArea = gimClctArea.GetQryStr
        strMduId = gimMduId.GetQryStr
        If gimSuite.GetQueryCode & "" <> "" Then
            If gimSuite.GetQryStr & "" <> "" Then
                strAMduId = gimSuite.GetQryStr & ""
            Else
                strAMduId = " In (" & gimSuite.GetQueryCode & " ) "
            End If
        End If
        
        strStrtName = gimStrtCode.GetQryStr
        strCMCode = gimCMCode.GetQryStr
        '2001/06/15
        strCompCode = " In ( " & gilCompCode.GetCodeNo & " )"
        strServiceType = Replace(gimServiceType.GetQueryCode, "'", "")
        
        intAddrType = IIf(optChargeAddress.Value = True, 1, 2)
        intGenOverdue = IIf(chkGenOverdue.Value = 1, 1, 0)
        intGenPRCust = IIf(chkGenPRCust.Value = 1, 1, 0)
        intToEndDate = IIf(chkToEndDate.Value = 1, 1, 0)
        '#5683 增加繳付類別 By Kin 2010/08/20
        If (gmdPayType.GetQryStr <> "") And (InStr(1, gmdPayType.GetDispStr, "全選") <= 0) Then
            strPayKinds = gmdPayType.GetQryStr
        End If
        If Len(gimCitemCode.GetQueryCode & "") > 0 Then
            
        
        End If
        '呼叫取參數 2011/11/30 Jacky 6153 MinChen
        If blnGetPara Then
            Call GetPara
            Exit Sub
        End If
        
        Dim lngRetCode As Variant
        On Error GoTo lTran
        gcnGi.BeginTrans
        '#6779 是否重建由參數判斷 By Kin 2014/05/27
        Dim blnRebuildFlag  As Boolean
        blnRebuildFlag = Val(GetRsValue("Select Nvl(ClearChargeBef,0) From " & GetOwner & "SO041", gcnGi)) = 1
        If Not DeleteTmp(blnRebuildFlag, False) Then     '刪除暫存資料
            gcnGi.RollbackTrans
            Exit Sub
        End If
            
        '#3096 增加傳入參數:產生停機客戶於本期之收費資料 By Kin 2007/07/18
        '#5683 增加繳付類別 By Kin 2010/08/20
        lngRetCode = SF_ACCOUNTING(gcnGi, gymClctYM.GetValue, Val(gdaShouldDate1.GetValue), _
                    Val(gdaShouldDate2.GetValue), strCompCode, strCitem, strCustCode, _
                    strAreaName, strServArea, strClctArea, strMduId, Empty, _
                    strStrtName, strAMduId, Val(intAddrType), Val(intGenOverdue), Val(intGenPRCust), _
                    Val(intToEndDate), txtCustId.Text, garyGi(0), garyGi(1), _
                    strServiceType, gilBankCode.GetCodeNo, strCMCode, _
                    chkStopPRCust.Value, "", "", strPayKinds, strRetMsg)
        If lngRetCode < 0 Then
            gcnGi.RollbackTrans
            MsgBox strRetMsg, vbCritical, giMsgWarning
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
       ' MsgBox "將參數傳送到 SF_C\GenChargel", vbExclamation, "訊息！"
       '將參數及結果新增一筆資料至出單作業 Log 檔（SO059) :<2000/03/22>
       Debug.Print strRetMsg
       Dim strInstSql As String, strPara As String
       strInstSql = "Insert Into " & GetOwner & "SO059 (UpdTime,UpdName,CompCode,ServiceType,Para,Result,RetCode) Values ('" & GetDTString(RightNow) & "','" & garyGi(1) & "'," & gilCompCode.GetCodeNo & ",'" & strServiceType & "',"
       If optInstAddress Then
          strPara = "0;"
       Else
          strPara = "1;"
       End If
       strPara = strPara & chkGenOverdue.Value & ";" & chkToEndDate.Value & ";" & chkGenPRCust.Value & ";" & chkATM.Value
       strPara = strPara & "出單年月：" & gymClctYM.GetValue(True) & " 　日期：" & gdaShouldDate1.Text & "~" & gdaShouldDate2.Text & vbCrLf
       strPara = strPara & "服務類別 : " & strServiceType & vbCrLf
       strPara = strPara & "公司別 : " & gilCompCode.GetDescription & vbCrLf
       strPara = strPara & "收費項目：" & gimCitemCode.GetQueryCode & vbCrLf
       strClassDispStr = gimClassCode.GetQueryCode
       'If Right(strClassDispStr, 1) <> RightB(strClassDispStr, 1) Then
       '    strClassDispStr = Replace(Left(strClassDispStr, Len(strClassDispStr)), "-", "")
       'End If
       strPara = strPara & "客戶類別:" & strClassDispStr & vbCrLf
       strPara = strPara & "行政區：" & gimAreaCode.GetQueryCode & vbCrLf
       strPara = strPara & "服務區：" & gimServCode.GetQueryCode & vbCrLf
       strPara = strPara & "收費區：" & gimClctArea.GetQueryCode & vbCrLf
       strPara = strPara & "收費方式：" & gimCMCode.GetQueryCode & vbCrLf
       strPara = strPara & "大樓：" & gimMduId.GetQueryCode & vbCrLf
       strPara = strPara & "套房：" & gimSuite.GetQueryCode & vbCrLf
       strPara = strPara & "街道：" & gimStrtCode.GetQueryCode & vbCrLf
       strPara = strPara & "地址依據：" & IIf(optChargeAddress.Value = True, "收費地址", "裝機地址") & vbCrLf
       strPara = strPara & "是否產生虛擬帳號：" & IIf(chkATM.Value = 1, "是", "否") & vbCrLf
       strPara = strPara & "產生欠費客戶於本期之收費資料：" & IIf(chkGenOverdue.Value = 1, "是", "否") & vbCrLf
       strPara = strPara & "產生待拆客戶於本期之收費資料：" & IIf(chkToEndDate.Value = 1, "是", "否") & vbCrLf
       strPara = strPara & "大樓個收戶收費至合約到期日：" & IIf(chkGenPRCust.Value = 1, "是", "否") & vbCrLf
       strPara = Replace(strPara, "'", "")
       strInstSql = strInstSql & "'" & MidMbcs(strPara, 1, 4000) & "','" & strRetMsg & "'," & lngRetCode & ")"
       gcnGi.Execute (Replace(strInstSql, Chr(0), ""))
       gcnGi.CommitTrans
       On Error GoTo chkErr
        Screen.MousePointer = vbDefault
        If Not blnAutoExec Then
            MsgBox strRetMsg, vbInformation, "訊息！"
        Else
            blnIsOK = True
            strReturnMsg = strRetMsg
            Unload Me
        End If
        'Unload Me
    Exit Sub
lTran:
    txtShowMessage.Visible = True
    txtShowMessage = strInstSql
    gcnGi.RollbackTrans
chkErr:
    Screen.MousePointer = vbDefault
    Call ErrSub(Me.Name, "cmdOk_Click")
End Sub

Private Sub cmdPara_Click()
    On Error GoTo chkErr
Dim rsSo059log As New ADODB.Recordset
Dim strmsg As String
        rsSo059log.MaxRecords = 1
        If Not GetRS(rsSo059log, "Select UpdTime,UpdName,para,Result from " & GetOwner & "So059 Order by UpdTime Desc") Then Exit Sub
        If Not rsSo059log.EOF Then
            strmsg = "最後執行時間：" & rsSo059log("UpdTime").Value & vbCrLf
            strmsg = strmsg & "最後執行人員：" & rsSo059log("UpdName").Value & vbCrLf
            strmsg = strmsg & "執  行  參  數  ：" & vbCrLf & Mid(rsSo059log("Para").Value, InStr(rsSo059log("Para"), Chr(255)) + 1) & vbCrLf
            strmsg = strmsg & "執  行  結  果  ：" & vbCrLf & rsSo059log("Result").Value & vbCrLf
            MsgBox strmsg, vbInformation, "訊息！"
        Else
            MsgBox "無上次執行參數！", vbInformation, "上次執行參數！"
        End If
        Call CloseRecordset(rsSo059log)
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdPara_Click"
End Sub

Private Sub cmdPrint_Click()
  On Error GoTo chkErr
'    Dim rsTmp As New ADODB.Recordset
'        If Not GetRS(rsTmp, "Select * From " & GetOwner & "Tmp009") Then Exit Sub
'        If rsTmp.RecordCount > 0 Then
'            With ViewForm
'                .uRecordset = rsTmp
'                .uIsSo3110A = True
'                .Show vbModal
'            End With
'        End If
'        Call Close3Recordset(rsTmp)
        
        Screen.MousePointer = vbHourglass
        Dim rstmp1 As New ADODB.Recordset
        Dim rsTmp2 As New ADODB.Recordset
        Dim rsTmp3 As New ADODB.Recordset
        Dim rsTmp4 As New ADODB.Recordset
        Dim strQry1 As String
        Dim strQry2 As String
        Dim strQry3 As String
        Dim strQry4 As String
        '#7118 Add the field of MduId to tmp009 that is related by so001 by kin 2018/06/07
        '#8634 Add SO017.Name By Kin 2020/08/07
        strQry1 = "Select A.*,B.MduId,(Select Name From " & GetOwner & " SO017 Where MduId = B.MduId And CompCode = " & gCompCode & " ) MduName " & _
                " From " & GetOwner & "TMP009 A," & GetOwner & "SO001 B  Where A.Uccode Is Not Null " & _
                " And Length(A.BillNo)=LengthB(A.BillNo) And Nvl(A.STFlag,0)=0 And A.CustId = B.CustId "
        strQry2 = "Select A.*,B.MduId,(Select Name From " & GetOwner & " SO017 Where MduId = B.MduId And CompCode = " & gCompCode & " ) MduName " & _
                " From " & GetOwner & "TMP009 A," & GetOwner & "SO001 B Where A.Uccode Is Null " & _
                " And Length(A.BillNo)=LengthB(A.BillNo) And Nvl(A.STFlag,0)=0 And A.CustId = B.CustId "
        strQry3 = "Select A.*,B.MduId,(Select Name From " & GetOwner & " SO017 Where MduId = B.MduId And CompCode = " & gCompCode & " ) MduName " & _
                " From " & GetOwner & "TMP009 A, " & GetOwner & "SO001 B Where Nvl(A.STFlag,0)=1 And A.CustId = B.CustId "
        strQry4 = "Select A.*,B.MduId,(Select Name From " & GetOwner & " SO017 Where MduId = B.MduId And CompCode = " & gCompCode & " ) MduName " & _
                " From " & GetOwner & "TMP009 A," & GetOwner & "SO001 B " & _
                " Where Length(A.BillNo)<> LengthB(A.BillNo) And A.CustId = B.CustId  " & _
                " And Nvl(A.STFlag,0)=0"
        
        If Not GetRS(rstmp1, strQry1, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If Not GetRS(rsTmp2, strQry2, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If Not GetRS(rsTmp3, strQry3, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If Not GetRS(rsTmp4, strQry4, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If (rstmp1.EOF) And (rsTmp2.EOF) And (rsTmp3.EOF) And (rsTmp4.EOF) Then
            MsgBox "無未出單資料！", vbInformation, "訊息"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        With frmSO3110B
            Set .rs1 = rstmp1
            Set .rs2 = rsTmp2
            Set .rs3 = rsTmp3
            Set .rs4 = rsTmp4
            .Show vbModal, Me
        End With
        'frmSO3110B.Show vbModal
    On Error Resume Next
    Screen.MousePointer = vbDefault
    Call CloseRecordset(rstmp1)
    Call CloseRecordset(rsTmp2)
    Call CloseRecordset(rsTmp3)
    Call CloseRecordset(rsTmp4)
  Exit Sub
chkErr:
    Screen.MousePointer = vbDefault
  Call ErrSub(Me.Name, "cmdPrint_Click")
End Sub
Private Function InsMDBData(ByRef rs As ADODB.Recordset, ByVal strTable As String) As Boolean
  On Error GoTo chkErr
    Dim rsMdb As New ADODB.Recordset
    Dim i As Long
    If cnn.State = adStateClosed Then Set cnn = GetTmpMdbCn
    If rs.EOF Then InsMDBData = True: Exit Function
    If Not GetRS(rsMdb, "Select * From " & strTable & " Where 1=0", cnn, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Function
    cnn.BeginTrans
    Do While Not rs.EOF
        rsMdb.AddNew
        For i = 0 To rsMdb.Fields.Count - 1
            rsMdb.Fields(rsMdb.Fields(i).Name) = rs.Fields(rsMdb.Fields(i).Name)
        Next i
        rsMdb.Update
        rs.MoveNext
    Loop
    cnn.CommitTrans
  On Error Resume Next
  InsMDBData = True
  Call ErrSub(Me.Name, "InsMDBData")
  Exit Function
chkErr:
  Call ErrSub(Me.Name, "InsMDBData")
End Function
Private Sub Form_Activate()
    On Error GoTo chkErr
        SetPara
        Screen.MousePointer = vbDefault
        If blnAutoExec Then
            cmdOk.Value = True
        End If
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Activate")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo chkErr
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF2 Then Call cmdOk_Click: Exit Sub
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_KeyDown")
End Sub

Private Sub Form_Load()
    On Error GoTo chkErr
        Set cnTmp = gcnGi
        If Alfa2 = True Then
            Call GetGlobal
        End If
        blnPaynowFlag = GetPaynowFlag
        Call subGil
        Call subGim
        Call SetDefaultValue
        Call CboAddItem
        If Val(GetRsValue("Select FLOWID From " & GetOwner & "SO041 ", gcnGi) & "") = 1 Then
            gimServiceType.Enabled = False
            gimCitemCode.IsReadOnly = True
            gimBillHeadFmt.Enabled = True
        Else
            gimServiceType.Enabled = True
            gimCitemCode.IsReadOnly = False
            gimCitemCode.SelectAll
            gimBillHeadFmt.Enabled = False
        End If
        
        
    Exit Sub
chkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Function ChkMaxLengthOK() As Boolean
    On Error GoTo chkErr
    Dim strChar As String
    Dim intLength  As Integer
        ChkMaxLengthOK = True
        strChar = txtCustId.Text
        Do
            intLength = InStr(1, txtCustId.Text, ",", vbTextCompare)
            strChar = Mid(strChar, intLength + 1)
            If InStr(1, strChar, ",", vbTextCompare) = 0 Then
                'If Len(Trim(strChar)) >= 8 Then ChkMaxLengthOK = False
                Exit Do
            End If
        Loop
    Exit Function
chkErr:
    ErrSub Me.Name, "ChkMaxLenthOK"
End Function

Private Sub subGil()
    On Error Resume Next
        SetgiList gilBankCode, "CodeNo", "Description", "CD018", , , , , , , "Where CheckATM=1", True
        SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
        SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱", , True
End Sub

Private Sub subGim()
    On Error GoTo chkErr
    '設定各複選元件的的屬性
    '收費項目 93/05/27 不過濾停用
    Dim aSQL As String
    aSQL = "select rownum,billheadfmt from " & GetOwner & "cd068 where nvl(stopflag,0)=0"
    Dim rsHeadFmt As New ADODB.Recordset
    Dim aHeadCode As String
    Dim aHeadDescription As String
    If Not GetRS(rsHeadFmt, aSQL, gcnGi, adUseClient, adOpenKeyset) Then Exit Sub
    rsHeadFmt.MoveFirst
    Do While Not rsHeadFmt.EOF
        If Len(aHeadCode & "") = 0 Then
            aHeadCode = rsHeadFmt("RowNum")
        Else
            aHeadCode = aHeadCode & "," & rsHeadFmt("RowNum")
        End If
        If Len(aHeadDescription & "") = 0 Then
            aHeadDescription = rsHeadFmt("billheadfmt")
        Else
            aHeadDescription = aHeadDescription & "," & rsHeadFmt("billheadfmt")
        End If
        rsHeadFmt.MoveNext
    Loop
    Call SetgiMultiAddItem(gimBillHeadFmt, aHeadCode, aHeadDescription, "代碼", "名稱")
    SetgiMulti gimServiceType, "CodeNo", "Description", "CD046", "代碼", "名稱"
    'SetgiMulti gimCitemCode, "CodeNo", "Description", "CD019", "代碼", "名稱", " Where PeriodFlag = 1 And Nvl(NoReplaceFlag,0) = 0 "
    '客戶類別
    SetgiMulti gimClassCode, "CodeNo", "Description", "CD004", "代碼", "名稱", , True
    '行政區
    SetgiMulti gimAreaCode, "CodeNo", "Description", "CD001", "代碼", "名稱", , True
    '服務區
    SetgiMulti gimServCode, "CodeNo", "Description", "CD002", "代碼", "名稱", , True
    '收費區
    SetgiMulti gimClctArea, "CodeNo", "Description", "CD040", "代碼", "名稱", , True
    '大樓
    SetgiMulti gimMduId, "Mduid", "Name", "SO017", "大樓編號", "大樓名稱"
    '套房
    SetgiMulti gimSuite, "Mduid", "Name", "SO202", "套房編號", "套房名稱"
    '街道
    SetgiMulti gimStrtCode, "CodeNo", "Description", "CD017", "代碼", "名稱", , True
    '收費方式
    SetgiMulti gimCMCode, "CodeNo", "Description", "CD031", "收費方式代碼", "收費方式名稱", , True
    '#5683 增加繳付類別 By Kin 2010/08/12
    SetgiMulti gmdPayType, "CodeNo", "Description", "CD112", "代碼", "名稱"
    '#5925 繳費類別增加預設值 By Kin 2011/04/12
    If blnPaynowFlag Then
        With gmdPayType
            Select Case Val(GetRsValue("SELECT NVL(PayKindDefault,0) FROM " & GetOwner & "SO041", gcnGi) & "")
                Case 0
                    .SelectAll
                Case 1
                    .SetQueryCode "0"
                Case 2
                    .SetQueryCode "1"
            End Select
        End With
    Else
        gmdPayType.Clear
        gmdPayType.Enabled = False
    End If
On Error Resume Next
    Close3Recordset rsHeadFmt
Exit Sub
chkErr:
    Call ErrSub(Me.Name, "subGim")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        Call ReleaseCOM(Me)
End Sub

Private Sub gdaShouldDate2_GotFocus()
    On Error Resume Next
        If gdaShouldDate2.GetValue = "" Then gdaShouldDate2.SetValue GetLastDate(gdaShouldDate1.GetValue(True))
End Sub
Private Sub CboAddItem()
    On Error Resume Next
        'If cboBillHeadFmt.ListCount > 0 Then Exit Sub
        If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr From " & GetOwner & "CD068 Where Nvl(Stopflag,0) = 0") Then Exit Sub
        Do While Not rsBillHeadFmt.EOF
            cboBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
            rsBillHeadFmt.MoveNext
        Loop
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
        If Not ChgComp(gilCompCode, "SO3100", "SO3110") Then Exit Sub
        Call subGil
        Call subGim
        'Call SelectServType(gilCompCode.GetCodeNo, gimServiceType)
        gimServiceType.SelectAll
        Call GiListFilter(gilBankCode, , gilCompCode.GetCodeNo)
        
'        gimBankId.Filter = gimBankId.Filter & IIf(gimBankId.Filter = "", " Where ", " And ") & " Upper(PrgName) = 'BANKCENTER'"
        Call GiMultiFilter(gimAreaCode, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimMduId, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimServCode, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimClctArea, , gilCompCode.GetCodeNo)
        Call GiMultiFilter(gimStrtCode, , gilCompCode.GetCodeNo)
        Call SetDefaultValue
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
End Sub

Private Sub gimBillHeadFmt_Change()
  On Error GoTo chkErr
    Screen.MousePointer = vbHourglass
    If Len(gimBillHeadFmt.GetQueryCode & "") = 0 Then
        gimCitemCode.Clear
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Dim rsCD019  As New ADODB.Recordset
    Dim strCD019SQL As String
    Dim strCitemCode As String
    If Len(gimBillHeadFmt.GetDescStr & "") = 0 Then
                strCD019SQL = "Select distinct CodeNo From " & GetOwner & "CD019 " & _
                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                      " And 1= 1)) And StopFlag <>1 " & _
                                                    " And PeriodFlag = 1 And Nvl(NoReplaceFlag,0) = 0"
    Else
        Dim aBillHeadFmt As String
        Dim aBillHeadFmtAry As Variant
        aBillHeadFmtAry = Split(gimBillHeadFmt.GetDispStr, ",")
        Dim i As Integer
        For i = LBound(aBillHeadFmtAry) To UBound(aBillHeadFmtAry)
            If Len(aBillHeadFmt & "") = 0 Then
                aBillHeadFmt = "'" & aBillHeadFmtAry(i) & "'"
            Else
                aBillHeadFmt = aBillHeadFmt & ",'" & aBillHeadFmtAry(i) & "'"
            End If
        Next
         strCD019SQL = "Select distinct CodeNo From " & GetOwner & "CD019 " & _
                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                      " And CD068.BillHeadFmt IN (" & aBillHeadFmt & ") )) And StopFlag <>1 " & _
                                                    " And PeriodFlag = 1 And Nvl(NoReplaceFlag,0) = 0"

    End If
        If Not GetRS(rsCD019, strCD019SQL) Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        gimCitemCode.Filter = gimCitemCode.Filter
        gimCitemCode.SetQueryCode strCitemCode
On Error Resume Next
  Close3Recordset rsCD019
  Screen.MousePointer = vbDefault
  Exit Sub
chkErr:
Screen.MousePointer = vbDefault
Call ErrSub(Me.Name, "gimBillHeadFmt_Change")

    'if gimBillHeadFmt.GetDispStr
    
'        strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
'                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
'                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
'                                                    " And CD068.BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1 " & _
'                                                    " And PeriodFlag = 1 And Nvl(NoReplaceFlag,0) = 0"
'
'        If Not GetRS(rsCD019, strCD019SQL) Then Exit Sub
'        If rsCD019.EOF Then Exit Sub
'        strCitemCode = rsCD019.GetString(, , , ",")
'        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
'        gimCitemCode.Filter = gimCitemCode.Filter
'        gimCitemCode.SetQueryCode strCitemCode
End Sub

Private Sub gimServiceType_Change()
    On Error Resume Next
        If gimServiceType.GetQueryCode = "" Then
            gimCitemCode.Clear
            Exit Sub
        End If
        gimCitemCode.Filter = "Where PeriodFlag = 1 And Nvl(NoReplaceFlag,0) = 0 And  (ServiceType In (" & gimServiceType.GetQueryCode & ") OR ServiceType is null)"
        gimClassCode.Filter = "Where (ServiceType In (" & gimServiceType.GetQueryCode & ") OR ServiceType is null)"
        gimCMCode.Filter = "Where (ServiceType In (" & gimServiceType.GetQueryCode & ") OR ServiceType is null)"
'        gimCitemCode.SelectAll
End Sub

Private Sub gymClctYM_Validate(Cancel As Boolean)
    On Error Resume Next
'        If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetValue gymClctYM.GetValue(True) & "/" & "01"
    If gdaShouldDate2.GetValue = "" Then
        gdaShouldDate2.SetValue GetLastDate(gymClctYM.GetValue(True))
    End If
End Sub

Private Sub txtCustId_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    '客戶編號只可輸入數值,逗號和減號
    If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 44 Or KeyAscii = 8 Or KeyAscii = 45 Then
        If KeyAscii = 44 Then
            If Asc(Right(" " & txtCustId.Text, 1)) = 44 Or txtCustId = "" Then KeyAscii = 9
        End If
        If Not ChkMaxLengthOK Then
           If Not (KeyAscii = 44 Or KeyAscii = 8) Then KeyAscii = 9
        End If
    Else
        KeyAscii = 9
    End If
End Sub

'設定預設值
Private Sub SetDefaultValue()
    On Error Resume Next
    Dim rsTmp As New ADODB.Recordset
    Dim lngPara14 As Long, lngPara2 As Long
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        
        gymClctYM.SetValue Format(RightDate, "YYYYMM")
        If Not GetRS(rsTmp, "Select para from " & GetOwner & "So059 Order by UpdTime Desc", gcnGi, , , , , , , 1) Then Exit Sub
        gdaShouldDate1.SetValue gymClctYM.GetValue(True) & "/" & "01"
        gdaShouldDate2.SetValue GetLastDate(gymClctYM.GetValue)
'        txtDay1 = 1
'        txtDay2 = GetLastDay(Year(rightnow), Month(rightnow))
        If rsTmp.EOF Then Exit Sub
        If Left(rsTmp("Para"), 1) = "0" Then
            optInstAddress = True
        Else
            optChargeAddress = True
        End If
        If Mid(rsTmp("Para"), 3, 1) = "1" Then chkGenOverdue.Value = 1
        If Mid(rsTmp("Para"), 5, 1) = "1" Then chkToEndDate.Value = 1
        If Mid(rsTmp("Para"), 7, 1) = "1" Then chkGenPRCust.Value = 1
        'If Mid(rsTmp("Para"), 9, 1) = "1" Then chkATM.Value = 1
        If Not GetSystemPara(rsTmp, gilCompCode.GetCodeNo, "", "SO043", "Para2,Para14") Then Exit Sub
        lngPara14 = Val(rsTmp("Para14") & "")
        lngPara2 = Val(rsTmp("Para2") & "")
        intCombineBill = Val(GetSystemParaItem("CombineBill", gilCompCode.GetCodeNo, "", "SO041", , True) & "")
        'cmdClearTmp.Visible = (intCombineBill = 1)
        
        optChargeAddress.Value = IIf(lngPara14 = 1, True, False)
        optInstAddress.Value = Not optChargeAddress.Value
        If lngPara2 = 0 Then chkGenOverdue.Value = 0: chkGenOverdue.Enabled = False
'        gimCitemCode.SelectAll
        '#3096 判斷SO043.BillStop來決定chkStopPRCust是否可使用 By Kin 2007/07/18
        If gcnGi.Execute("Select nvl(BillStop,0) From " & GetOwner & "SO043 Where CompCode=" & gilCompCode.GetCodeNo)(0) = 0 Then
            chkStopPRCust.Enabled = True
        Else
            chkStopPRCust.Enabled = False
        End If
        'gilBankCode.ListIndex = 1
     Exit Sub
chkErr:
    Call ErrSub(Me.Name, "SetDefaultValue")
End Sub

'必要欄位檢查
Private Function IsDataOk() As Boolean
'必要欄位：出單年月，起始日，截止日（截止日需大於起始日）
On Error GoTo chkErr
    IsDataOk = False
    If Not ChkDTok Then Exit Function
    
    If Not MustExist(gymClctYM, 1, "出單年月") Then Exit Function
    If Not MustExist(gdaShouldDate1, 1, "收集年月起始日") Then Exit Function
    If Not MustExist(gdaShouldDate2, 1, "收集年月截止日") Then Exit Function
    If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
    
    If Val(gdaShouldDate1.Text) > Val(gdaShouldDate2.Text) Then
           Call MsgDateRangeX("收集年月")
           gdaShouldDate1.SetFocus
           Exit Function
    End If
    If gimCitemCode.GetQueryCode = "" Then Call MsgMustBe("收費項目"): Exit Function
    If gimServiceType.Enabled Then
        If gimServiceType.GetQueryCode = "" Then Call MsgMustBe("服務類別"): Exit Function
    End If
    'If chkATM.Value = 1 And gilBankCode.GetCodeNo = "" Then
    '    MsgBox "如產生虛擬帳號,則銀行代碼須有值！", vbExclamation, "訊息！"
    '    gilBankCode.SetFocus
    '    Exit Function
    'End If
    
    IsDataOk = True
Exit Function
chkErr:
    Call ErrSub(Me.Name, "IsDataOk")
End Function

Private Function DeleteTmp(Optional blnClearTmp As Boolean = False, Optional ShowResult As Boolean = True) As Boolean
    On Error GoTo chkErr
    '#6636 改用呼叫後端方式重建 By Kin 2013/11/20
'        If blnClearTmp Then gcnGi.Execute "Truncate Table So032"
'        gcnGi.Execute "Truncate Table " & GetOwner & "So065"
'        gcnGi.Execute "Truncate Table " & GetOwner & "So066"
'        gcnGi.Execute "Truncate Table " & GetOwner & "So067"
'        gcnGi.Execute "Truncate Table " & GetOwner & "So068"
    
    Dim p_RetMsg As String
    Dim ComSF_RebuildTmpCharge As New ADODB.Command
    With ComSF_RebuildTmpCharge
       .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
       .Parameters.Append .CreateParameter("P_COMPCODE", adVarNumeric, adParamInput, , gCompCode)
       .Parameters.Append .CreateParameter("p_REBUILDFLAG", adVarNumeric, adParamInput, , IIf(blnClearTmp, 1, 0))
       .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, p_RetMsg)
         Set .ActiveConnection = gcnGi
        .CommandText = "SF_RebuildTmpCharge"
        .CommandType = adCmdStoredProc
        .Execute
        p_RetMsg = .Parameters("P_RETMSG").Value & ""
        If Val(.Parameters("Return_value").Value & "") < 0 Then
            MsgBox p_RetMsg, vbCritical, "錯誤"
            DeleteTmp = False
            Exit Function
        Else
            If ShowResult Then
                MsgBox "清除完成", vbInformation, gimsgPrompt
            End If
           DeleteTmp = True
        End If
    End With
    Exit Function
chkErr:
    ErrSub Me.Name, "DeleteTmp"
End Function


Public Function SF_ACCOUNTING(ByVal cnConn As ADODB.Connection, ByVal p_YM As String, ByVal P_DAY1 As Long, _
ByVal P_DAY2 As Long, ByVal p_CompSQL As String, ByVal p_CitemSQL As String, _
ByVal p_ClassSQL As String, ByVal P_AREASQL As String, ByVal p_ServSQL As String, _
ByVal p_ClctAreaSQL As String, ByVal P_MDUSQL As String, ByVal p_ProSQL As String, ByVal p_StrtSQL As String, ByVal p_AMduSQL As String, _
ByVal P_ADDRTYPE As Long, ByVal p_GENOVERDUE As Long, ByVal p_GENPRCUST As Long, _
ByVal P_TOENDDATE As Long, ByVal P_CUSTIDLIST As String, ByVal P_UPDEN As String, _
ByVal P_UPDNAME As String, ByVal p_SERVICETYPE As String, ByVal P_BANKCODE As String, _
ByVal p_CMSQL As String, ByVal p_StopPRCust As Long, ByVal p_FileName As String, _
ByVal p_FaciSeqNoSQL As String, ByVal p_PayKind As String, ByRef p_RetMsg As String) As Integer
On Error GoTo chkErr
        Dim ComSF_ACCOUNTING As New ADODB.Command
        If cnConn Is Nothing Then
                'MsgBox vmsgSendConn, vbCritical, vmsgPrompt
                Exit Function
        End If

        With ComSF_ACCOUNTING
                .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
                .Parameters.Append .CreateParameter("P_YM", adVarChar, adParamInput, 2000, p_YM)
                .Parameters.Append .CreateParameter("P_DAY1", adVarNumeric, adParamInput, , P_DAY1)
                .Parameters.Append .CreateParameter("P_DAY2", adVarNumeric, adParamInput, , P_DAY2)
                .Parameters.Append .CreateParameter("P_COMPSQL", adVarChar, adParamInput, 2000, p_CompSQL)
                .Parameters.Append .CreateParameter("P_CITEMSQL", adVarChar, adParamInput, 10000, p_CitemSQL)
                .Parameters.Append .CreateParameter("P_CLASSSQL", adVarChar, adParamInput, 2000, p_ClassSQL)
                .Parameters.Append .CreateParameter("P_AREASQL", adVarChar, adParamInput, 2000, P_AREASQL)
                .Parameters.Append .CreateParameter("P_SERVSQL", adVarChar, adParamInput, 2000, p_ServSQL)
                .Parameters.Append .CreateParameter("P_CLCTAREASQL", adVarChar, adParamInput, 2000, p_ClctAreaSQL)
                .Parameters.Append .CreateParameter("P_MDUSQL", adVarChar, adParamInput, 2000, P_MDUSQL)
                .Parameters.Append .CreateParameter("P_STRTSQL", adVarChar, adParamInput, 2000, p_StrtSQL)
                .Parameters.Append .CreateParameter("P_AMDUSQL", adVarChar, adParamInput, 2000, p_AMduSQL)
                .Parameters.Append .CreateParameter("p_ProSQL", adVarChar, adParamInput, 2000, p_ProSQL)
                .Parameters.Append .CreateParameter("P_ADDRTYPE", adVarNumeric, adParamInput, , P_ADDRTYPE)
                .Parameters.Append .CreateParameter("P_GENOVERDUE", adVarNumeric, adParamInput, , p_GENOVERDUE)
                .Parameters.Append .CreateParameter("P_GENPRCUST", adVarNumeric, adParamInput, , p_GENPRCUST)
                .Parameters.Append .CreateParameter("P_TOENDDATE", adVarNumeric, adParamInput, , P_TOENDDATE)
                .Parameters.Append .CreateParameter("P_CUSTIDLIST", adVarChar, adParamInput, 2000, P_CUSTIDLIST)
                .Parameters.Append .CreateParameter("P_UPDEN", adVarChar, adParamInput, 2000, P_UPDEN)
                .Parameters.Append .CreateParameter("P_UPDNAME", adVarChar, adParamInput, 2000, P_UPDNAME)
                .Parameters.Append .CreateParameter("P_SERVICETYPE", adVarChar, adParamInput, 2000, p_SERVICETYPE)
                .Parameters.Append .CreateParameter("P_BANKCODE", adVarChar, adParamInput, 2000, P_BANKCODE)
                .Parameters.Append .CreateParameter("P_CMSQL", adVarChar, adParamInput, 2000, p_CMSQL)
                .Parameters.Append .CreateParameter("P_STOPPRCUST", adNumeric, adParamInput, , p_StopPRCust)
                '#4030 為了訂單增加一個Tmp檔的暫存檔,在此功能傳入空字即可 By Kin 2008/08/12
                .Parameters.Append .CreateParameter("p_FileName", adVarChar, adParamInput, 2000, p_FileName)
                '#5569 增加設備序號參數,整批出帳傳空值就好 By Kin 2010/03/03
                .Parameters.Append .CreateParameter("p_FaciSeqNoSQL", adVarChar, adParamInput, 2000, p_FaciSeqNoSQL)
                .Parameters.Append .CreateParameter("p_PayKindSql", adVarChar, adParamInput, 2000, p_PayKind)
                .Parameters.Append .CreateParameter("p_NightRun", adVarNumeric, adParamInput, , 0)
                .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, p_RetMsg)
                Set .ActiveConnection = cnConn
                .CommandText = "SF_ACCOUNTING"
                .CommandType = adCmdStoredProc
                .Execute
                p_RetMsg = .Parameters("P_RETMSG").Value & ""
                SF_ACCOUNTING = Val(.Parameters("Return_value").Value & "")
        End With
        Exit Function
chkErr:
        ErrSub "clsStoreFunction", "SF_ACCOUNTING"
End Function
