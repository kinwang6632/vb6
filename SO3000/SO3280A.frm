VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{0CE6E756-92AE-11D3-8311-0080C8453DDF}#5.0#0"; "GiYM.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Object = "{6DDA0569-6454-11D3-9612-0048544B09DF}#4.1#0"; "GiTime.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.1#0"; "csMulti.ocx"
Begin VB.Form frmSO3280A 
   BorderStyle     =   1  '單線固定
   Caption         =   "收費單號調整 [SO3280A]"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SO3280A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame gmdPayType 
      Caption         =   "共用條件"
      ForeColor       =   &H00FF0000&
      Height          =   675
      Left            =   150
      TabIndex        =   44
      Top             =   6120
      Width           =   7065
      Begin Gi_Multi.GiMulti gimPayType 
         Height          =   375
         Left            =   90
         TabIndex        =   18
         Top             =   240
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   661
         ButtonCaption   =   "繳  付  類  別"
         DataType        =   2
         FontSize        =   9
         FontName        =   "新細明體"
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "單據合併調整項目"
      ForeColor       =   &H00FF0000&
      Height          =   885
      Left            =   150
      TabIndex        =   42
      Top             =   6810
      Width           =   7065
      Begin VB.CheckBox chkUpdPrtSNo 
         Caption         =   "調整印單序號"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3900
         TabIndex        =   51
         Top             =   540
         Width           =   1905
      End
      Begin VB.CheckBox chkUpdBillCloseDate 
         Caption         =   "調整繳費截止日期"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1920
         TabIndex        =   50
         Top             =   540
         Width           =   1905
      End
      Begin VB.CheckBox chkUpdUCcode 
         Caption         =   "調整未收原因"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   150
         TabIndex        =   49
         Top             =   540
         Width           =   1485
      End
      Begin VB.CheckBox chkUpdCreateEn 
         Caption         =   "產單人員"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5580
         TabIndex        =   22
         Top             =   240
         Width           =   1545
      End
      Begin VB.CheckBox chkCmCode 
         Caption         =   "收費方式"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3570
         TabIndex        =   21
         Top             =   240
         Width           =   1185
      End
      Begin VB.CheckBox chkCreateTime 
         Caption         =   "產生日期"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1920
         TabIndex        =   20
         Top             =   240
         Width           =   1185
      End
      Begin VB.CheckBox chkShouldDate 
         Caption         =   "應收日期"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   150
         TabIndex        =   19
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdPrevRpt 
      Caption         =   "上次統計結果"
      Height          =   375
      Left            =   2130
      TabIndex        =   24
      Top             =   7770
      Width           =   1365
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "F2. 確定"
      Default         =   -1  'True
      Height          =   375
      Left            =   300
      TabIndex        =   23
      Top             =   7800
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&X)"
      Height          =   375
      Left            =   5850
      TabIndex        =   25
      Top             =   7770
      Width           =   1365
   End
   Begin VB.Frame fraData 
      Height          =   6045
      Left            =   120
      TabIndex        =   27
      ToolTipText     =   "本功能只針對未收之T單做合併"
      Top             =   30
      Width           =   7155
      Begin VB.ComboBox cboTBillHeadFmt 
         Height          =   315
         Left            =   2040
         Style           =   2  '單純下拉式
         TabIndex        =   47
         Top             =   3720
         Width           =   4995
      End
      Begin CS_Multi.CSmulti gimTCitemCode 
         Height          =   375
         Left            =   180
         TabIndex        =   46
         Top             =   4500
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   661
         ButtonCaption   =   "收  費  項  目"
         IsReadOnly      =   -1  'True
      End
      Begin CS_Multi.CSmulti gimBCitemCode 
         Height          =   375
         Left            =   180
         TabIndex        =   45
         Top             =   2640
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   661
         ButtonCaption   =   "B單收費項目"
         IsReadOnly      =   -1  'True
      End
      Begin VB.CheckBox chkShouldDateMerge 
         Caption         =   "相同應收日期才併單"
         Enabled         =   0   'False
         Height          =   225
         Left            =   180
         TabIndex        =   4
         Top             =   1170
         Width           =   2235
      End
      Begin VB.TextBox txtCustId 
         Height          =   315
         Left            =   4380
         MaxLength       =   200
         TabIndex        =   5
         Top             =   1500
         Width           =   2610
      End
      Begin Gi_Time.GiTime gdtCreateTime2 
         Height          =   345
         Left            =   3270
         TabIndex        =   11
         Top             =   2250
         Width           =   1845
         _ExtentX        =   3254
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
      Begin Gi_Time.GiTime gdtCreateTime1 
         Height          =   315
         Left            =   1050
         TabIndex        =   10
         Top             =   2280
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
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
      Begin CS_Multi.CSmulti gimBCitemCode1 
         Height          =   345
         Left            =   6150
         TabIndex        =   26
         Top             =   5640
         Visible         =   0   'False
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         ButtonCaption   =   "B單收費項目"
      End
      Begin VB.ComboBox cboBillHeadFmt 
         Height          =   315
         Left            =   1980
         Style           =   2  '單純下拉式
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   4995
      End
      Begin VB.CheckBox chkCombine 
         Caption         =   "對應不到收費單的資料是否強制轉換"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   5700
         Value           =   1  '核取
         Width           =   3495
      End
      Begin Gi_Date.GiDate gdaProcessDate 
         Height          =   315
         Left            =   3285
         TabIndex        =   3
         Top             =   1095
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
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
         Left            =   1065
         TabIndex        =   1
         Top             =   660
         Width           =   3570
         _ExtentX        =   6297
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
         FldWidth2       =   2500
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilCompCode 
         Height          =   315
         Left            =   1065
         TabIndex        =   0
         Top             =   240
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
      Begin Gi_YM.GiYM gymClctYM1 
         Height          =   330
         Left            =   1065
         TabIndex        =   6
         Top             =   1905
         Width           =   900
         _ExtentX        =   1588
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
      Begin Gi_YM.GiYM gymClctYM2 
         Height          =   330
         Left            =   2385
         TabIndex        =   7
         Top             =   1905
         Width           =   900
         _ExtentX        =   1588
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
      Begin CS_Multi.CSmulti gimBUCCode 
         Height          =   345
         Left            =   180
         TabIndex        =   12
         Top             =   3000
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   609
         ButtonCaption   =   "未 收 費 原 因"
      End
      Begin Gi_Date.GiDate gdaShouldDate2 
         Height          =   345
         Left            =   2730
         TabIndex        =   14
         Top             =   4140
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
         Left            =   1140
         TabIndex        =   13
         Top             =   4140
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
      Begin CS_Multi.CSmulti gimTUCCode 
         Height          =   345
         Left            =   180
         TabIndex        =   15
         Top             =   4890
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   609
         ButtonCaption   =   "未 收 費 原 因"
      End
      Begin Gi_Multi.GiMulti gimBillType 
         Height          =   375
         Left            =   180
         TabIndex        =   16
         Top             =   5280
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   661
         ButtonCaption   =   "單  據  類  別"
         DataType        =   2
         DIY             =   -1  'True
         FontSize        =   9
         FontName        =   "新細明體"
      End
      Begin Gi_Date.GiDate gdaOweDate1 
         Height          =   345
         Left            =   4380
         TabIndex        =   8
         Top             =   1890
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
      Begin Gi_Date.GiDate gdaOweDate2 
         Height          =   345
         Left            =   5880
         TabIndex        =   9
         Top             =   1890
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "多帳戶產生依據設定"
         Height          =   195
         Left            =   210
         TabIndex        =   48
         Top             =   3780
         Width           =   1755
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號(以,相隔或以-取範圍)"
         Height          =   195
         Left            =   1770
         TabIndex        =   43
         Top             =   1590
         Width           =   2565
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   3000
         TabIndex        =   41
         Top             =   2340
         Width           =   195
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "產生時間"
         Height          =   195
         Left            =   180
         TabIndex        =   40
         Top             =   2340
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   5610
         TabIndex        =   39
         Top             =   1950
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "應收日期"
         Height          =   195
         Left            =   3510
         TabIndex        =   38
         Top             =   1980
         Width           =   780
      End
      Begin VB.Label lblBillHeadFmt 
         AutoSize        =   -1  'True
         Caption         =   "多帳戶產生依據設定"
         Height          =   195
         Left            =   150
         TabIndex        =   37
         Top             =   780
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblShouldDate 
         AutoSize        =   -1  'True
         Caption         =   "應收日期"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   4215
         Width           =   780
      End
      Begin VB.Label lblunti2 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2400
         TabIndex        =   35
         Top             =   4215
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "被合併單據條件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   34
         Top             =   3420
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "合併單據條件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   33
         Top             =   1560
         Width           =   1440
      End
      Begin VB.Label lblServiceType 
         AutoSize        =   -1  'True
         Caption         =   "服務類別"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   720
         Width           =   780
      End
      Begin VB.Label lblCompCode 
         AutoSize        =   -1  'True
         Caption         =   "公司別"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   330
         Width           =   585
      End
      Begin VB.Label lblNote2B 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   195
         Left            =   2055
         TabIndex        =   30
         Top             =   1965
         Width           =   195
      End
      Begin VB.Label lblNote2A 
         AutoSize        =   -1  'True
         Caption         =   "收集年月"
         Height          =   195
         Left            =   180
         TabIndex        =   29
         Top             =   1980
         Width           =   780
      End
      Begin VB.Label lblProcessDate 
         AutoSize        =   -1  'True
         Caption         =   "處理日期"
         Height          =   195
         Left            =   2490
         TabIndex        =   28
         Top             =   1170
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmSO3280A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmer:  Jacky Chen
' Last Modify: 2003/6/3
Option Explicit
Dim intPara24 As Integer
Dim rsBillHeadFmt As New ADODB.Recordset
Dim rsTBillHeadFmt As New ADODB.Recordset
Dim strPBillType As String   '是否啟用跨服務
Dim strPCitemStr As String  '非週期性項目與P服務
Dim blnIFlag As Boolean     '啟動跨服務後，收費項目裡是否有包含I服務
Dim strCustIdStr As String
Private blnGetPara As Boolean
Private blnIsOK As Boolean
Public rsM As New ADODB.Recordset
Public rsD As New ADODB.Recordset
Private blnAutoExec As Boolean
Private strReturnMsg As String

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

'Private blnShouldDateMerge As Boolean
Private Sub cboBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
    Dim rsCD019Mutli As New ADODB.Recordset
    Dim strCD019SQL As String
        If rsBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsBillHeadFmt.AbsolutePosition = cboBillHeadFmt.ListIndex + 1
        cboTBillHeadFmt.Text = cboBillHeadFmt.Text
        rsTBillHeadFmt.AbsolutePosition = cboTBillHeadFmt.ListIndex + 1
        
        '#7049 改用CD068A.CitemCode By Kin 2015/07/14
        strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                                    " And CD068.BillHeadFmt = '" & rsBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1"
'        If Not GetRS(rsCD019, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & rsBillHeadFmt("CitemCodeStr") & ") And StopFlag=0") Then Exit Sub
        If Not GetRS(rsCD019, strCD019SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        '***************************************************************************************
'        gimBCitemCode.Filter = "Where PeriodFlag = 1"
'        gimTCitemCode.Filter = ""
'        If strCitemCode <> "" Then
'            gimBCitemCode.Filter = "Where CodeNo In (" & strCitemCode & ") And PeriodFlag = 1"
'            gimTCitemCode.Filter = "Where CodeNo In (" & strCitemCode & ")"
'        End If
        '****************************************************************************************
        '#6801 T單收費項目用Filter會有一些沒被選到，改成用SQL語法重新指定 By Kin 2014/05/28
        gimBCitemCode.Filter = ""
        gimTCitemCode.Filter = ""
        Dim bCitemSQL As String
        Dim tCitemSQL As String
        bCitemSQL = "Select CodeNo ,Description From " & GetOwner & "CD019 Where 1 = 1  " & _
                        " And StopFlag <> 1 And PeriodFlag = 1  Order  By CodeNo "
        tCitemSQL = "Select CodeNo ,Description From " & GetOwner & "CD019 Where 1= 1" & _
                     " And StopFlag <> 1 Order By CodeNo  "
        'strCitemCode = "Select CitemCode From " & GetOwner & "CD068A Where  CD068A.BillHeadFmt = '" & cboBillHeadFmt.Text & "'"
        '#7049 改變畫面的收費項目 By Kin 2015/07/15
        Dim strCD068ACitem As String
        strCD068ACitem = "Select CitemCode From " & GetOwner & "CD068A Where  CD068A.BillHeadFmt = '" & cboBillHeadFmt.Text & "'"
        If Len(strCitemCode) > 0 Then
            

            bCitemSQL = "Select CodeNo ,Description From " & GetOwner & "CD019 Where CodeNo In (" & strCD068ACitem & ") " & _
                        " And Nvl(StopFlag,0)  =  0 And PeriodFlag = 1  Order  By CodeNo "
            
            tCitemSQL = "Select CodeNo ,Description From " & GetOwner & "CD019 Where CodeNo In (" & strCD068ACitem & ") " & _
                       " And Nvl(StopFlag,0) = 0 Order By CodeNo  "
                        
            SetMsQry gimBCitemCode, bCitemSQL, "收費項目代碼,收費項目名稱"
            SetMsQry gimTCitemCode, tCitemSQL, "收費項目代碼,收費項目名稱"
        Else
            SetMsQry gimBCitemCode, bCitemSQL, "收費項目代碼,收費項目名稱"
            SetMsQry gimTCitemCode, tCitemSQL, "收費項目代碼,收費項目名稱"
        End If
        '#2955 判斷所選的多媒体帳戶是否有啟用跨服務，如果有要過濾B單只剩P服務與非週期性的項目 By Kin 2007/04/19 For Jackie
        If rsBillHeadFmt("MutiServiceUnion") = 1 Then '啟動
            '#2955 如果有啟動P服務檢查收費項目裡面是否有I服務 By Kin 2007/04/24
             'blnIFlag = GetRsValue("Select Count(*) From " & GetOwner & "CD019 Where CodeNo In (" & strCitemCode & ") And ServiceType='I'", gcnGi)
             '#7049 改用CD068A.CitemCode By Kin 2015/07/14
             blnIFlag = GetRsValue("Select Count(*) From " & GetOwner & "CD019 Where CodeNo In (" & strCD068ACitem & ") And ServiceType='I'", gcnGi)
             
            '#2955 只留P服務與非週期性的收費項目
            '#7049 改用CD068A.CitemCode By Kin 2015/07/14
'            If Not GetRS(rsCD019Mutli, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & strCitemCode & ") And PeriodFlag=0 And ServiceType='P'", , adUseClient, adOpenKeyset) Then Exit Sub
            If Not GetRS(rsCD019Mutli, "Select CodeNo From " & GetOwner & "CD019 Where CodeNo In (" & strCD068ACitem & ") And PeriodFlag=0 And ServiceType='P'", , adUseClient, adOpenKeyset) Then Exit Sub
            If rsCD019Mutli.RecordCount > 0 Then
                strPBillType = "'B'"
                strPCitemStr = rsCD019Mutli.GetString(, , , ",")
                strPCitemStr = " IN (" & Left(strPCitemStr, Len(strPCitemStr) - 1) & ")"
            Else
                '雖然有啟動，但找不到任何的P服務的非週期性項目
                strPBillType = Empty
                strPCitemStr = Empty
                blnIFlag = False
            End If
        Else
            strPBillType = Empty
            strPCitemStr = Empty
            blnIFlag = False
        End If
        gimBCitemCode.SetQueryCode strCitemCode
        gimTCitemCode.SetQueryCode strCitemCode
        CloseRecordset rsCD019
        CloseRecordset rsCD019Mutli
End Sub



Private Sub cboTBillHeadFmt_Click()
    On Error Resume Next
    Dim strCitemCode As String
    Dim rsCD019  As New ADODB.Recordset
    Dim strCD019SQL As String
        If rsTBillHeadFmt.RecordCount = 0 Then Exit Sub
        rsTBillHeadFmt.AbsolutePosition = cboTBillHeadFmt.ListIndex + 1
        strCD019SQL = "Select CodeNo From " & GetOwner & "CD019 " & _
                            " Where CodeNo In ( Select CitemCode From " & GetOwner & "CD068A " & _
                            "         Where Exists (Select * From " & GetOwner & "CD068 Where CD068.BillHeadFmt = CD068A.BillHeadFmt " & _
                                                    " And CD068.BillHeadFmt = '" & rsTBillHeadFmt("BillHeadFmt") & "' )) And StopFlag <>1"
        If Not GetRS(rsCD019, strCD019SQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        If rsCD019.EOF Then Exit Sub
        strCitemCode = rsCD019.GetString(, , , ",")
        strCitemCode = Left(strCitemCode, Len(strCitemCode) - 1)
        gimTCitemCode.Filter = ""
        Dim tCitemSQL As String
        tCitemSQL = "Select CodeNo ,Description From " & GetOwner & "CD019 Where 1 = 1" & _
                     " And StopFlag <> 1 Order By CodeNo  "
 
        Dim strCD068ACitem As String
        strCD068ACitem = "Select CitemCode From " & GetOwner & "CD068A Where  CD068A.BillHeadFmt = '" & cboTBillHeadFmt.Text & "'"
        If Len(strCitemCode) > 0 Then
            tCitemSQL = "Select CodeNo ,Description From " & GetOwner & "CD019 Where CodeNo In (" & strCD068ACitem & ") " & _
                       " And Nvl(StopFlag,0) = 0 Order By CodeNo  "
                                    
            SetMsQry gimTCitemCode, tCitemSQL, "收費項目代碼,收費項目名稱"
        Else
            SetMsQry gimTCitemCode, tCitemSQL, "收費項目代碼,收費項目名稱"
        End If
             
        gimTCitemCode.SetQueryCode strCitemCode
        CloseRecordset rsCD019
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo chkErr
        Unload Me
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdCancel_Click"
End Sub

Private Sub cmdPrevRpt_Click()
  On Error GoTo chkErr
    PreviousRpt GetPrinterName(5), "SO3280A.rpt", "收費單號調整 [SO3280A]"
    PreviousRpt GetPrinterName(5), "SO3280A1.rpt", "收費單號調整 [SO3280A]"
  Exit Sub
chkErr:
    Call ErrSub(Me.Name, "cmdPrevRpt_Click")
End Sub

Private Sub Form_Activate()
    On Error GoTo chkErr
        SetPara
        Screen.MousePointer = vbDefault
        If blnAutoExec Then
            cmdOk.Value = True
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo chkErr
        If Not ChkGiList(KeyCode) Then Exit Sub
        If KeyCode = vbKeyF2 Then
            If cmdOk.Enabled Then Call cmdOK_Click
            KeyCode = 0
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "Form_KeyDown"
End Sub

Private Sub Form_Load()
    On Error GoTo ER
        subGim
        subGil
        Call CboAddItem
        DefaultValue
        chkShouldDateMerge.Value = GetRsValue("Select nvl(ShouldDateMerge,0) From " & GetOwner & "SO041 SO041 where SysID=" & gilCompCode.GetCodeNo, gcnGi)
        'chkShouldDateMerge.Value = blnShouldDateMerge * -1
        If chkShouldDateMerge Then
            gdaOweDate1.SetValue ""
            gdaOweDate2.SetValue ""
            gdaOweDate1.Enabled = False
            gdaOweDate2.Enabled = False
        End If
    Exit Sub
ER:
    If ErrHandle(, True, , "Form_Load") Then Resume 0
End Sub

Private Sub subGil()
    On Error GoTo chkErr
        SetgiList gilServiceType, "CodeNo", "Description", "CD046", , , , , 1, 12
        SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGil"
End Sub

Private Sub CboAddItem()
    On Error Resume Next
        'If cboBillHeadFmt.ListCount > 0 Then Exit Sub
        If Not GetRS(rsBillHeadFmt, "Select BillHeadFmt,CitemCodeStr,MutiServiceUnion From " & GetOwner & "CD068") Then Exit Sub
        Set rsTBillHeadFmt = rsBillHeadFmt.Clone
        Do While Not rsBillHeadFmt.EOF
            cboBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
            cboTBillHeadFmt.AddItem rsBillHeadFmt("BillHeadFmt") & ""
            rsBillHeadFmt.MoveNext
        Loop
End Sub

Private Sub DefaultValue()
    On Error GoTo chkErr
    Dim intPara23 As Integer
        gilCompCode.SetCodeNo garyGi(9)
        gilCompCode.Query_Description
        
        gimBillType.SetDispStr "臨時收費單"
        gimBillType.SetQueryCode "T"
        
        '處理日期
        gdaProcessDate.SetValue RightDate
        ' 收集年月範圍 = 當日之年月至當日之年月
        gymClctYM1.SetValue (Format(RightNow, "YYYYMM"))
        gymClctYM2.SetValue (Format(RightNow, "YYYYMM"))
        intPara24 = GetSystemParaItem("Para24", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043")
        intPara23 = Val(GetSystemParaItem("Para23", gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, "SO043") & "")
        If intPara24 = 1 Then
            gilServiceType.Clear
            gilServiceType.Visible = False
            lblServiceType.Visible = False
            lblBillHeadFmt.Visible = True
            cboBillHeadFmt.Visible = True
            cboTBillHeadFmt.Visible = True
            cboBillHeadFmt.ListIndex = 0
            cboTBillHeadFmt.ListIndex = 0
            'gimBCitemCode.Enabled = False
            'gimTCitemCode.Enabled = False
            '  用媒體入帳才Lock
            'If intPara23 = 2 Then
            '    gimBCitemCode.Enabled = False
            '    gimTCitemCode.Enabled = False
            'End If
        Else
            gilServiceType.ListIndex = 1
            gilServiceType.Visible = True
            lblServiceType.Visible = True
            lblBillHeadFmt.Visible = False
            cboBillHeadFmt.Visible = False
            cboTBillHeadFmt.Visible = False
        End If
        
    Exit Sub
chkErr:
    ErrSub Me.Name, "DefalutValue"
End Sub

Private Sub subGim()
    On Error GoTo chkErr
        SetgiMulti gimBCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱", , True
        SetgiMulti gimTCitemCode, "CodeNo", "Description", "CD019", "收費項目代碼", "收費項目名稱", , True
        SetgiMulti gimBUCCode, "CodeNo", "Description", "CD013", "未收費原因代碼", "未收費原因名稱", " Where Nvl(REFNO,0) NOT IN(3,7) AND NVL(PayOK,0)=0 ", True
        SetgiMulti gimTUCCode, "CodeNo", "Description", "CD013", "未收費原因代碼", "未收費原因名稱", " Where Nvl(REFNO,0) NOT IN(3,7) AND NVL(PayOK,0)=0 ", True
        SetgiMulti gimPayType, "CodeNo", "Description", "CD112", "代碼", "名稱"
        Call SetgiMultiAddItem(gimBillType, "I,M,P,T", "裝機單,維修單,拆機單,臨時收費單", "單據代號", "單據名稱")
        If GetPaynowFlag Then
'            gimPayType.SetQueryCode "0"
            '#5925 繳費類別增加預設值 By Kin 2011/04/12
            With gimPayType
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
            gimPayType.Clear
            gimPayType.Enabled = False
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "subGim"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        
        Call ReleaseCOM(Me)
End Sub

Private Sub gdaOweDate1_GotFocus()
    On Error Resume Next
    If gdaOweDate1.GetValue = "" Then gdaOweDate1.SetValue (RightDate)
End Sub

Private Sub gdaOweDate2_GotFocus()
    On Error Resume Next
    If gdaOweDate1.GetValue <> "" Then gdaOweDate2.SetValue gdaOweDate1.GetValue
End Sub

Private Sub gdaShouldDate1_GotFocus()
    On Error Resume Next
        If gdaShouldDate1.GetValue = "" Then gdaShouldDate1.SetValue RightDate
End Sub

Private Sub gdaShouldDate2_GotFocus()
    On Error Resume Next
    '#2955測試不OK,要帶與gdaShouldDate1一樣的值,不是RightDate By Kin 2007/06/15 For Penny
        If gdaShouldDate1.GetValue <> "" Then gdaShouldDate2.SetValue gdaShouldDate1.GetValue
End Sub

Private Sub gdaShouldDate2_Validate(Cancel As Boolean)
    On Error Resume Next
        If gdaShouldDate1.GetValue = "" Or gdaShouldDate2.GetValue = "" Then Exit Sub
        If gdaShouldDate2.GetValue < gdaShouldDate1.GetValue Then
            MsgDateRangeX
            Cancel = True
        End If
End Sub

Private Sub gdtCreateTime1_GotFocus()
    On Error Resume Next
    If gdtCreateTime1.GetValue = "" Then gdtCreateTime1.SetValue (RightDate)
End Sub

Private Sub gdtCreateTime2_GotFocus()
  On Error Resume Next
    If gdtCreateTime1.GetValue <> "" Then gdtCreateTime2.SetValue GetDayLastTime(gdtCreateTime1.GetValue(True))
End Sub

Private Sub gilCompCode_Change()
    On Error Resume Next
    Dim rsTmp As New ADODB.Recordset
    Dim strTmpSql As String
        If Not ChgComp(gilCompCode, "SO3310", "SO3311") Then Exit Sub
        Call subGil
        Call subGim
        Call SelectServType(gilCompCode.GetCodeNo, gilServiceType)
        gilServiceType.SetCodeNo ""
        gilServiceType.Query_Description
        gilServiceType.ListIndex = 1
        '問題集 2761 測試不ok，因為文件規格有異動需加入單據合併調整項目 2006/12/13
        '#4077 增加UpdCreateEn參數 By Kin 2008/09/18
        '#7397 Add 3 datafield on the  screen conditions that are unable to be modified by kin 2017/04/05
        strTmpSql = "Select nvl(UpdShoulddate,0) as UpdShoulddate,nvl(UpdCreatetime,0) as UpdCreatetime," & _
                    "nvl(UpdCMCode,0) as UpdCMCode,nvl(UpdCreateEn,0) as UpdCreateEn,   " & _
                    "nvl(UpdUCcode,0) as UpdUCcode,nvl(UpdBillCloseDate,0) as UpdBillCloseDate,nvl(UpdPrtSNo,0) as UpdPrtSNo " & _
                    " from " & GetOwner & _
                    "SO041 where SysID=" & gilCompCode.GetCodeNo
        If Not GetRS(rsTmp, strTmpSql) Then Exit Sub
        chkShouldDate = Val(rsTmp("UpdShoulddate") & "")
        chkCreateTime = Val(rsTmp("UpdCreatetime") & "")
        chkCmCode = Val(rsTmp("UpdCMCode") & "")
        chkUpdCreateEn = Val(rsTmp("UpdCreateEn") & "")
        chkUpdUCcode = Val(rsTmp("UpdUCcode") & "")
        chkUpdBillCloseDate = Val(rsTmp("UpdBillCloseDate") & "")
        chkUpdPrtSNo = Val(rsTmp("UpdPrtSNo") & "")
        
        CloseRecordset rsTmp
End Sub

Private Sub gilServiceType_Change()
    On Error Resume Next
        Call GiMultiFilter(gimBCitemCode, gilServiceType.GetCodeNo)
        gimBCitemCode.Filter = gimBCitemCode.Filter & IIf(Trim(gimBCitemCode.Filter) = "", " Where ", " And ") & " PeriodFlag = 1"
        Call GiMultiFilter(gimBUCCode, gilServiceType.GetCodeNo, , , " Nvl(REFNO,0) NOT IN(3,7) AND NVL(PayOK,0)=0 And Nvl(StopFlag,0)=0 ")
        Call GiMultiFilter(gimTCitemCode, gilServiceType.GetCodeNo)
        Call GiMultiFilter(gimTUCCode, gilServiceType.GetCodeNo, , , " Nvl(REFNO,0) NOT IN(3,7) AND NVL(PayOK,0)=0 And Nvl(StopFlag,0)=0 ")
        
End Sub





Private Sub gymClctYM1_Validate(Cancel As Boolean)
    On Error GoTo chkErr
        Call CheckYM
    Exit Sub
chkErr:
    ErrSub Me.Name, "gymClctYM1_Validate"
End Sub

Private Sub gymClctYM2_Validate(Cancel As Boolean)
    On Error GoTo chkErr
        Call CheckYM
    Exit Sub
chkErr:
    ErrSub Me.Name, "gymClctYM2_Validate"
End Sub

Private Sub CheckYM()
    On Error GoTo chkErr
        If gymClctYM1.GetValue <> "" And gymClctYM2.GetValue <> "" Then
            If gymClctYM2.GetValue < gymClctYM1.GetValue Then
                gymClctYM2.SetValue (gymClctYM1.GetValue)
            End If
        End If
        If gymClctYM1.GetValue <> "" And gymClctYM2.GetValue = "" Then
            gymClctYM2.SetValue (gymClctYM1.GetValue)
        End If
        If gymClctYM2.GetValue <> "" And gymClctYM1.GetValue = "" Then
            gymClctYM1.SetValue (gymClctYM2.GetValue)
        End If
        '合併單據應收日期
        If gdaOweDate2.GetValue <> "" And gdaOweDate1.GetValue = "" Then
            gdaOweDate1.SetValue (gdaOweDate2.GetValue)
        End If
        If gdaOweDate1.GetValue <> "" And gdaOweDate2.GetValue = "" Then
            gdaOweDate2.SetValue (gdaOweDate1.GetValue)
        End If
        '合併單據產生日期
        If gdtCreateTime2.GetValue <> "" And gdtCreateTime1.GetValue = "" Then
            gdtCreateTime1.SetValue (gdtCreateTime2.GetValue)
        End If
        If gdtCreateTime1.GetValue <> "" And gdtCreateTime2.GetValue = "" Then
            gdtCreateTime2.SetValue (gdtCreateTime1.GetValue)
        End If
    Exit Sub
chkErr:
    ErrSub Me.Name, "CheckYM"
End Sub

Private Function IsDataOk()
    On Error GoTo chkErr
        IsDataOk = False
        
        If Not ChkDTok Then Exit Function
        If Not MustExist(gilCompCode, 2, "公司別") Then Exit Function
        If intPara24 = 0 Then If Not MustExist(gilServiceType, 2, "服務類別") Then Exit Function
        If Not MustExist(gdaProcessDate, 1, "處理日期") Then Exit Function
        If Not MustExist(gymClctYM1, 1, "收集年月起始日") Then Exit Function
        If Not MustExist(gymClctYM2, 1, "收集年月截止日") Then Exit Function
        If gimBCitemCode.GetQueryCode = "" Then MsgMustBe ("B單收費項目"): gimBCitemCode.SetFocus: Exit Function
        If Not MustExist(gdaShouldDate1, 1, "應收起始日期") Then Exit Function
        If Not MustExist(gdaShouldDate2, 1, "應收截止日期") Then Exit Function
        If gimTCitemCode.GetQueryCode = "" Then MsgMustBe ("T單收費項目"): gimTCitemCode.SetFocus: Exit Function
        If gdaOweDate1.GetValue > gdaOweDate2.GetValue Then MsgBox "應收起始日期不得大於終止日!!", vbInformation, "警告訊息": gdaOweDate2.SetFocus: Exit Function
        If gdtCreateTime1.GetValue > gdtCreateTime2.GetValue Then MsgBox "產生起始時間不得大於終止時間!!", vbInformation, "警告訊息": gdtCreateTime2.SetFocus: Exit Function
        IsDataOk = True
    Exit Function
chkErr:
    ErrSub Me.Name, "IsDataOk"
End Function
' 3.  確定鍵或F2按下:
Private Sub cmdOK_Click()
    On Error GoTo chkErr
    Dim strMsg As String
    Dim rsTmp As New ADODB.Recordset
    Dim intCount As Long
    Dim intError As Integer
    Dim aryCustId() As String
    Dim blnShwCombItem As Boolean
    Dim strPayKinds As String
    blnShwCombItem = False
    
        Call CheckYM
        '‧  必要欄位檢查, 若不過不可繼續。 除了複選元件外, 所有欄位皆為必要
        strCustIdStr = Empty
        '#3478 增加客戶編號條件 By Kin 2008/02/21
        If Not IsDataOk Then Exit Sub
        If Trim(txtCustId) <> "" Then
            Call ParseWord(txtCustId, aryCustId)
            strCustIdStr = Join(aryCustId, ",")
                   
        End If
        '呼叫取參數 2011/11/30 Jacky 6153 MinChen
        If blnGetPara Then
            Call GetPara
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        ReadyGoPrint
        
        Call subChoose
        strPayKinds = ""
        '#5756 增加繳付類別參數 By Kin 2010/08/30
        If (gimPayType.GetQryStr <> "") And (InStr(1, gimPayType.GetQryStr, "全選") <= 0) Then
            strPayKinds = gimPayType.GetQryStr
        End If
        If SF_CombineBT(garyGi(1), gilCompCode.GetCodeNo, gilServiceType.GetCodeNo, _
                        gdaProcessDate.GetValue, gymClctYM1.GetValue, gymClctYM2.GetValue, "IN (" & gimBCitemCode.GetQueryCode & ")", _
                        gimBUCCode.GetQryStr, gdaShouldDate1.GetValue, gdaShouldDate2.GetValue, "IN (" & gimTCitemCode.GetQueryCode & ")", _
                        gimTUCCode.GetQryStr, gimBillType.GetQueryCode, chkCombine.Value, gdaOweDate1.GetValue, gdaOweDate2.GetValue, _
                        gdtCreateTime1.GetValue, gdtCreateTime2.GetValue, strPBillType, strPCitemStr, blnIFlag, strCustIdStr, strPayKinds, _
                        cboBillHeadFmt.Text, cboTBillHeadFmt.Text, _
                        strMsg) <= 0 Then
             
            MsgBox strMsg, vbCritical, "警告"
        Else
            If Not blnAutoExec Then
                MsgBox strMsg, vbInformation, gimsgPrompt
                strSQL = "Select * From " & GetOwner & "SO109 Where CombineId = 0 "
                intError = 1
                intCount = GetRsValue("Select count(*) As intCount From " & GetOwner & "SO109 SO109 Where CombineId = 0 ")
                intError = 2
                
                If intCount = 0 Then
                    MsgBox "查無帳單合併資料!!", vbInformation, "訊息"
                Else
                    blnShwCombItem = True
                    Call PrintRpt(GetPrinterName(5), "SO3280A.rpt", "SO3280", Me.Caption, strSQL, strChooseString, _
                                        , True, , , "Para24 = " & intPara24, GiPaperPortrait)
                End If
    
                If chkCombine.Value = 0 Then
                    Dim S As String
                    
                    strSQL = "Select * From " & GetOwner & "SO109 Where CombineId = 1 "
                    If rsTmp.State = adStateOpen Then rsTmp.Close
                    intError = 3
                    
                    intCount = GetRsValue("Select count(*) As intCount From " & GetOwner & "SO109 SO109 Where CombineId = 1 ")
                    intError = 4
                    
                    If intCount = 0 Then
                     MsgBox "查無帳單未合併資料!!", vbInformation, "訊息"
                    Else
                     Call PrintRpt(GetPrinterName(5), "SO3280A1.rpt", "SO3280", Me.Caption, strSQL, strChooseString, , True, , , , GiPaperPortrait)
                    End If
    
                End If
                '#5439 如果有啟動合併帳單要Show出提醒訊息 By Kin 2010/01/06
                If blnShwCombItem Then
                    If Val(GetRsValue("SELECT NVL(STARTCOMBCITEM,0) FROM " & GetOwner & "SO041", gcnGi) & "") = 1 Then
                        MsgBox "請記得重新執行收費項目合併作業!!", vbInformation, "訊息"
                    End If
                End If
            Else
                blnIsOK = True
                strReturnMsg = strMsg
                Unload Me
            End If
        End If
        If rsTmp.State = adStateOpen Then rsTmp.Close
        Set rsTmp = Nothing
        Screen.MousePointer = vbDefault
    Exit Sub
chkErr:
    ErrSub Me.Name, "cmdOk_Click: ErrorNo:" & intError
End Sub
Public Function SF_CombineBT(ByVal p_UserId As String, ByVal p_CompCode As String, _
    ByVal p_ServiceType As String, ByVal p_ProcDate As Long, ByVal p_YM1 As String, _
    ByVal p_YM2 As String, ByVal p_BCitemStr As String, ByVal p_BUCCode As String, _
    ByVal p_ShouldDate1 As String, ByVal p_ShouldDate2 As String, ByVal p_TCitemStr As String, _
    ByVal p_TUCCode As String, ByVal p_TBillType As String, ByVal p_IMPTCombine As String, _
    ByRef p_BShouldDate1 As String, ByRef p_BShouldDate2 As String, _
    ByRef p_BCreateTime1 As String, ByRef p_BCreateTime2 As String, _
    ByRef p_PBillType As String, ByRef p_PCitemStr As String, ByVal p_iServiceFlag As Boolean, _
    ByVal p_CustStr As String, ByVal p_PayKindSql As String, ByVal p_BillHeadFmt As String, _
    ByVal p_BillHeadFmtT As String, _
    ByRef P_RETMSG As String) As Long
    '#5756 增加p_PayKindSql參數 By Kin 2010/08/30
    On Error GoTo chkErr
        Dim ComSF_CombineBT As New ADODB.Command

        With ComSF_CombineBT
                .Parameters.Append .CreateParameter("RETURN_VALUE", adVarNumeric, adParamReturnValue)
                .Parameters.Append .CreateParameter("p_UserId", adVarChar, adParamInput, 2000, p_UserId)
                .Parameters.Append .CreateParameter("p_CompCode", adVarNumeric, adParamInput, 2000, p_CompCode)
                .Parameters.Append .CreateParameter("p_ServiceType", adVarChar, adParamInput, 2000, p_ServiceType)
                .Parameters.Append .CreateParameter("p_ProcDate", adVarChar, adParamInput, 2000, p_ProcDate)
                .Parameters.Append .CreateParameter("P_YM1", adVarNumeric, adParamInput, 2000, p_YM1)
                .Parameters.Append .CreateParameter("P_YM2", adVarNumeric, adParamInput, 2000, p_YM2)
                '.Parameters.Append .CreateParameter("p_BCitemStr", adVarChar, adParamInput, 20000, p_BCitemStr)
                '#7049 測試不OKp_BCitemStr 隨便給值就好
                .Parameters.Append .CreateParameter("p_BCitemStr", adVarChar, adParamInput, 4000, "0")
                .Parameters.Append .CreateParameter("p_BUCCode", adVarChar, adParamInput, 2000, IIf(p_BUCCode = "", Null, p_BUCCode))
                .Parameters.Append .CreateParameter("p_ShouldDate1", adVarChar, adParamInput, 2000, p_ShouldDate1)
                .Parameters.Append .CreateParameter("p_ShouldDate2", adVarChar, adParamInput, 2000, p_ShouldDate2)
'                .Parameters.Append .CreateParameter("p_TCitemStr", adVarChar, adParamInput, 20000, p_TCitemStr)
                .Parameters.Append .CreateParameter("p_TCitemStr", adVarChar, adParamInput, 2000, "0")
                .Parameters.Append .CreateParameter("p_TUCCode", adVarChar, adParamInput, 2000, IIf(p_TUCCode = "", Null, p_TUCCode))
                .Parameters.Append .CreateParameter("p_TBillType", adVarChar, adParamInput, 2000, IIf(p_TBillType = "", Null, p_TBillType))
                .Parameters.Append .CreateParameter("p_IMPTCombine", adVarChar, adParamInput, 2000, p_IMPTCombine)
                .Parameters.Append .CreateParameter("p_BShouldDate1", adVarChar, adParamInput, 2000, IIf(p_BShouldDate1 = "", Null, p_BShouldDate1))
                .Parameters.Append .CreateParameter("p_BShouldDate2", adVarChar, adParamInput, 2000, IIf(p_BShouldDate2 = "", Null, p_BShouldDate2))
                .Parameters.Append .CreateParameter("p_BCreateTime1", adVarChar, adParamInput, 2000, IIf(p_BCreateTime1 = "", Null, p_BCreateTime1))
                .Parameters.Append .CreateParameter("p_BCreateTime2", adVarChar, adParamInput, 2000, IIf(p_BCreateTime2 = "", Null, p_BCreateTime2))
                '#2955 增加傳入後端3個參數 1.跨服務別的單據類型 2.跨服務別的收費項目 3.收費項目是否有包含I服務 By Kin 2007/04/24 For Jackie
                .Parameters.Append .CreateParameter("p_PBillType", adVarChar, adParamInput, 2000, IIf(strPBillType & "" = "", Null, strPBillType))
                .Parameters.Append .CreateParameter("p_PCitemStr", adVarChar, adParamInput, 4000, IIf(strPCitemStr & "" = "", Null, strPCitemStr))
                .Parameters.Append .CreateParameter("p_iServiceFlag", adVarNumeric, adParamInput, 2000, blnIFlag * -1)
                '#3478 增加客戶編號參數 By Kin 2008/02/21
                .Parameters.Append .CreateParameter("p_CustStr", adVarChar, adParamInput, 4000, IIf(strCustIdStr = Empty, Null, strCustIdStr))
                .Parameters.Append .CreateParameter("p_PayKindSql", adVarChar, adParamInput, 2000, p_PayKindSql)
                '#7049 增加 多帳戶產生依據 By Kin 2015/07/14
                .Parameters.Append .CreateParameter("p_BillHeadFmt", adVarChar, adParamInput, 2000, p_BillHeadFmt)
                .Parameters.Append .CreateParameter("p_BillHeadFmtT", adVarChar, adParamInput, 2000, p_BillHeadFmtT)
                .Parameters.Append .CreateParameter("P_RETMSG", adVarChar, adParamOutput, 2000, P_RETMSG)
                
                Set .ActiveConnection = gcnGi
                .CommandText = "SF_CombineBT"
                .CommandType = adCmdStoredProc
                .Execute
                P_RETMSG = .Parameters("P_RETMSG").Value & ""
                SF_CombineBT = Val(.Parameters("Return_value").Value & "")
        End With
   
        Exit Function
chkErr:
        SF_CombineBT = -99: P_RETMSG = "程式產生:其他錯誤"
        ErrSub Me.Name, "SF_CombineBT"
End Function

Private Sub subChoose()
On Error Resume Next

    strChooseString = "公司別  : " & gilCompCode.GetDescription & ";" & _
                      "服務類別: " & gilServiceType.GetDescription & ";" & _
                      "收集年月: " & gymClctYM1.GetValue(True) & "~" & gymClctYM2.GetValue(True) & ";" & _
                      "B單收費項目: " & gimBCitemCode.GetDispStr & ";" & _
                      "B單未收費原因: " & gimBUCCode.GetDispStr & ";" & _
                      "應收日期  : " & gdaShouldDate1.GetValue(True) & "~" & gdaShouldDate2.GetValue(True) & ";" & _
                      "T單收費項目  : " & gimTCitemCode.GetDispStr & ";" & _
                      "T單未收費原因: " & gimTCitemCode.GetDispStr
End Sub

Private Sub txtCustId_KeyPress(keyAscii As Integer)
  On Error Resume Next
    If keyAscii >= 48 And keyAscii <= 57 Or keyAscii = 44 Or keyAscii = 8 Or keyAscii = 45 Then
        If keyAscii = 44 Or keyAscii = 45 Then
            If Asc(Right(" " & txtCustId.Text, 1)) = 44 Or Asc(Right(" " & txtCustId.Text, 1)) = 45 Or txtCustId = "" Then keyAscii = 9
        End If
        If Not ChkMaxLengthOK(txtCustId) Then
           If Not (keyAscii = 44 Or keyAscii = 8 Or keyAscii = 45) Then keyAscii = 9
        End If
    Else
        keyAscii = 9
    End If
End Sub
