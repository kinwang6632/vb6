VERSION 5.00
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{074771B7-6ADE-11D5-B178-00485456B028}#2.0#0"; "csMulti.ocx"
Begin VB.Form frmSO11A1A 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "SO11A1A[ ���~�q�ʩ��Ӭd�� ]"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   6525
   Icon            =   "SO11A1A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6525
   StartUpPosition =   1  '���ݵ�������
   Begin VB.CommandButton cmdExit 
      Caption         =   "���}(&X)"
      Height          =   405
      Left            =   5370
      TabIndex        =   9
      Top             =   2160
      Width           =   1005
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "F2.�T�w"
      Height          =   405
      Left            =   180
      TabIndex        =   8
      Top             =   2160
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   "�d�߱���"
      Height          =   1515
      Left            =   60
      TabIndex        =   0
      Top             =   510
      Width           =   6375
      Begin VB.CheckBox chkHistoryData 
         Caption         =   "��ܾ��v���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4590
         TabIndex        =   12
         Top             =   300
         Width           =   1545
      End
      Begin CS_Multi.CSmulti gimOrderWay 
         Height          =   405
         Left            =   150
         TabIndex        =   5
         Top             =   720
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   714
         ButtonCaption   =   "�q�ʺ޹D"
      End
      Begin Gi_Date.GiDate gdaOrderDate1 
         Height          =   345
         Left            =   1290
         TabIndex        =   2
         Top             =   270
         Width           =   1395
         _ExtentX        =   2461
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
      Begin Gi_Date.GiDate gdaOrderDate2 
         Height          =   345
         Left            =   3030
         TabIndex        =   4
         Top             =   270
         Width           =   1395
         _ExtentX        =   2461
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
      Begin CS_Multi.CSmulti gimOrderEn 
         Height          =   405
         Left            =   180
         TabIndex        =   6
         Top             =   1680
         Visible         =   0   'False
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   714
         ButtonCaption   =   "�q�ʤH��"
      End
      Begin CS_Multi.CSmulti gimSTBSNO 
         Height          =   405
         Left            =   150
         TabIndex        =   7
         Top             =   1050
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   714
         ButtonCaption   =   "STB�]��"
      End
      Begin VB.Label Label2 
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2790
         TabIndex        =   3
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Label1 
         Caption         =   "�q�ʤ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   885
      End
   End
   Begin prjGiList.GiList gilCompCode 
      Height          =   315
      Left            =   990
      TabIndex        =   10
      Top             =   90
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      F2Corresponding =   -1  'True
   End
   Begin VB.Label Label3 
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
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmSO11A1A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsPvod As ADODB.Recordset

Private Sub cmdExit_Click()
  On Error Resume Next
    CloseRecordset rsPvod
    Unload Me
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ChkErr
    Screen.MousePointer = vbHourglass
    
    Call ClearWhere
    If gdaOrderDate1.Text <> Empty Then
        fOrderDate1 = gdaOrderDate1.GetValue
    End If
    If gdaOrderDate2.Text <> Empty Then
        fOrderDate2 = gdaOrderDate2.GetValue
    End If
    If gimOrderWay.GetQryStr <> Empty Then
        fOrderWayCode = gimOrderWay.GetQueryCode
    End If
    If gimOrderEn.GetQryStr <> Empty Then
        fOrderEn = gimOrderEn.GetQueryCode
    End If
    If gimSTBSNO.GetQryStr <> Empty Then
        fSTBSeqNos = gimSTBSNO.GetQryFld(3)
    End If
    fShowHistory = chkHistoryData.Value
    Set rsPvod = GetQryData
    If rsPvod.RecordCount <= 0 Then
        MsgBox "�d�L�����ơI", vbInformation, "�T��"
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        Call ShowDataView
    End If
    Screen.MousePointer = vbDefault
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "cmdOK_Click")
End Sub
Private Sub ClearWhere()
  On Error Resume Next
    fBillNo = Empty
    fBillNoAndItem = Empty
    fBillNoItem = Empty
    fOrderDate1 = Empty
    fOrderDate2 = Empty
    fSTBSeqNos = Empty
    fOrderWayCode = Empty
    fOrderEn = Empty
    fSTBSeqNo = Empty
    fCitemCode = Empty
    fProductType = 99
End Sub

Private Sub DefaultValue()
  On Error GoTo ChkErr
'    Dim rsField As New ADODB.Recordset
'    Dim i As Long
'
    If fOrderDate1 <> Empty Then
        gdaOrderDate1.SetValue fOrderDate1
    Else
        gdaOrderDate1.SetValue Format(RightNow, "yyyymm") & "01"
    End If
    If fOrderDate2 <> Empty Then
        gdaOrderDate2.SetValue fOrderDate2
    Else
        gdaOrderDate2.SetValue GetLastDate(RightNow)
    End If
    If fOrderEn <> Empty Then
        gimOrderEn.SetQueryCode fOrderEn
    End If
    If fOrderWayCode <> Empty Then
        gimOrderWay.SetQueryCode fOrderWayCode
    End If

    
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "DefaultValue")
End Sub
'Private Sub ClearWhere()
'  On Error Resume Next
'    fOrderDate1 = Empty
'    fOrderDate2 = Empty
'    fOrderEn = Empty
'    fOrderWayCode = Empty
'End Sub

Private Sub Form_Activate()
  On Error Resume Next
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
    If KeyCode = vbKeyF2 Then
        cmdOK_Click
    End If
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
  On Error GoTo ChkErr
    
    
    Call subGil
    Call DefaultValue
    'Call subGim
    
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "Form_Load")
End Sub
Public Property Let uOrderDate1(ByVal vData As String)
  On Error Resume Next
    fOrderDate1 = vData
End Property
Public Property Let uOrderDate2(ByVal vData As String)
  On Error Resume Next
    fOrderDate2 = vData
End Property
Public Property Let uOrderEn(ByVal vData As String)
  On Error Resume Next
    fOrderEn = vData
    
End Property
Public Property Let uOrderWayCode(ByVal vData As String)

End Property
Private Sub subGim()
  On Error GoTo ChkErr
    Dim strSTBQry As String
    strSTBQry = Empty
    strSTBQry = "SELECT ROWNUM,FaciSno,SEQNO FROM " & GetOwner & "SO004 " & _
            " WHERE CUSTID = " & gCustId & " AND COMPCODE = " & gCompCode & _
            " AND SEQNO IN ( SELECT STBSNO  FROM " & GetOwner & "SO004 " & _
            " WHERE CUSTID = " & gCustId & " AND COMPCODE = " & gCompCode & _
            " AND DVRsizecode IS Not Null And CMOPENDATE IS NOT NULL " & _
            " AND NVL(CMOpenDate,TO_DATE('10000101','YYYYMMDD')) > NVL(CMCloseDate,TO_DATE('10000101','YYYYMMDD')))"
    Call SetgiMulti(gimOrderEn, "EmpNo", "EmpName", "CM003", "�N�X", "�W��", , True)
    Call SetgiMulti(gimOrderWay, "CodeNo", "Description", "CD114", "�N�X", "�W��", , True)
    SetGimQry gimSTBSNO, strSTBQry, _
            , , "����,�]�ƧǸ�,�]�Ƭy����", "1000,3000,0", False
'
'    SetMsQry csmCitem2, strQry, _
'                             , , "�N�X,���O����,����,���B,���b�b��,���O�覡,�_�l��,�I���,�Ǹ�,PK,RowID,���~�Ǹ�,�ӽФH�m�W,���ئW��", _
'                             "500,1800,460,560,1800,850,850,850,770,1,1,850,1600,1600", True, , , , , _
'                             , , , , , , , , , , , , , , , , , , , , , "###,###,###", , , "EE/MM/DD", "EE/MM/DD", , , , , 10
    
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGim")
End Sub
Private Sub subGil()
  On Error GoTo ChkErr
    SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
    If gCompCode = Empty Then
        gCompCode = garyGi(9)
    End If
    gilCompCode.SetCodeNo gCompCode
    gilCompCode.Query_Description
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "subGil")
End Sub
Private Sub gilCompCode_Change()
  On Error GoTo ChkErr
    
    garyGi(16) = gilCompCode.GetOwner
    'garyGi(16) = Empty
    gCompCode = gilCompCode.GetCodeNo
    Call subGil
    Call subGim
  Exit Sub
ChkErr:
  Call ErrSub(Me.Name, "gilCompCode_Change")
End Sub


