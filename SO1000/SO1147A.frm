VERSION 5.00
Object = "{005346C2-8184-11D3-BD70-0080C8F80BC4}#5.5#0"; "GiList.ocx"
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Object = "{445AE19A-297E-11D3-A5F8-0080AD912008}#43.3#0"; "GiDate.ocx"
Object = "{82351433-9094-11D1-A24B-00A0C932C7DF}#1.5#0"; "AniGIF.ocx"
Object = "{A0C51418-555F-11D3-BC10-0050BA01A3A9}#9.9#0"; "GiMulti.ocx"
Begin VB.Form frmSO1147A 
   BorderStyle     =   1  '��u�T�w
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   ForeColor       =   &H8000000F&
   Icon            =   "SO1147A.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   11160
   StartUpPosition =   1  '���ݵ�������
   Visible         =   0   'False
   Begin VB.PictureBox pic1 
      Height          =   675
      Left            =   3660
      ScaleHeight     =   615
      ScaleWidth      =   3555
      TabIndex        =   29
      Top             =   2835
      Visible         =   0   'False
      Width           =   3615
      Begin AniGIFCtrl.AniGIF AniGIF1 
         Height          =   480
         Left            =   120
         TabIndex        =   30
         Top             =   60
         Width           =   480
         BackColor       =   12632256
         PLaying         =   -1  'True
         Transparent     =   -1  'True
         Speed           =   1
         Stretch         =   0
         AutoSize        =   -1  'True
         SequenceString  =   ""
         Sequence        =   0
         HTTPProxy       =   ""
         HTTPUserName    =   ""
         HTTPPassword    =   ""
         MousePointer    =   0
         GIF             =   "SO1147A.frx":0442
         ExtendWidth     =   847
         ExtendHeight    =   847
         Loop            =   0
         AutoRewind      =   0   'False
         Synchronized    =   -1  'True
      End
      Begin VB.Label Label4 
         Appearance      =   0  '����
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "�PMSS���ƨt�Τ�����,�еy��..."
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   690
         TabIndex        =   31
         Top             =   210
         Width           =   2700
      End
   End
   Begin VB.Frame fraForm 
      BorderStyle     =   0  '�S���ؽu
      Height          =   6525
      Left            =   0
      TabIndex        =   12
      Top             =   -105
      Width           =   11085
      Begin VB.OptionButton optP 
         Caption         =   "�u��"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6930
         TabIndex        =   34
         Top             =   6090
         Width           =   675
      End
      Begin VB.OptionButton optT 
         Caption         =   "T��"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6180
         TabIndex        =   33
         Top             =   6090
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.CheckBox chkMSS 
         Caption         =   "�ϥΪF�Ƥ����ҲչB��"
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   8445
         TabIndex        =   27
         Top             =   270
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Frame Frame2 
         Height          =   3915
         Left            =   195
         TabIndex        =   15
         Top             =   675
         Width           =   10650
         Begin VB.CommandButton btnExcel 
            Caption         =   "�պ����h���B ....."
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6495
            TabIndex        =   4
            Top             =   3420
            Width           =   1755
         End
         Begin prjGiGridR.GiGridR ggrData 
            Height          =   2205
            Left            =   195
            TabIndex        =   2
            Top             =   600
            Width           =   10230
            _ExtentX        =   18045
            _ExtentY        =   3889
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Gi_Date.GiDate gdCancelDate 
            Height          =   315
            Left            =   9210
            TabIndex        =   1
            Top             =   210
            Width           =   1080
            _ExtentX        =   1905
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
         Begin Gi_Multi.GiMulti gimStbItem 
            Height          =   375
            Left            =   180
            TabIndex        =   3
            Top             =   2955
            Width           =   10290
            _ExtentX        =   18150
            _ExtentY        =   661
            ButtonCaption   =   "STB �t��N�X"
            DataType        =   2
            FldCaption1     =   "�t��N�X"
            FldCaption2     =   "�t��W��"
            FontSize        =   9
            FontName        =   "�s�ө���"
         End
         Begin VB.Label lblCaption 
            BackColor       =   &H8000000A&
            Caption         =   "�ݵ�(��)Box "
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   450
            TabIndex        =   20
            Top             =   285
            Width           =   1770
         End
         Begin VB.Label Label1 
            Caption         =   "���M��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   8550
            TabIndex        =   19
            Top             =   255
            Width           =   780
         End
         Begin VB.Label Label2 
            Caption         =   "�t�󦩰����B"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   870
            TabIndex        =   18
            Top             =   3495
            Width           =   1200
         End
         Begin VB.Label lblAmount 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
            BackColor       =   &H80000005&
            Caption         =   "0"
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   2130
            TabIndex        =   17
            Top             =   3435
            Width           =   1305
         End
         Begin VB.Label lblCheckAmount 
            Alignment       =   1  '�a�k���
            Appearance      =   0  '����
            BackColor       =   &H80000005&
            Caption         =   "0"
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   8355
            TabIndex        =   16
            Top             =   3450
            Width           =   1635
         End
      End
      Begin VB.CommandButton cmdReciveAmount 
         Caption         =   "���ͦ��O���(&C)"
         Enabled         =   0   'False
         Height          =   345
         Left            =   8745
         TabIndex        =   10
         Top             =   6075
         Width           =   1650
      End
      Begin VB.Frame fraRenttoSale 
         Appearance      =   0  '����
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   660
         TabIndex        =   13
         Top             =   4740
         Width           =   4155
         Begin VB.CheckBox chkRent 
            Appearance      =   0  '����
            BackColor       =   &H8000000A&
            Caption         =   "����P ( �Ӳ����ӯ��� �A�Φ��� )"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   14
            Top             =   195
            Width           =   2985
         End
      End
      Begin prjGiList.GiList gilReturnItme 
         Height          =   300
         Left            =   7470
         TabIndex        =   7
         Top             =   5640
         Width           =   2910
         _ExtentX        =   5133
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
         FldWidth1       =   600
         FldWidth2       =   1980
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilItem 
         Height          =   270
         Left            =   1860
         TabIndex        =   9
         Top             =   5700
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   476
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth1       =   600
         FldWidth2       =   1980
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilCompCode 
         Height          =   285
         Left            =   1575
         TabIndex        =   0
         Top             =   285
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         FldWidth1       =   600
         FldWidth2       =   2500
         F5Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilPTCode 
         Height          =   300
         Left            =   7455
         TabIndex        =   6
         Top             =   5265
         Width           =   2910
         _ExtentX        =   5133
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
         FldWidth1       =   600
         FldWidth2       =   1980
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilPTCodePay 
         Height          =   285
         Left            =   1860
         TabIndex        =   8
         Top             =   5310
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   503
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth1       =   600
         FldWidth2       =   1980
         F2Corresponding =   -1  'True
      End
      Begin prjGiList.GiList gilClctEn 
         Height          =   285
         Left            =   7440
         TabIndex        =   5
         Top             =   4890
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FldWidth1       =   1000
         FldWidth2       =   2000
         F2Corresponding =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "���O�H��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   6495
         TabIndex        =   32
         Top             =   4950
         Width           =   810
      End
      Begin VB.Label lblPtCodePay 
         AutoSize        =   -1  'True
         Caption         =   "���O�I�ں���"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   630
         TabIndex        =   28
         Top             =   5385
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�h�O�I�ں���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6180
         TabIndex        =   26
         Top             =   5340
         Width           =   1170
      End
      Begin VB.Label lblReasonCode 
         AutoSize        =   -1  'True
         Caption         =   "���h�O�ζ���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6210
         TabIndex        =   25
         Top             =   5715
         Width           =   1170
      End
      Begin VB.Label lblFRAmount 
         Alignment       =   1  '�a�k���
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         Enabled         =   0   'False
         Height          =   270
         Left            =   1875
         TabIndex        =   24
         Top             =   6105
         Width           =   1650
      End
      Begin VB.Label lblRAmount 
         Caption         =   "�������B"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1005
         TabIndex        =   23
         Top             =   6135
         Width           =   795
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1020
         TabIndex        =   22
         Top             =   5745
         Width           =   780
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   270
         X2              =   10830
         Y1              =   6495
         Y2              =   6495
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   240
         X2              =   10830
         Y1              =   4725
         Y2              =   4725
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "��  �q  �O"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   750
         TabIndex        =   21
         Top             =   330
         Width           =   765
      End
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
      Left            =   8850
      TabIndex        =   11
      Top             =   6510
      Width           =   1410
   End
End
Attribute VB_Name = "frmSO1147A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//
'// �u�沣��(�����^/�榬�^)/�W�ߧ@�~(����P/�榬�^)
Dim WithEvents rsSO004   As ADODB.Recordset
Attribute rsSO004.VB_VarHelpID = -1
Dim rsTmp As New ADODB.Recordset
'//
''Public gServiceType As String   '' �~���ǤJ���A�ȧO
''Public strCustid As String   '' �~���ǤJ���Ȥ�s��
'' �~���ǤJ���M�@�~����
Public intType As Integer '' 0. STB ���M  1.����P   2.�P��h�^
'//
Public pClctEn As String    '' �~���ǤJ �u�{�H���N�X (intType �O 0 ���~�ݭn�A�_�h���ާ@�̥N��)
Public pClctName As String  '' �~���ǤJ �u�{�H���N�X (intType �O 0 ���~�ݭn�A�_�h���ާ@�̦W��)
Public pBillNO As String   '' �~���ǤJ���渹(intType �O 0 ���~�ݭn�A�_�h�ۦ�զ��{�ɳ渹)
Public pItem As String   '' �~���ǤJ���]�ƶ���(intType �O 0 ���~�ݭn�A�_�h�T�w�� 1�A�p�G���ͪ��O���O��ƫh�� 2   )
Public pSeqNo As String
Private blnFailFlag As Boolean
Private blnMSS As Boolean
Public blnFromCall As Boolean   ''  �ˬd�O�_�O�ѥ~���I�s

Private lngContractOldAmount As Long
Private lngContractShouldAmount As Long
Private blnContractCustomer As Boolean

Private Sub cmdCancel_Click()
      Unload Me
End Sub

Private Sub btnExcel_Click()
'//
Dim lngReturn As Long
Dim intMonthRange As Integer
Dim rs As New ADODB.Recordset
Dim blnAppDaysRange As Boolean  '' �O���O�_�W�LŲ��� �A�O���ܳ]�� true  �_�h�]�� false

'//
On Error GoTo ChkErr

      Call SetDefaultClct
     
     If IsDataOk(False) = False Then Exit Sub
     
     If rsSO004.RecordCount = 0 Then
           MsgBox "�S�����Ȥ᪺�����]�Ƹ�ơA���\�ඵ�صL�k�ϥ�  !!", _
                        vbOKOnly + vbInformation, "�T�� "
           Exit Sub
     End If

    '// intMonthRange = DateDiff("m", rsSo004("InstDate"), gdCancelDate.GetValue(True))
     If CDate(gdCancelDate.GetValue(True)) < CDate(rsSO004("InstDate")) Then
          MsgBox "���M�餣�i�H�p��I��� !!", vbOKOnly + vbInformation, "�T���@"
          Exit Sub
     End If
     
     ''intMonthRange = DateDiff("m", CDate(gdCancelDate.GetValue(True)), rsSO004("InstDate"))
     intMonthRange = DateDiff("m", rsSO004("InstDate"), CDate(gdCancelDate.GetValue(True)))
     If CDate(gdCancelDate.GetValue(True)) >= DateAdd("m", intMonthRange, rsSO004("InstDate")) Then intMonthRange = intMonthRange + 1
  
     rs.CursorLocation = adUseClient
     rs.Open _
          "SELECT * FROM " & GetOwner & "SO046 WHERE COMPCODE  = " & gilCompCode.GetCodeNo, _
          gcnGi, adOpenKeyset, adLockReadOnly
     
     If rs.RecordCount = 0 Then
           MsgBox "�S���]�wSTB �h�O�ѼƳ]�w������ơA���\�ඵ�صL�k�ϥ�  !!", _
                        vbOKOnly + vbInformation, "�T�� "
                        rs.Close
                        Set rs = Nothing
           Exit Sub
      End If
     
     Select Case rsSO004("BuyCode")
        Case 1  ' �p�G�O�R�i�h�^��
 '''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""''''
 ''  2002/12/23  -- �վ�{�����e--
 ''  �H �U�o�@�q��Ų���������
 ''  ���p�O�R�h�� ( BuyCode = 1 ), �˵�  Ų���
 ''-------------------------------------------------------------------------------------------------------------
     Dim rsSO043 As New ADODB.Recordset
     Dim intAppreciateDays As Integer
     Dim blnCancel As Boolean
     
     blnCancel = False
     blnAppDaysRange = False
     
     rsSO043.CursorLocation = adUseClient
     rsSO043.Open "SELECT AppreciateDays FROM " & GetOwner & "SO043", gcnGi, adOpenKeyset, adLockReadOnly
     intAppreciateDays = IIf(IsNull(rsSO043("AppreciateDays")), 0, rsSO043("AppreciateDays"))
     
     If intAppreciateDays <> 0 Then
            If rsSO004("BUYCODE") = 1 Then
                If DateDiff("d", CDate(rsSO004("InstDate")), CDate(gdCancelDate.GetValue(True))) > intAppreciateDays Then
                        MsgBox "�ثe�]�Ƹ˸m�w�g�W�LŲ����A�H�h�O����p��h�O��� !!  ", vbOKOnly + vbInformation, "�T��"
                        SetgiList gilReturnItme, "CodeNo", "Description", "CD019", , , , , , , " Where   periodflag  =  0 AND REFNO = 7  "
                        blnAppDaysRange = True
                Else
                        MsgBox "�ثe�]�Ƹ˸m�bŲ������A�H���B�h�O !!  ", vbOKOnly + vbInformation, "�T��"
                        lngReturn = rsSO004("Amount")
                        SetgiList gilReturnItme, "CodeNo", "Description", "CD019", , , , , , , " Where   periodflag  =  0 AND REFNO = 13  "
                End If
                If gilReturnItme.GetListCount <= 0 Then
                        gilReturnItme.Filter = " where periodflag  = 0 AND SIGN = '-' "
                End If
            End If
     Else
            MsgBox "Ų�����ƥ��]�A�H�U�������Ͱh�O��ơA���@�P�O !!", vbOKOnly + vbInformation
            blnAppDaysRange = True
     End If
     rsSO043.Close
     Set rsSO043 = Nothing
     If blnCancel = True Then
            Exit Sub
     End If

 '''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
      If blnAppDaysRange = True Then
                 If blnMSS = False Then
'                If chkMSS.Value = 0 Then
                         ''  �H�U�� SO046  �����x�]�w
                        If intMonthRange >= rs("STARTM1") And intMonthRange <= rs("STOPM1") Then
                            lngReturn = rs("PRICEM1")
                        ElseIf intMonthRange > rs("STARTM2") And intMonthRange <= rs("STOPM2") Then
                            lngReturn = rs("PRICEM2")
                        ElseIf intMonthRange > rs("STARTM3") And intMonthRange <= rs("STOPM3") Then
                            lngReturn = rs("PRICEM3")
                        ElseIf intMonthRange > rs("STARTM4") And intMonthRange <= rs("STOPM4") Then
                            lngReturn = rs("PRICEM4")
                        ElseIf intMonthRange > rs("STARTM5") And intMonthRange <= rs("STOPM5") Then
                            lngReturn = rs("PRICEM5")
                        End If
                Else
                      Dim O As Object
                      Dim strMeessage As String
                      Dim blnSucc As Boolean
                      Dim strDarte As String
                      strDarte = "0" & Replace(gdCancelDate.GetOriginalValue, "/", "")
                      blnSucc = False
                  ''    blnSucc = MSSReturn(O, 1, rsSO004("SmartCardNo"), strDarte, lngReturn, strMeessage)
                      Pic1.Visible = True
                      blnSucc = MSSReturn(O, 1, rsSO004("FaciSNO"), strDarte, lngReturn, strMeessage)
                      If blnSucc = False Then
                           Call MsgBox(strMeessage, vbOKOnly + vbCritical, "�T�� ")
                           Pic1.Visible = False
                           Exit Sub
                      End If
                       Pic1.Visible = False
                      
                End If
        End If
        Case 3 '' �p�G�O�ӯ��h�^��'
             ''  2003/11/12  ���p�O�j����A�h�����պ⤧��A�i��t�~�@�Ӭy�{
               If intMonthRange <= rs("MonthCount") Then
                            lngReturn = Val(rsSO004("Deposit") & "")
                        ElseIf intMonthRange > rs("MONTHCOUNT") Then
                            lngReturn = Val(rsSO004("Deposit") & "") - (intMonthRange - rs("MONTHCOUNT")) * rs("BACKPRICE")
              End If
       
            
        Case 4 '' �p�G�O�ɰh�^��'
                lngReturn = Val(rsSO004("Deposit") & "")
    End Select
    
'' -----------------------------------------------------------------------------------------------------------------------------------
   lngReturn = lngReturn - lblAmount.Caption
    lblCheckAmount.Caption = IIf(lngReturn > 0, lngReturn, 0)
    
''20031117 �o�@�q�p��j����n�h�����B
    
    If rsSO004("BuyCode") = 3 And rsSO004("ContractCust") = 1 Then
                        Dim rsSO7300H As New ADODB.Recordset
                        Dim lngMonth As Long  '' �j�����
                        Dim lngMonthBackPrice As Long  '' ��h���B
                        Dim lngGiveMonthAmount As Long
                        blnContractCustomer = True
                        Call GetRS(rsSO7300H, "SELECT * FROM SO046 WHERE CompCode=" & gilCompCode.GetCodeNo, gcnGi, adUseClient, adOpenKeyset)
                        lngMonth = rsSO7300H("ContractMonth")
                        lngMonthBackPrice = rsSO7300H("BackPrice")
                        If DateDiff("D", CDate(rsSO004("InstDate")), Date) + 1 < lngMonth * 30 Then  '' �H����
                                        lngGiveMonthAmount = (DateDiff("D", CDate(rsSO004("InstDate")), Date) + 1) * _
                                                                              (lngMonthBackPrice / 30)
                                        If lngGiveMonthAmount > lngReturn Then  ''�w�ذe���B�j��O����
                                            lngContractOldAmount = lngGiveMonthAmount
                                            lngContractShouldAmount = lngReturn
                                            MsgBox "�ӳ]�Ƭ��j����A�]�H��" & vbCrLf & _
                                                           "�w�ذe���믲�O�ε���" & lngGiveMonthAmount & "�A�w�g�W�L�O�������h���B �A" & vbCrLf & _
                                                          "�t�αN�|���ͤ@�������O�ε���O�������h���B " & lngReturn & "����� !!", vbInformation, "���M�T��"
                                        Else
                                            lngContractOldAmount = lngGiveMonthAmount
                                            lngContractShouldAmount = lngGiveMonthAmount
                                            MsgBox "�ӳ]�Ƭ��j����A�]�H��" & vbCrLf & _
                                                          "�w�ذe���믲�O�ε���" & lngGiveMonthAmount & "�A���W�L�O�������h���B  �A" & vbCrLf & _
                                                          "�t�αN�|���ͤ@�������O�ε���w�ذe���믲�O" & lngGiveMonthAmount & "����� !!", vbInformation, "���M�T��"
                                        End If
                        Else
                                        blnContractCustomer = False
                                        lngContractOldAmount = 0
                                        lngContractShouldAmount = 0
                        End If
                         rsSO7300H.Close
                         Set rsSO7300H = Nothing
    Else
                        blnContractCustomer = False
                        lngContractOldAmount = 0
                        lngContractShouldAmount = 0
    End If
    rs.Close
    Set rs = Nothing
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "btnExcel_Click")
End Sub
Private Sub cmdReciveAmount_Click()

    Dim strMessage As String
On Error GoTo ChkErr
    If IsDataOk(True) = False Then Exit Sub

  '//  ���p�պ���B�������ܡA�N���Ͱh�O���
  
' '''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""''''
' ''  2002/12/23  -- �վ�{�����e--
' ''  �H �U�o�@�q��Ų���������
' ''-------------------------------------------------------------------------------------------------------------
' If intType = 2 Then  '' ���p�O�R�h�� , �˵�  Ų���
'     Dim rsSO043 As New ADODB.Recordset
'     Dim intAppreciateDays As Integer
'     Dim blnCancel As Boolean
'
'     blnCancel = False
'     rsSO043.CursorLocation = adUseClient
'     rsSO043.Open "SELECT AppreciateDays FROM SO043", gcnGi, adOpenKeyset, adLockReadOnly
'     intAppreciateDays = IIf(IsNull(rsSO043("AppreciateDays")), 0, rsSO043("AppreciateDays"))
'     If intAppreciateDays <> 0 Then
'            If rsSO004("BUYCODE") = 1 Then
'                If DateDiff("d", CDate(rsSO004("InstDate")), CDate(gdCancelDate.GetValue(True))) > intAppreciateDays Then
'                       If MsgBox("�ثe�]�Ƹ˸m�w�g�W�LŲ����A�T�w�j��h�O��  ??  ", vbYesNo + vbInformation + vbDefaultButton2, "�T��") = vbNo Then
'                              MsgBox "�z�����F�h�O����A�t�αN���|���Ͱh�O��� !!"
'                              blnCancel = True
'                        Else
'                              If MsgBox("�z��ܱj��h�O�A�N�~�򲣥Ͱh�O��� !!", _
'                                  vbYesNo + vbInformation, "�T�� ") = vbNo Then
'                                       blnCancel = True
'                              End If
'                       End If
'                End If
'            End If
'     Else
'           MsgBox "Ų�����ƥ��]�A�H�U�������Ͱh�O��ơA���@�P�O !!", vbOKOnly + vbInformation
'     End If
'     rsSO043.Close
'     Set rsSO043 = Nothing
'     If blnCancel = True Then
'            Exit Sub
'     End If
' End If
' '''""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
    If intType > 0 And blnFromCall = False Then gcnGi.BeginTrans
''200311117   �p�G�O�j����A���ͤ@�������@�������O����

    If blnContractCustomer = True Then
        Call GenerateSO033Data(4)
        ' ���ͤ@�����O��ƨñN��Բ�J SO033
        If InsertIntoSO033 = False Then  '
            gcnGi.RollbackTrans
            MsgBox "�����j����A���͸j���᦬�O��Ƹ�ƥ��� !!", vbOKOnly + vbInformation, "�T��  "
            Exit Sub
         Else
            MsgBox "�����j����A���\���ͤ@���j���᦬�O��� !!", vbOKOnly + vbInformation, "�T��  "
        End If
    End If

    '#1653 �p�GUser���T��,�������ͤ@�����O���,���A�^�ǵ�SO111XC By Kin 2007/06/04
    If lblCheckAmount.Caption > 0 Then
        Call GenerateSO033Data(intType)
        If intType = 0 And optP.Value Then
            Unload Me
            Exit Sub
        End If
        ' ���ͤ@���h�O�O��ƨñN��Բ�J SO033
        If InsertIntoSO033 = False Then  '
            gcnGi.RollbackTrans
            MsgBox "���Ͱh�O��Ƹ�ƥ��� !!", vbOKOnly + vbInformation, "�T��  "
            Exit Sub
        End If
    Else
        MsgBox "���h���B�� 0 �A�����Ͱh�O��� !!"
    End If
   
   '//  ���p�O����P�h�A���ͤ@�����O��T
   
    If intType = 1 Then
        Call GenerateSO033Data(3)
        ' ���ͤ@�����O��ƨñN��Բ�J SO033
        If InsertIntoSO033 = False Then  '
            gcnGi.RollbackTrans
            MsgBox "���ͦ��O��Ƹ�ƥ��� !!", vbOKOnly + vbInformation, "�T��  "
            Exit Sub
        End If
    End If
   
 
   
 '//  �I�s�F�ƼҲ� , �n�����ʸ�T�H�Χ��]�ƪ��A
   
    Select Case intType
        Case 1  '''����P
    '         If CallRentToSale(gilCompCode.GetCodeNo, rsTmp("BILLNO"), _
    '                                     rsSO004("FaciSNo"), rsSO004("SmartCardNo"), rsSO004("SEQNo")) = False Then
              
             If CallRentToSale(gilCompCode.GetCodeNo, rsTmp("BILLNO"), _
                                         rsSO004("FaciSNo") & "", rsSO004("SmartCardNo") & "", rsSO004("SEQNo"), rsTmp("ShouldAmt")) = False Then
              
                gcnGi.RollbackTrans
                Exit Sub
            End If
             strMessage = " ���\���ͤ@���h�O��ƥH�Τ@�����O��� !! "
        Case 2  '' �P��h�^
            If CallSaleReturn _
                         (gilCompCode.GetCodeNo, rsTmp("BILLNO"), rsSO004("FaciSNo"), _
                          rsSO004("SmartCardNo") & "", gdCancelDate.GetValue, rsTmp.Fields("OldAmt"), _
                          pItem, rsSO004("SEQNo"), rsSO004("CODENO") & "", rsSO004("DESCRIPTION") & "") = False Then
                frmSO1147A.Pic1.Visible = False
                gcnGi.RollbackTrans
                blnFailFlag = True
                Exit Sub
             Else
                blnFailFlag = False
             End If
             strMessage = " ���\���ͤ@���h�O��� !! "
    End Select
    '#1653 �p�G�O�q�u��I�s,�ӥB�S�O���T��,�h��Show�T��,�������^�u��
    If intType = 0 Then
        Unload Me
        Exit Sub
    End If
    If blnFromCall = False Then gcnGi.CommitTrans
    Call MsgBox(strMessage, vbOKOnly + vbInformation, "�T���@")
    Unload Me
 Exit Sub
    
ChkErr:
    gcnGi.RollbackTrans
    Call ErrSub(Me.Name, "cmdReciveAmount_Click")
 
End Sub

Private Sub Form_Load()

On Error GoTo ChkErr
     Dim blnOpen As Boolean
     
     Call SubGil
     Call subGim
     
     Set rsSO004 = New ADODB.Recordset
     blnOpen = OpenData
     If blnOpen = False Then
         frmSO1100BMDI.mnuSO11471.Tag = "1"
         MsgBox "���Ȥ�L�����]�Ƹ�� !!", vbOKOnly + vbInformation, "�T��"
         Unload Me
         Exit Sub
     End If
    
    blnFailFlag = False
    'gdCancelDate.Text = Format(DateAdd("d", 1, Date), "ee/mm/dd")
    gdCancelDate.Text = GetDTString(DateAdd("d", 1, Date))
    '#1653 ���]�����P�Ҧ��A�Щ�W�[������O���ﶵ�A��User��ܰh�O���جO���bT��Τu�� By Kin 2007/06/04
    Select Case intType
        Case 0
            Me.Caption = "STB ���M�@�~"
            optP.Enabled = True
        Case 1
            Me.Caption = "����P�@�~"
            chkRent.Value = 1
            lblPtCodePay.Enabled = True
            gilPTCodePay.Enabled = True
            gilPTCodePay.ListIndex = 1
            lblItem.Enabled = True
            gilItem.Enabled = True
            gilItem.ListIndex = 1
            lblRAmount.Enabled = True
            lblFRAmount.Enabled = True
        Case 2
            Me.Caption = "�P�h�^�@�~"
    End Select
    Call subGrd
    blnMSS = IIf(RPxx(gcnGi.Execute("SELECT LinkMss FROM  " & GetOwner & "CD039 WHERE  CodeNo =" & gilCompCode.GetCodeNo).GetString) = 1, True, False)
    Exit Sub
    
    lngContractOldAmount = 0
    lngContractShouldAmount = 0
    blnContractCustomer = False
    
ChkErr:
    Call ErrSub(Me.Name, "Form_Load")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
   If intType <> 0 And blnFromCall = False Then CloseRecordset rsSO004
   If intType <> 0 And blnFromCall = False Then CloseRecordset rsTmp
   '#1653 ���u��@���Ū��O��,����u��A���@��Update By Kin 2007/06/04
   If intType = 0 And optT Then
        CloseRecordset rsSO004
        CloseRecordset rsTmp
        Set rsTmp = gcnGi.Execute("Select * From " & GetOwner & "SO033 Where 1=0")
   End If
   ReleaseCOM Me
   blnFromCall = False
'    rsSO004.Close
'    If rsTmp.State = 1 And intType <> 0 Then Set rsTmp = Nothing
'    Set rsSO004 = Nothing
End Sub

Private Sub gilCompCode_Change()
   garyGi(6) = gilCompCode.GetOwner
   gimStbItem.Filter = " WHERE COMPCODE =" & gilCompCode.GetCodeNo
End Sub
Private Sub gilItem_Change()
    Dim intCodeno As Integer
    Dim rs As ADODB.Recordset
    If Len(gilItem.GetCodeNo) > 0 Then
        intCodeno = gilItem.GetCodeNo
        Set rs = New ADODB.Recordset
        rs.Open "SELECT AMOUNT FROM " & GetOwner & "CD019 WHERE CODENO  = " & gilItem.GetCodeNo, gcnGi, adOpenKeyset, adLockReadOnly
        lblFRAmount.Caption = rs(0)
          rs.Close
         Set rs = Nothing
    End If
  
End Sub

Private Sub gimStbItem_Change()

Dim strFilter As String
Dim strSQL As String
Dim lngCost As Long
Dim rs As New ADODB.Recordset

On Error GoTo ChkErr

If gimStbItem.GetQryStr <> "" Or gimStbItem.GetQueryCode <> "" Then

        If gimStbItem.GetQryStr = "" Then
            strFilter = ""
        Else
           strFilter = "  AND  CD028.CODENO  " & gimStbItem.GetQryStr
        End If
        
        rs.CursorLocation = adUseClient
        strSQL = "SELECT SUM(COST) FROM " & GetOwner & "CD028 WHERE  COMPCODE  = " & gilCompCode.GetCodeNo & strFilter
        rs.Open strSQL, gcnGi, adOpenKeyset, adLockReadOnly
        If rs.RecordCount > 0 Then
            lblAmount.Caption = rs(0)
        Else
            lblAmount.Caption = 0
        End If
        rs.Close
        Set rs = Nothing
Else
        lblAmount.Caption = 0
End If

Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "gimStbItem_Change")
End Sub
Private Sub lblCheckAmount_Change()

If Len(lblCheckAmount.Caption) > 0 Then
   If lblCheckAmount.Caption > 0 Then
       cmdReciveAmount.Enabled = True
   Else
       cmdReciveAmount.Enabled = False
   End If
Else
   cmdReciveAmount.Enabled = False
End If

End Sub
Private Sub SubGil()
On Error GoTo ChkErr

    SetgiList gilItem, "CodeNo", "Description", "CD019", , , , , , , " Where periodflag  = 0 AND REFNO = 8  AND StopFlag  = 0  "
    
    If gilItem.GetListCount <= 0 Then
       gilItem.Filter = " where periodflag  = 0 AND SIGN = '+' "
    End If
    
    SetgiList gilPTCode, "CodeNo", "Description", "CD032", , , , , , , " Where NVL(SERVICETYPE,'" & gServiceType & "') = '" & gServiceType & "'  AND StopFlag  = 0  "
    gilPTCode.ListIndex = 1
    
    SetgiList gilPTCodePay, "CodeNo", "Description", "CD032", , , , , , , " Where   NVL(SERVICETYPE,'" & gServiceType & "') = '" & gServiceType & "' AND StopFlag  = 0  "
    
    SetgiList gilReturnItme, "CodeNo", "Description", "CD019", , , , , , , " Where   periodflag  =  0 AND REFNO = 7  AND StopFlag  = 0  "
    
    If gilReturnItme.GetListCount <= 0 Then
       gilReturnItme.Filter = " where periodflag  = 0 AND SIGN = '-' "
    End If
    
    SetgiList gilClctEn, "EmpNo", "EmpName", "CM003", , , , , 10, 20, , True
    
    SetgiList gilCompCode, "CodeNo", "Description", "CD039", , , , , 3, 12, GetCompCodeFilter(gilCompCode)
    gilCompCode.SetCodeNo gCompCode
    gilCompCode.Query_Description
    
  Exit Sub
ChkErr:
  ErrSub Me.Name, "subGil"
End Sub
Private Sub subGim()
    Dim strFilter As String
  
    strFilter = " Where COMPCODE  = " & gilCompCode.GetCodeNo
    SetgiMulti gimStbItem, "CodeNo", "Description", "CD028", "�N�X", "�W��", strFilter, True
End Sub

Private Function IsDataOk(ByVal blnReciveAmount As Boolean) As Boolean

On Error GoTo ChkErr

IsDataOk = False
If blnReciveAmount = True Then

   If Len(Trim(gilPTCode.Text)) = 0 Then
        MsgBox "�z���ӿ���@���h�O�I�ں������� !!", vbInformation + vbOKOnly, "�T�� "
        gilPTCode.SetFocus
        Exit Function
     End If
     
    If Len(Trim(gilReturnItme.Text)) = 0 Then
        MsgBox "�z���ӿ���@�����h�O�ζ��� !!", vbInformation + vbOKOnly, "�T�� "
        gilReturnItme.SetFocus
        Exit Function
     End If
     
     If chkRent.Value = 1 Then
         
        If Len(Trim(gilPTCodePay.Text)) = 0 Then
               MsgBox "�z���ӿ���@�����O�I�ں������� !!", vbInformation + vbOKOnly, "�T�� "
                gilPTCodePay.SetFocus
                Exit Function
         End If
     
          If Len(Trim(gilItem.Text)) = 0 Then
                 MsgBox "�z���ӿ���@���������� !!", vbInformation + vbOKOnly, "�T�� "
                 gilItem.SetFocus
                 Exit Function
          End If
     End If
End If

     If Not IsDate(gdCancelDate.Text) = True Then
          MsgBox "�нT�{���M�� !!  ", vbInformation + vbOKOnly, "�T��"
          gdCancelDate.SetFocus
          Exit Function
     End If
     
     If Len(gilClctEn.Text) = 0 Then
          MsgBox "����w���O�H�� !!  ", vbInformation + vbOKOnly, "�T��"
          gilClctEn.SetFocus
          Exit Function
     End If
     
     IsDataOk = True
     Exit Function
ChkErr:
    Call ErrSub(Me.Name, "subGrd")
End Function

Private Sub subGrd()

Dim mFlds As New prjGiGridR.GiGridFlds

On Error GoTo ChkErr

        With mFlds
            .Add "FaciSNo", , , , False, "        �Ǹ�             ", vbLeftJustify
            .Add "FaciName", , , , False, "      �~�W        ", vbLeftJustify
            .Add "SmartCardNo", , , , False, "      ���z�d��        ", vbLeftJustify
            .Add "BuyName", , , , False, "�R��覡", vbLeftJustify
            .Add "InstDate", giControlTypeDate, , , False, "�w�ˤ��       ", vbLeftJustify
         ''   .Add "InitPlaceNo", , , , False, "��m              ", vbLeftJustify
            .Add "DESCRIPTION", , , , False, "��m              ", vbLeftJustify
            .Add "Deposit", , , , False, "�O����", vbRightJustify
            .Add "PRDate", giControlTypeDate, , , False, "���     ", vbLeftJustify
          '''  .Add "Note", , , , False, Space(20) & "         �Ƶ�            ", vbLeftJustify
        End With
        
        ggrData.AllFields = mFlds
        ggrData.SetHead
        Set ggrData.Recordset = rsSO004
        ggrData.Refresh
        ggrData.Visible = True
      
            
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "subGrd")
End Sub

Private Function OpenData() As Boolean
    Dim intBuyCode As Integer
    Dim strBuySql As String
    Dim strSQL As String
On Error GoTo ChkErr
    Select Case intType
        Case 0
            strSQL = "SELECT SO004.* , CD056.CODENO , CD056.DESCRIPTION  FROM  " & GetOwner & "SO004 ," & GetOwner & "CD056 " & _
                        "where " & _
                        "Seqno  = '" & pSeqNo & _
                        "' AND  CD056.CODENO(+) = SO004.InitPlaceNo AND SO004.PRDATE IS NULL  " & _
                        "AND InstDate IS NOT NULL "
        Case Else
            If intType = 1 Then
                strBuySql = " AND SO004.BUYCODE  =3 "
            Else
                strBuySql = " AND (SO004.BUYCODE  =1 OR SO004.BUYCODE  =3 OR SO004.BUYCODE = 4  OR SO004.BUYCODE = 5 )  "
            End If
            strSQL = "SELECT SO004.* , CD056.CODENO ,CD056.DESCRIPTION  FROM  " & GetOwner & "SO004 , " & GetOwner & "CD022 , " & GetOwner & "CD056 WHERE  " & _
                            "SO004.CUSTID = " & gCustId & "  AND  " & _
                            "SO004.FACICODE = CD022.CODENO  AND  " & _
                            "CD022 .REFNO   = 3  AND CompCode =" & _
                            gCompCode & _
                            " AND CD056.CODENO(+)= SO004.InitPlaceNo  AND SO004.PRDATE IS NULL  " & _
                            "AND InstDate IS NOT NULL "
            strSQL = strSQL & strBuySql
            
    End Select
    If rsSO004.State = 1 Then rsSO004.Close
    Set rsSO004 = New ADODB.Recordset
    rsSO004.CursorLocation = adUseClient
    rsSO004.Open strSQL, gcnGi
    If rsSO004.RecordCount = 0 Then
       OpenData = False
    Else
       OpenData = True
    End If
    
Exit Function
ChkErr:
    Call ErrSub(Me.Name, "OpenData")
End Function

Private Sub GenerateSO033Data(pintType As Integer)

'// �H�U���ثe�L�k�F��
'// ���q�O�O�_�H�W�h�ǨӬ��D
'// �A�ȧO�u��  "C" ��

    Dim strSQL As String
    Dim rsUC As New ADODB.Recordset  '// ���o������]�N�X
    Dim rsCust As New ADODB.Recordset  '// ���o�P�a�}�������������
    
    Dim strOldShouldAmt As String    '' �������B & �X����B
    Dim strShouldAmt As String   ''
    
On Error GoTo ChkErr

'' �̾ڤ��P��pintType ���o�U���h�O��ƪ�������줺�e
'#1653 �p�G�O�q���u��Ǩ�,�ӥBUser�O���T�檺��,�h�������ͤ@��T��,���ζǨӪ����u�沣�� By Kin 2007/06/04
    Select Case pintType
        Case 0  '' STB ���M�@�~
            strOldShouldAmt = 0 - lblCheckAmount.Caption
            If optT Then
                pBillNO = GetInvoiceNo("T", gServiceType)
                pItem = 1
                pClctEn = gilClctEn.GetCodeNo
                pClctName = gilClctEn.GetDescription
            End If
        Case 1  ''"����P�@�~
            pBillNO = GetInvoiceNo("T", gServiceType)
            pItem = 1
            strOldShouldAmt = 0 - lblCheckAmount.Caption
            pClctEn = gilClctEn.GetCodeNo
            pClctName = gilClctEn.GetDescription
        Case 2  ''"�P�h�^�@�~"
            If blnFailFlag = False Then
                pBillNO = GetInvoiceNo("T", gServiceType)
            End If
            pItem = 1
            strOldShouldAmt = 0 - lblCheckAmount.Caption
            pClctEn = gilClctEn.GetCodeNo
            pClctName = gilClctEn.GetDescription
        Case 3  '//  �o�@��������P�ɡA�ΥH���ͤ@�����O����
         '' pBillNO = GetInvoiceNo("T", gServiceType)
            pItem = 2
            strOldShouldAmt = lblFRAmount.Caption
            pClctEn = gilClctEn.GetCodeNo
            pClctName = gilClctEn.GetDescription
        Case 4  '//  �o�@�����j����  �t�~���ͪ��@�����O����
            pClctEn = gilClctEn.GetCodeNo
            pClctName = gilClctEn.GetDescription
    End Select

    If lblCheckAmount.Caption > 0 Then
 
 '// �H�U�ΥH���o������]�N�X
        strSQL = "SELECT CODENO , DESCRIPTION FROM " & GetOwner & "CD013  WHERE  " & _
                          "NVL(SERVICETYPE,'" & gServiceType & "') = '" & gServiceType & "' AND REFNO  =1 "
        
        rsUC.CursorLocation = adUseClient
        rsUC.Open strSQL, gcnGi, adOpenKeyset, adLockReadOnly
        
        strSQL = "SELECT  " & _
                       "SO001.INSTADDRNO,SO001.SERVCode , SO001.ClctAreaCode ,SO001.ClassCode1," & _
                        "SO014.STRTCODE ,SO014.MDUID , SO014.NODENO ,SO014.AREACODE  " & _
                        "FROM " & _
                        GetOwner & "SO001," & GetOwner & "SO014 " & _
                        "WHERE " & _
                        "SO001.InstAddrNo  = SO014.ADDRNO  AND SO001.CUSTID   = " & gCustId
                    
        rsCust.CursorLocation = adUseClient
        rsCust.Open strSQL, gcnGi, adOpenKeyset, adLockReadOnly
      
    ''  �إ߼Ȧs��
        Dim strSELECT  As String
    
        strSELECT = "select " & _
                             "Custid,BillNo,Item," & _
                             "CitemCode,CitemName," & _
                             "OldAmt,ShouldDate," & _
                             "ShouldAmt,ClctEn," & _
                             "ClctName,CMCode," & _
                             "CMName,UCCode," & _
                             "UCName,Note,CreateTime,CreateEn,CompCode,AddrNo,ServCode," & _
                             "ClctAreaCode,NodeNo,PTCode,PTName,OldClctEn," & _
                             "OldClctName,ClassCode,UpdTime,UpdEn,ServiceType,Quantity ," & _
                             "CancelFlag,AreaCode,StrtCode,MduId,FaciSeqNo ,FaciSNo   " & _
                             "from " & GetOwner & "so033 where 1=0 "
                            
        If Not GetRS(rsTmp, strSELECT, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
        Set rsTmp.ActiveConnection = Nothing
        With rsTmp
            .AddNew
           
            .Fields("CustID") = gCustId '' �Ȥ�s�� '0
            .Fields("BillNo") = pBillNO '1
            
            If pintType = 4 Then
                .Fields("Item") = 4
            Else
                .Fields("Item") = pItem  '2
            End If
            If pintType = 3 Then
                .Fields("CitemCode") = gilItem.GetCodeNo '3
                .Fields("CitemName") = gilItem.GetDescription
            Else
                If pintType = 4 Then
                    .Fields("CitemCode") = RPxx(gcnGi.Execute("SELECT CODENO FROM CD019 WHERE REFNO = 14 And StopFlag = 0").GetString)
                    .Fields("CitemName") = RPxx(gcnGi.Execute("SELECT Description FROM CD019 WHERE REFNO = 14 And StopFlag = 0").GetString)
                Else
                    .Fields("CitemCode") = gilReturnItme.GetCodeNo '3
                    .Fields("CitemName") = gilReturnItme.GetDescription
                End If
            End If
            
            If pintType <> 4 Then
                .Fields("OldAmt") = strOldShouldAmt
            Else
               .Fields("OldAmt") = lngContractOldAmount
            End If
            .Fields("ShouldDate") = gdCancelDate.GetValue(True)
            
            If pintType <> 4 Then
                .Fields("ShouldAmt") = strOldShouldAmt
            Else
                .Fields("ShouldAmt") = lngContractShouldAmount
            End If
    
            .Fields("ClctEn") = pClctEn '8
            .Fields("ClctName") = pClctName '9
            .Fields("CMCode") = 1 '10
            .Fields("CMName") = GetRsValue("Select Description From " & GetOwner & "CD031 Where CodeNo = 1") & "" '11
           ' .Fields("CMName") = "�H�u" '11
            .Fields("UCCode") = rsUC(0) '12
            .Fields("UCName") = rsUC(1) '13
             Dim strReturn As String
            If pintType <> 4 Then
                 If CInt(strOldShouldAmt) > 0 Then strReturn = "��" Else strReturn = "�h"
                .Fields("Note") = "���W�� (�Ǹ� = " & rsSO004("FaciSNo") & ") �����M��" & strReturn & "�O�� " '14
            Else
                .Fields("Note") = "���W�� (�Ǹ� = " & rsSO004("FaciSNo") & ") ���M�j���������O�� " '14
            End If
            .Fields("CreateTime") = Format(Now, "YYYY/MM/DD HH:MM:SS") '15
            .Fields("CreateEn") = garyGi(0) '16
            .Fields("CompCode") = gilCompCode.GetCodeNo '17
            
            .Fields("AddrNo") = rsCust("INSTADDRNO") '18
            .Fields("ServCode") = rsCust("ServCode") '19
            .Fields("ClctAreaCode") = rsCust("ClctAreaCode") '20
            .Fields("NodeNo") = rsCust("NodeNo") '21
            
            If pintType = 3 Then
                 .Fields("PTCode") = gilPTCodePay.GetCodeNo
                 .Fields("PTName") = gilPTCodePay.GetDescription '23
            Else
                 .Fields("PTCode") = gilPTCode.GetCodeNo '22
                 .Fields("PTName") = gilPTCode.GetDescription '23
            End If
         
            .Fields("OldClctEn") = pClctEn '24
            .Fields("OldClctName") = pClctName '25
            
            .Fields("ClassCode") = rsCust("ClassCode1") '26
            .Fields("UpdTime") = GetDTString(Now) '27
            .Fields("UpdEn") = garyGi(1) '28
            .Fields("ServiceType") = gServiceType  '29
            
            .Fields("Quantity") = 0  ''30
            .Fields("CancelFlag") = 0   ''31
            
            .Fields("AreaCode") = rsCust("AreaCode")  ''32
            .Fields("StrtCode") = rsCust("StrtCode")  ''33
            .Fields("MduId") = rsCust("MduId")  ''34
            .Fields("FaciSeqNo") = NoZero(rsSO004("SeqNo"))
            .Fields("FaciSNo") = NoZero(rsSO004("FaciSNo"))
            
            .Update
        End With
    End If
    rsUC.Close
    rsCust.Close
    Set rsUC = Nothing
    Set rsCust = Nothing
Exit Sub

ChkErr:
    Call ErrSub(Me.Name, "GenerateSO033Data")
End Sub
'//  �N��z�n����� �A��JSO033 ��ƪ�
Private Function InsertIntoSO033() As Boolean

Dim strInsertInto As String

On Error GoTo ChkErr

   InsertIntoSO033 = False
   rsTmp.MoveFirst
  strInsertInto = "INSERT INTO " & GetOwner & "SO033 "
  
  strInsertInto = strInsertInto & _
                        "(Custid,BillNo,Item," & _
                        "CitemCode,CitemName," & _
                        "OldAmt,ShouldDate," & _
                        "ShouldAmt,ClctEn," & _
                        "ClctName,CMCode," & _
                        "CMName,UCCode," & _
                        "UCName,Note,CreateTime,CreateEn,CompCode,AddrNo,ServCode," & _
                        "ClctAreaCode,NodeNo,PTCode,PTName,OldClctEn," & _
                        "OldClctName,ClassCode,UpdTime,UpdEn,ServiceType,Quantity," & _
                        "CancelFlag,AreaCode,StrtCode,MduId,FaciSeqNo ,FaciSNo )"
             
  strInsertInto = strInsertInto & " VALUES "
  
   strInsertInto = strInsertInto & "(" & GetNullString(rsTmp(0), giStringV) & " ," & GetNullString(rsTmp(1), giStringV) & "," & _
                                                        GetNullString(rsTmp(2), giLongV) & " ," & GetNullString(rsTmp(3), giLongV) & "," & _
                                                        GetNullString(rsTmp(4), giStringV) & " ," & GetNullString(rsTmp(5), giLongV) & _
                                                        ",TO_DATE('" & rsTmp(6) & "','YYYY/MM/DD' )," & GetNullString(rsTmp(7), giLongV) & "," & _
                                                        GetNullString(rsTmp(8), giStringV) & " ," & GetNullString(rsTmp(9), giStringV) & "," & _
                                                        GetNullString(rsTmp(10), giLongV) & " ," & GetNullString(rsTmp(11), giStringV) & "," & _
                                                        GetNullString(rsTmp(12), giLongV) & " ," & GetNullString(rsTmp(13), giStringV) & "," & _
                                                        GetNullString(rsTmp(14), giStringV) & "," & GetNullString(rsTmp(15), giDateV, , True) & "," & _
                                                        GetNullString(rsTmp(16), giStringV) & "," & GetNullString(rsTmp(17), giLongV) & "," & _
                                                        GetNullString(rsTmp(18), giLongV) & " ," & GetNullString(rsTmp(19), giStringV) & "," & _
                                                        GetNullString(rsTmp(20), giStringV) & " ," & GetNullString(rsTmp(21), giStringV) & "," & _
                                                        GetNullString(rsTmp(22), giLongV) & " ," & GetNullString(rsTmp(23), giStringV) & "," & _
                                                        GetNullString(rsTmp(24), giStringV) & "," & GetNullString(rsTmp(25), giStringV) & "," & _
                                                        GetNullString(rsTmp(26), giStringV) & " ," & GetNullString(rsTmp(27), giStringV) & "," & _
                                                        GetNullString(rsTmp(28), giStringV) & " ," & GetNullString(rsTmp(29), giStringV) & "," & GetNullString(rsTmp(30), giLongV) & "," & _
                                                        GetNullString(rsTmp(31), giLongV) & " ," & GetNullString(rsTmp(32), giStringV) & "," & _
                                                        GetNullString(rsTmp(33), giLongV) & " ," & GetNullString(rsTmp(34), giStringV) & "," & _
                                                        GetNullString(rsTmp(35), giStringV) & " ," & GetNullString(rsTmp(36), giStringV) & ")"
 
 gcnGi.Execute strInsertInto
 InsertIntoSO033 = True
Exit Function

ChkErr:
    Call ErrSub(Me.Name, "InsertIntoSO033")
End Function

Public Property Get rsReturn() As ADODB.Recordset
      Set rsReturn = rsTmp
End Property

Private Sub rsSO004_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo ChkErr
If Not rsSO004.EOF And Not rsSO004.BOF Then
    Select Case rsSO004("BUYCODE")
     Case 1  ''''  �R (�h�O)
            SetgiList gilReturnItme, "CodeNo", "Description", "CD019", , , , , , , " Where   periodflag  =  0 AND REFNO = 7   AND StopFlag  = 0  AND  ByHouse=0   "
     Case 3  ''''  ��
'            If rsSO004("ContractCust") = 1 Then
'                 SetgiList gilReturnItme, "CodeNo", "Description", "CD019", , , , , , , " Where   periodflag  =  0 AND REFNO = 14  AND StopFlag  = 0  "
'            Else
                 SetgiList gilReturnItme, "CodeNo", "Description", "CD019", , , , , , , " Where   periodflag  =  0 AND REFNO = 11  AND StopFlag  = 0  AND  ByHouse=0 "
'            End If
     Case 4  ''''  ��
           SetgiList gilReturnItme, "CodeNo", "Description", "CD019", , , , , , , " Where   periodflag  =  0 AND REFNO = 12  AND StopFlag  = 0 AND  ByHouse=0  "
          
    End Select
    
    If gilReturnItme.GetListCount <= 0 Then
       gilReturnItme.Filter = " where periodflag  = 0 AND SIGN = '-' AND ByHouse = 0 "
    End If
    gilReturnItme.ListIndex = 1
End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "rsSO004_MoveComplete")
End Sub
Private Sub SetDefaultClct()
Dim strTmpSql As String
Dim rsTmpClct As New ADODB.Recordset
On Error GoTo ChkErr
    strTmpSql = "Select B.ClctEn,B.ClctName From " & _
                        GetOwner & "So001 A," & GetOwner & "So014 B Where " & _
                        "A.ChargeAddrNo=B.AddrNo And " & _
                        "A.CustId = " & rsSO004("CustId") & " And A.CompCode = " & gilCompCode.GetCodeNo
                        
                    
    If Not GetRS(rsTmpClct, strTmpSql, gcnGi) Then Exit Sub
    
    If rsTmpClct.RecordCount > 0 Then
            gilClctEn.SetCodeNo rsTmpClct("ClctEn") & ""
            gilClctEn.SetDescription rsTmpClct("ClctName") & ""
    End If
Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "SetDefaultClct")
End Sub
