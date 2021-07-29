VERSION 5.00
Object = "{3BB22E57-824A-11D3-BD70-0080C8F80BC4}#3.4#0"; "GiGridR.ocx"
Begin VB.Form frmSO114FB 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "���ʰO��[SO114FB]"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   10455
   Icon            =   "SO114FB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   10455
   StartUpPosition =   1  '���ݵ�������
   Begin prjGiGridR.GiGridR ggrData 
      Height          =   4125
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   7276
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
End
Attribute VB_Name = "frmSO114FB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FInvSeqNo As String
Private rsLog As New ADODB.Recordset
Public Property Let uInvSeqNo(ByVal vData As String)
  FInvSeqNo = vData
End Property

Private Sub Form_Load()
    Call subGrd
    Call OpenData
End Sub
Private Sub subGrd()
    On Error Resume Next
    Dim mFlds As New GiGridFlds
    Dim mFlds2 As New GiGridFlds
    With mFlds
        .Add "FuncType", , , , , "���ʺ���"
        .Add "UpdTime", , , , , "���ʮɶ�" & Space(10)
        .Add "UpdEn", , , , , "���ʤH��" & Space(10)
        .Add "InvSeqNo", , , , , "�y���s��     "
        .Add "ChargeTitle", , , , , "����H" & Space(40)
        .Add "InvoiceType", , , , , "�o������"
        .Add "InvNo", , , , , "�o���Τ@�s��"
        .Add "InvTitle", , , , , "�o�����Y" & Space(50)
        .Add "InvAddress", , , , , "�o���a�}" & Space(50)
        .Add "ChargeAddress", , , , , "���O�a�}" & Space(50)
        .Add "MailAddress", , , , , "�l�H�a�}" & Space(50)
        .Add "StopFlag", , , , , "����"
        .Add "StopDate", giControlTypeDate, , , , "���Τ��" & Space(5)
        .Add "InvPurposeName", , , , , "�o���γ~�W��" & Space(10)
        .Add "PreInvoice", , , , , "�O�_�w�}�o��"
        .Add "InvoiceKind", , , , , "�o���}�ߺ���"
        .Add "BillMailKind", , , , , "�b��H�e�覡"
        .Add "DenRecName", , , , , "���س��W��"
        .Add "DenRecDate", giControlTypeDate, , , , "���z�ذe���" & Space(10)
        .Add "ApplyInvDate", giControlTypeDate, , , , "�ӽйq�l�p����o�����"
        
    End With
    ggrData.AllFields = mFlds
    ggrData.SetHead
End Sub
Private Sub OpenData()
  On Error GoTo ChkErr
    Dim aSQL As String
    aSQL = "SELECT * FROM " & GetOwner & "SO138LOG WHERE InvSeqNo=" & FInvSeqNo & _
        " ORDER BY UpdTime DESC"
    If Not GetRS(rsLog, aSQL, gcnGi, adUseClient, adOpenKeyset, adLockOptimistic) Then Exit Sub
    Set ggrData.Recordset = rsLog
    ggrData.Refresh
    Exit Sub
ChkErr:
    Call ErrSub(Me.Name, "OpenData")
End Sub


Private Sub ggrData_ShowCellData(ByVal Col As Long, ByVal Row As Long, ByVal giFld As prjGiGridR.GiGridFld, ByVal Fld As ADODB.Field, Value As Variant)
    On Error Resume Next
    If giFld.FieldName = "InvoiceType" Then
        Select Case Value
            Case 0
                Value = "�K�}"
            Case 2
                Value = "�G�p��"
            Case 3
                Value = "�T�p��"
        End Select
    End If
    If giFld.FieldName = "StopFlag" Then
        Select Case Value
            Case 0
                Value = "�_"
            Case 1
                Value = "�O"
            Case Else
                Value = "�_"
        End Select
    End If
    If giFld.FieldName = "PreInvoice" Then
        Select Case Value
            Case 0
                Value = "��}��"
            Case 1
                Value = "�w�}��"
            Case Else
                Value = "��}��"
        End Select
    End If
    If giFld.FieldName = "InvoiceKind" Then
        Select Case Value
            Case 0
                Value = "�q�l�p���"
            Case 1
                Value = "�q�l�o��"
            Case Else
                Value = "�q�l�p���"
                
        End Select
    End If
    '0=�@��b��,1=�q�l�b��
    If giFld.FieldName = "BillMailKind" Then
        Select Case Value
            Case 0
                Value = "�@��b��"
            Case 1
                Value = "�q�l�b��"
            Case Else
                Value = "�@��b��"
        End Select
    End If
    '1=�ק�, 2=�R��, 3=�s�W
    If giFld.FieldName = "FuncType" Then
        Select Case Value
            Case 1
                Value = "�ק�"
            Case 2
                Value = "�R��"
            Case 3
                Value = "�s�W"
        End Select
    End If
End Sub
