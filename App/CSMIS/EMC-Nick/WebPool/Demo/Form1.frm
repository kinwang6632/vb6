VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   8385
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2250
      TabIndex        =   4
      Top             =   570
      Width           =   5265
   End
   Begin VB.CommandButton Command3 
      Caption         =   "GetCommand"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8705
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "GetRecordSet"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GetConnection"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
'這是呼叫使用 Delphi 寫的 Connection Pool
  
  Dim aPool As New PoolManager
  Dim aConnection As ADODB.Connection
  
  ' 使用 COM+ Pool 元件的 GetConnection 方法
  Set aConnection = aPool.GetConnection
  
  Dim aRecordSet As New ADODB.Recordset
    
  '取得的 ADO Connection 已經 Open, 不須先 Open
  '用 ADO RecordSet 的 Open 方法使用取得的 Conneciton 會有錯誤訊息
  'aRecordSet.Open "select * from so001 where rownum <= 100", aConnection, adOpenKeyset, adLockOptimistic
  
  '直接使用 ADO Connection 的 Execute 取 RecordSet 方法 OK
  Set aRecordSet = aConnection.Execute(Text1.Text)
  Set DataGrid1.DataSource = aRecordSet
  
End Sub

Private Sub Command2_Click()
  '這是呼叫使用 Delphi 寫的 Connection Pool
  
  'Dim aPool As New PoolManager
    
  'Dim aRecordSet As ADODB.Recordset
    
  '使用 COM+ Pool 元件的 GetRecordSet 方法
  'Set aRecordSet = aPool.GetRecordSet("SELECT * FROM INV099 where yearmonth like '2004%'")
  'Set DataGrid1.DataSource = aRecordSet
  
    Dim o(20) As PoolManager
    Dim cn(20) As ADODB.Connection
    Dim intLoop As Integer
    For intLoop = 1 To 20
        'Set o(intLoop) = CreateObject("WebPool.PoolManager")
        Set o(intLoop) = New PoolManager
        Set cn(intLoop) = o(intLoop).GetConnection
        DoEvents
        Debug.Print o(intLoop).GetRecordSet(Text1.Text).GetString
        
        Debug.Print intLoop
        Set o(intLoop) = Nothing
        Set cn(intLoop) = Nothing
        DoEvents
    Next
  
  
End Sub

Private Sub Command3_Click()
    Dim aPool As New PoolManager
    Dim aCommand As ADODB.Command
    Dim aRecordSet As ADODB.Recordset
    
    Set aCommand = aPool.GetCommand
    aCommand.CommandText = "select * from so001 where rownum <= 100"
    Set aRecordSet = aCommand.Execute
    Set DataGrid1.DataSource = aRecordSet
End Sub
