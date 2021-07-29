object frmSO8C10: TfrmSO8C10
  Left = 251
  Top = 114
  Caption = 'frmSO8C10'
  ClientHeight = 503
  ClientWidth = 742
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object Splitter1: TSplitter
    Left = 0
    Top = 245
    Width = 742
    Height = 3
    Cursor = crVSplit
    Align = alTop
  end
  object pnlMultiData: TPanel
    Left = 0
    Top = 248
    Width = 742
    Height = 214
    Align = alClient
    TabOrder = 0
    object DBGrid1: TDBGrid
      Left = 1
      Top = 1
      Width = 740
      Height = 212
      Align = alClient
      Color = clInfoBk
      DataSource = dsrSo126
      ReadOnly = True
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -13
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
      OnTitleClick = DBGrid1TitleClick
      Columns = <
        item
          Expanded = False
          FieldName = 'EMPNAME'
          Width = 120
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'GETPAPERDATE'
          Width = 100
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'PREFIX'
          Title.Caption = #23383#38957
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'BEGINNUM'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'ENDNUM'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'TOTALPAPERCOUNT'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'OPERATOR'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'RETURNDATE'
          Width = 100
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'CLEARDATE'
          Title.Caption = #28165#38500#26085#26399
          Width = 100
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'NOTE'
          Width = 400
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'UPDTIME'
          Visible = True
        end>
    end
  end
  object pnlSingleData: TPanel
    Left = 0
    Top = 103
    Width = 742
    Height = 142
    Align = alTop
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 1
    object StaticText5: TStaticText
      Left = 17
      Top = 12
      Width = 52
      Height = 20
      Caption = #38936#21934#26085#26399
      TabOrder = 8
    end
    inline fraChineseYMD3: TfraChineseYMD
      Left = 72
      Top = 6
      Width = 78
      Height = 31
      TabOrder = 0
      ExplicitLeft = 72
      ExplicitTop = 6
      ExplicitWidth = 78
      ExplicitHeight = 31
      inherited mseYMD: TMaskEdit
        Top = 3
        Height = 24
        OnExit = fraChineseYMD3mseYMDExit
        ExplicitTop = 3
        ExplicitHeight = 24
      end
    end
    object StaticText6: TStaticText
      Left = 161
      Top = 12
      Width = 52
      Height = 20
      Caption = #38936#21934#20154#21729
      TabOrder = 9
    end
    object StaticText7: TStaticText
      Left = 30
      Top = 71
      Width = 40
      Height = 20
      Caption = #27969#27700#34399
      TabOrder = 10
    end
    object DBEdit1: TDBEdit
      Left = 216
      Top = 12
      Width = 121
      Height = 24
      TabStop = False
      Color = clBtnFace
      DataField = 'EMPNAME'
      DataSource = dsrSo126
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      TabOrder = 11
    end
    object DBEdit2: TDBEdit
      Left = 72
      Top = 73
      Width = 121
      Height = 24
      DataField = 'BEGINNUM'
      DataSource = dsrSo126
      TabOrder = 3
    end
    object StaticText8: TStaticText
      Left = 198
      Top = 75
      Width = 13
      Height = 20
      Caption = '~'
      TabOrder = 12
    end
    object StaticText9: TStaticText
      Left = 344
      Top = 71
      Width = 28
      Height = 20
      Caption = #24373#25976
      TabOrder = 13
    end
    object DBEdit4: TDBEdit
      Left = 375
      Top = 69
      Width = 68
      Height = 24
      DataField = 'TOTALPAPERCOUNT'
      DataSource = dsrSo126
      TabOrder = 4
    end
    object DBEdit3: TDBEdit
      Left = 212
      Top = 70
      Width = 121
      Height = 24
      TabStop = False
      AutoSize = False
      BevelInner = bvNone
      BevelOuter = bvNone
      Color = clBtnFace
      DataField = 'ENDNUM'
      DataSource = dsrSo126
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 14
    end
    object btnGetPaperUser: TButton
      Left = 341
      Top = 12
      Width = 23
      Height = 25
      Caption = '...'
      TabOrder = 1
      OnClick = btnGetPaperUserClick
    end
    object StaticText10: TStaticText
      Left = 41
      Top = 42
      Width = 28
      Height = 20
      Caption = #23383#38957
      TabOrder = 15
    end
    object DBEdit5: TDBEdit
      Left = 72
      Top = 40
      Width = 121
      Height = 24
      DataField = 'PREFIX'
      DataSource = dsrSo126
      TabOrder = 2
    end
    object StaticText12: TStaticText
      Left = 199
      Top = 42
      Width = 120
      Height = 20
      Caption = '('#22823#23567#23531#26159#19981#19968#27171#30340#21908')'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 16
    end
    object StaticText14: TStaticText
      Left = 16
      Top = 108
      Width = 52
      Height = 20
      Caption = #36864#22238#26085#26399
      TabOrder = 17
    end
    object StaticText15: TStaticText
      Left = 161
      Top = 108
      Width = 52
      Height = 20
      Caption = #28165#26597#26085#26399
      TabOrder = 18
    end
    inline fraChineseYMD5: TfraChineseYMD
      Left = 216
      Top = 100
      Width = 78
      Height = 31
      TabOrder = 6
      ExplicitLeft = 216
      ExplicitTop = 100
      ExplicitWidth = 78
      ExplicitHeight = 31
      inherited mseYMD: TMaskEdit
        Top = 3
        Height = 24
        OnExit = fraChineseYMD5mseYMDExit
        ExplicitTop = 3
        ExplicitHeight = 24
      end
    end
    object DBMemo1: TDBMemo
      Left = 465
      Top = 14
      Width = 262
      Height = 111
      DataField = 'NOTE'
      DataSource = dsrSo126
      ScrollBars = ssBoth
      TabOrder = 7
    end
    object StaticText16: TStaticText
      Left = 431
      Top = 16
      Width = 28
      Height = 20
      Caption = #20633#35387
      TabOrder = 19
    end
    inline fraChineseYMD4: TfraChineseYMD
      Left = 72
      Top = 102
      Width = 78
      Height = 31
      TabOrder = 5
      ExplicitLeft = 72
      ExplicitTop = 102
      ExplicitWidth = 78
      ExplicitHeight = 31
      inherited mseYMD: TMaskEdit
        Top = 3
        Height = 24
        ExplicitTop = 3
        ExplicitHeight = 24
      end
    end
  end
  object pnlQuery: TPanel
    Left = 0
    Top = 0
    Width = 742
    Height = 103
    Align = alTop
    BevelInner = bvLowered
    BevelOuter = bvNone
    TabOrder = 2
    object StaticText1: TStaticText
      Left = 17
      Top = 11
      Width = 52
      Height = 20
      Caption = #38936#21934#26085#26399
      TabOrder = 6
    end
    object StaticText2: TStaticText
      Left = 284
      Top = 11
      Width = 52
      Height = 20
      Caption = #38936#21934#20154#21729
      TabOrder = 7
    end
    object StaticText3: TStaticText
      Left = 30
      Top = 74
      Width = 40
      Height = 20
      Caption = #27969#27700#34399
      TabOrder = 8
    end
    object btnQuery: TBitBtn
      Left = 624
      Top = 41
      Width = 75
      Height = 25
      Caption = #26597#35426
      TabOrder = 5
      OnClick = btnQueryClick
    end
    inline fraChineseYMD1: TfraChineseYMD
      Left = 72
      Top = 6
      Width = 82
      Height = 31
      TabOrder = 0
      ExplicitLeft = 72
      ExplicitTop = 6
      ExplicitWidth = 82
      ExplicitHeight = 31
      inherited mseYMD: TMaskEdit
        Top = 4
        Height = 24
        ExplicitTop = 4
        ExplicitHeight = 24
      end
    end
    object edtPaperNum: TEdit
      Left = 72
      Top = 72
      Width = 156
      Height = 24
      TabOrder = 4
    end
    object StaticText4: TStaticText
      Left = 156
      Top = 15
      Width = 13
      Height = 20
      Caption = '~'
      TabOrder = 9
    end
    inline fraChineseYMD2: TfraChineseYMD
      Left = 178
      Top = 8
      Width = 83
      Height = 27
      TabOrder = 1
      ExplicitLeft = 178
      ExplicitTop = 8
      ExplicitWidth = 83
      ExplicitHeight = 27
      inherited mseYMD: TMaskEdit
        Top = 2
        Height = 24
        ExplicitTop = 2
        ExplicitHeight = 24
      end
    end
    object lbxEmp: TListBox
      Left = 340
      Top = 8
      Width = 277
      Height = 85
      TabStop = False
      Columns = 4
      ItemHeight = 16
      TabOrder = 10
    end
    object btnEmp: TBitBtn
      Left = 623
      Top = 8
      Width = 25
      Height = 25
      Caption = '...'
      TabOrder = 2
      OnClick = btnEmpClick
    end
    object StaticText11: TStaticText
      Left = 41
      Top = 44
      Width = 28
      Height = 20
      Caption = #23383#38957
      TabOrder = 11
    end
    object edtPrefix: TEdit
      Left = 71
      Top = 41
      Width = 67
      Height = 24
      MaxLength = 5
      TabOrder = 3
    end
    object StaticText13: TStaticText
      Left = 145
      Top = 42
      Width = 120
      Height = 20
      Caption = '('#22823#23567#23531#26159#19981#19968#27171#30340#21908')'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 12
    end
  end
  object pnlCommand: TPanel
    Left = 0
    Top = 462
    Width = 742
    Height = 41
    Align = alBottom
    TabOrder = 3
    object btnAppend: TBitBtn
      Left = 40
      Top = 8
      Width = 75
      Height = 25
      Caption = #26032#22686
      TabOrder = 0
      OnClick = btnAppendClick
    end
    object btnEdit: TBitBtn
      Left = 120
      Top = 8
      Width = 75
      Height = 25
      Caption = #20462#25913
      TabOrder = 1
      OnClick = btnEditClick
    end
    object btnDelete: TBitBtn
      Left = 200
      Top = 8
      Width = 75
      Height = 25
      Caption = #21034#38500
      TabOrder = 2
      OnClick = btnDeleteClick
    end
    object btnCancelExit: TBitBtn
      Left = 376
      Top = 8
      Width = 75
      Height = 25
      Caption = #32080#26463
      TabOrder = 4
      OnClick = btnCancelExitClick
    end
    object btnSave: TBitBtn
      Left = 288
      Top = 8
      Width = 75
      Height = 25
      Caption = #23384#27284
      TabOrder = 3
      OnClick = btnSaveClick
    end
    object btnReUse: TBitBtn
      Left = 464
      Top = 8
      Width = 75
      Height = 25
      Caption = #37559#21934#32396#29992
      TabOrder = 5
      OnClick = btnReUseClick
    end
    object btnPrint: TBitBtn
      Left = 550
      Top = 8
      Width = 75
      Height = 25
      Caption = #21015#21360
      TabOrder = 6
      OnClick = btnPrintClick
    end
  end
  object dsrSo126: TDataSource
    DataSet = dtmMain3.cdsSo126
    OnDataChange = dsrSo126DataChange
    Left = 92
    Top = 308
  end
end
