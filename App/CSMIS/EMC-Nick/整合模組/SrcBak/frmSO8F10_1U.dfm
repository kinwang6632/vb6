object frmSO8F10_1: TfrmSO8F10_1
  Left = 236
  Top = 141
  Width = 645
  Height = 400
  Caption = 'SO8F10_1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object Splitter1: TSplitter
    Left = 0
    Top = 74
    Width = 637
    Height = 3
    Cursor = crVSplit
    Align = alBottom
  end
  object pnlCommand: TPanel
    Left = 0
    Top = 319
    Width = 637
    Height = 53
    Align = alBottom
    TabOrder = 0
    object btnAppend: TBitBtn
      Left = 16
      Top = 8
      Width = 72
      Height = 31
      Caption = #26032#22686
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      OnClick = btnAppendClick
    end
    object btnEdit: TBitBtn
      Left = 88
      Top = 8
      Width = 75
      Height = 31
      Caption = #20462#25913
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      OnClick = btnEditClick
    end
    object btnDelete: TBitBtn
      Left = 163
      Top = 8
      Width = 75
      Height = 31
      Caption = #21034#38500
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      OnClick = btnDeleteClick
    end
    object btnCancelExit: TBitBtn
      Left = 312
      Top = 8
      Width = 75
      Height = 31
      Caption = #32080#26463
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 4
      OnClick = btnCancelExitClick
    end
    object btnSave: TBitBtn
      Left = 239
      Top = 8
      Width = 75
      Height = 31
      Caption = #23384#27284
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      OnClick = btnSaveClick
    end
    object btnEditData: TBitBtn
      Left = 408
      Top = 8
      Width = 145
      Height = 31
      Caption = #32232#36655#20998#39006#23565#29031#36039#26009
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 5
      OnClick = btnEditDataClick
    end
    object btnSynchronization: TBitBtn
      Left = 553
      Top = 8
      Width = 73
      Height = 31
      Caption = #36039#26009#21516#27493
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 6
      OnClick = btnSynchronizationClick
    end
  end
  object pnlMultiData: TPanel
    Left = 0
    Top = 77
    Width = 637
    Height = 242
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 1
    object DBGrid1: TDBGrid
      Left = 2
      Top = 2
      Width = 633
      Height = 238
      Align = alClient
      Color = clInfoBk
      DataSource = dsrCD067A
      ReadOnly = True
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -13
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
      OnDblClick = DBGrid1DblClick
      OnTitleClick = DBGrid1TitleClick
      Columns = <
        item
          Expanded = False
          FieldName = 'TABLENAME'
          Width = 120
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'TABLEDESCRIPTION'
          Width = 150
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'OPERATOR'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'UPDTIME'
          Width = 150
          Visible = True
        end>
    end
  end
  object pnlSingleData: TPanel
    Left = 0
    Top = 0
    Width = 637
    Height = 74
    Align = alClient
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 2
    object Label1: TLabel
      Left = 576
      Top = 8
      Width = 54
      Height = 16
      Alignment = taRightJustify
      Caption = 'Label1'
      Font.Charset = ANSI_CHARSET
      Font.Color = clRed
      Font.Height = -16
      Font.Name = #32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
    object StaticText5: TStaticText
      Left = 17
      Top = 12
      Width = 83
      Height = 20
      Caption = 'Table Name '
      TabOrder = 2
    end
    object StaticText7: TStaticText
      Left = 22
      Top = 47
      Width = 72
      Height = 20
      Caption = 'Description'
      TabOrder = 3
    end
    object DBEdit2: TDBEdit
      Left = 104
      Top = 45
      Width = 465
      Height = 24
      DataField = 'TABLEDESCRIPTION'
      DataSource = dsrCD067A
      TabOrder = 1
    end
    object DBEdit5: TDBEdit
      Left = 104
      Top = 8
      Width = 209
      Height = 24
      DataField = 'TABLENAME'
      DataSource = dsrCD067A
      TabOrder = 0
    end
  end
  object dsrCD067A: TDataSource
    DataSet = dtmMain4.adoCD067A
    Left = 368
    Top = 8
  end
end
