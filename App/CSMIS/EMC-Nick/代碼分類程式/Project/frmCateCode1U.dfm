object frmCateCode1: TfrmCateCode1
  Left = 85
  Top = 39
  Width = 600
  Height = 400
  Caption = 'SO8F10_1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object Splitter1: TSplitter
    Left = 0
    Top = 87
    Width = 592
    Height = 3
    Cursor = crVSplit
    Align = alBottom
  end
  object pnlCommand: TPanel
    Left = 0
    Top = 332
    Width = 592
    Height = 41
    Align = alBottom
    TabOrder = 0
    object btnAppend: TBitBtn
      Left = 24
      Top = 8
      Width = 75
      Height = 25
      Caption = #26032#22686
      TabOrder = 0
      OnClick = btnAppendClick
    end
    object btnEdit: TBitBtn
      Left = 112
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
    object btnEditData: TBitBtn
      Left = 472
      Top = 8
      Width = 113
      Height = 25
      Caption = #32232#36655#20998#39006#23565#29031#36039#26009
      TabOrder = 5
      OnClick = btnEditDataClick
    end
  end
  object pnlMultiData: TPanel
    Left = 0
    Top = 90
    Width = 592
    Height = 242
    Align = alBottom
    BevelInner = bvLowered
    TabOrder = 1
    object DBGrid1: TDBGrid
      Left = 2
      Top = 2
      Width = 588
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
          Width = 64
          Visible = True
        end>
    end
  end
  object pnlSingleData: TPanel
    Left = 0
    Top = 0
    Width = 592
    Height = 87
    Align = alClient
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 2
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
    DataSet = dtmMain.adoCD067A
    Left = 368
    Top = 8
  end
end
