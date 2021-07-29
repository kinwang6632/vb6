object frmSO8A60: TfrmSO8A60
  Left = 303
  Top = 171
  Width = 696
  Height = 480
  Caption = #29256#27402#25104#26412#20998#25892#27604#20363#35373#23450
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object pnlMultiData: TPanel
    Left = 0
    Top = 265
    Width = 688
    Height = 159
    Align = alClient
    BevelOuter = bvLowered
    TabOrder = 0
    object DBGrid1: TDBGrid
      Left = 1
      Top = 1
      Width = 686
      Height = 157
      Align = alClient
      Color = clInfoBk
      DataSource = dsrSo112A
      ReadOnly = True
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -13
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
      Columns = <
        item
          Expanded = False
          FieldName = 'CODENO'
          Title.Caption = #25910#36027#38917#30446#20195#30908
          Width = 80
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'DESCRIPTION'
          Title.Caption = #25910#36027#38917#30446
          Width = 180
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'OPERATOR'
          Title.Caption = #30064#21205#20154#21729
          Width = 80
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'UPDTIME'
          Title.Caption = #30064#21205#26178#38291
          Width = 120
          Visible = True
        end>
    end
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 424
    Width = 688
    Height = 29
    Panels = <>
  end
  object pnlSingleData: TPanel
    Left = 0
    Top = 65
    Width = 688
    Height = 200
    Align = alTop
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 2
    object pgcPercentData: TPageControl
      Left = 367
      Top = 2
      Width = 319
      Height = 196
      ActivePage = TabSheet1
      Align = alRight
      Style = tsFlatButtons
      TabOrder = 0
      object TabSheet1: TTabSheet
        Caption = #20998#25892#26126#32048
        object dbgPercentXmlData: TDBGrid
          Left = 0
          Top = 0
          Width = 311
          Height = 162
          Hint = #25353#21491#37749#32232#36655#36039#26009
          Align = alClient
          Color = clTeal
          DataSource = dsrPercentXmlData
          ParentShowHint = False
          PopupMenu = PopupMenu1
          ShowHint = True
          TabOrder = 0
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -13
          TitleFont.Name = 'MS Sans Serif'
          TitleFont.Style = []
          OnEditButtonClick = dbgPercentXmlDataEditButtonClick
          Columns = <
            item
              ButtonStyle = cbsEllipsis
              Expanded = False
              FieldName = 'sProviderName'
              ReadOnly = True
              Title.Caption = #20379#25033#21830#21517#31281
              Width = 100
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'nPercent'
              Title.Caption = #20998#25892#30334#20998#27604
              Visible = True
            end>
        end
      end
    end
    object pnlMasterData: TPanel
      Left = 2
      Top = 2
      Width = 365
      Height = 196
      Align = alClient
      TabOrder = 1
      object DBText1: TDBText
        Left = 104
        Top = 24
        Width = 65
        Height = 17
        DataField = 'CODENO'
        DataSource = dsrSo112A
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
      end
      object StaticText1: TStaticText
        Left = 22
        Top = 24
        Width = 76
        Height = 20
        Caption = #25910#36027#38917#30446#20195#30908
        TabOrder = 0
      end
      object StaticText2: TStaticText
        Left = 46
        Top = 51
        Width = 52
        Height = 20
        Caption = #25910#36027#38917#30446
        TabOrder = 1
      end
      object DBLookupComboBox1: TDBLookupComboBox
        Left = 103
        Top = 47
        Width = 256
        Height = 24
        DataField = 'CODENO'
        DataSource = dsrSo112A
        KeyField = 'CODENO'
        ListField = 'REFDATA'
        ListSource = dsrCd019
        TabOrder = 2
        OnCloseUp = DBLookupComboBox1CloseUp
      end
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 0
    Width = 688
    Height = 65
    Align = alTop
    BevelInner = bvLowered
    TabOrder = 3
    object btnAppend: TBitBtn
      Left = 32
      Top = 24
      Width = 75
      Height = 25
      Caption = #26032#22686
      TabOrder = 0
      OnClick = btnAppendClick
    end
    object btnEdit: TBitBtn
      Left = 112
      Top = 24
      Width = 75
      Height = 25
      Caption = #20462#25913
      TabOrder = 1
      OnClick = btnEditClick
    end
    object btnDelete: TBitBtn
      Left = 192
      Top = 24
      Width = 75
      Height = 25
      Caption = #21034#38500
      TabOrder = 2
      OnClick = btnDeleteClick
    end
    object btnCancel: TBitBtn
      Left = 272
      Top = 24
      Width = 75
      Height = 25
      Caption = #21462#28040
      TabOrder = 3
      OnClick = btnCancelClick
    end
    object btnExit: TBitBtn
      Left = 432
      Top = 24
      Width = 75
      Height = 25
      Caption = #32080#26463
      TabOrder = 5
      OnClick = btnExitClick
    end
    object btnSave: TBitBtn
      Left = 352
      Top = 24
      Width = 75
      Height = 25
      Caption = #20786#23384
      TabOrder = 4
      OnClick = btnSaveClick
    end
  end
  object dsrSo112A: TDataSource
    DataSet = dtmMain2.cdsSo112A
    Left = 208
    Top = 144
  end
  object dsrCd019: TDataSource
    DataSet = dtmMain2.cdsCd019A
    Left = 160
    Top = 144
  end
  object dsrPercentXmlData: TDataSource
    DataSet = dtmMain2.cdsPercentXmlData
    OnStateChange = dsrPercentXmlDataStateChange
    Left = 280
    Top = 144
  end
  object PopupMenu1: TPopupMenu
    Left = 396
    Top = 184
    object miInsert: TMenuItem
      Caption = #26032#22686
      OnClick = miInsertClick
    end
    object miEdit: TMenuItem
      Caption = #20462#25913
      OnClick = miEditClick
    end
    object miDelete: TMenuItem
      Caption = #21034#38500
      OnClick = miDeleteClick
    end
    object miCancel: TMenuItem
      Caption = #21462#28040
      OnClick = miCancelClick
    end
    object miSave: TMenuItem
      Caption = #20786#23384
      OnClick = miSaveClick
    end
  end
  object dsrSO113ForUSeful: TDataSource
    DataSet = dtmMain2.cdsSO113ForUSeful
    Left = 160
    Top = 192
  end
end
