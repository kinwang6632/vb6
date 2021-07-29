object frmSO8B10: TfrmSO8B10
  Left = 4
  Top = 0
  Width = 796
  Height = 547
  Caption = 'frmSO8B10'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 0
    Top = 474
    Width = 788
    Height = 3
    Cursor = crVSplit
    Align = alBottom
  end
  object pnlEdit: TPanel
    Left = 0
    Top = 0
    Width = 788
    Height = 185
    Align = alTop
    TabOrder = 0
    object Label2: TLabel
      Left = 4
      Top = 53
      Width = 84
      Height = 13
      Alignment = taRightJustify
      Caption = #25910#36027#38917#30446#21517#31281
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label3: TLabel
      Left = 27
      Top = 90
      Width = 56
      Height = 13
      Alignment = taRightJustify
      Caption = #20323#37329#21934#20301
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object dtxCharge: TDBText
      Left = 95
      Top = 53
      Width = 136
      Height = 20
      Color = clInfoBk
      DataField = 'DESCRIPTION'
      DataSource = dsrSO120
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object Label1: TLabel
      Left = 272
      Top = 14
      Width = 42
      Height = 13
      Alignment = taRightJustify
      Caption = #21443#32771#34399
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label4: TLabel
      Left = 46
      Top = 13
      Width = 42
      Height = 13
      Alignment = taRightJustify
      Caption = #20844#21496#21029
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label11: TLabel
      Left = 27
      Top = 132
      Width = 56
      Height = 13
      Alignment = taRightJustify
      Caption = #26989#21209#27963#21205
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label7: TLabel
      Left = 268
      Top = 54
      Width = 182
      Height = 13
      Alignment = taRightJustify
      Caption = #38971#36947#38656#25910#35222#28415#22810#23569#22825#25165#32102#20323#37329
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object btnQueryCharge: TBitBtn
      Left = 236
      Top = 53
      Width = 24
      Height = 24
      Caption = '...'
      TabOrder = 1
      OnClick = btnQueryChargeClick
    end
    object dedRefNo: TDBEdit
      Left = 327
      Top = 13
      Width = 170
      Height = 21
      DataField = 'REF_NO'
      DataSource = dsrSO120
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      MaxLength = 3
      ParentFont = False
      TabOrder = 4
    end
    object cobComp: TComboBox
      Left = 95
      Top = 10
      Width = 138
      Height = 21
      Style = csDropDownList
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 0
    end
    object cobPayUnit: TComboBox
      Left = 95
      Top = 88
      Width = 141
      Height = 21
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 2
      OnChange = cobPayUnitChange
      Items.Strings = (
        '1-'#30334#20998#27604
        '2-'#20803)
    end
    object gbxFirst: TGroupBox
      Left = 272
      Top = 90
      Width = 201
      Height = 81
      Caption = '   '#39318#27425' - '#20184#27454#26041#24335'('#25512#24291#20154')'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 6
      object Label5: TLabel
        Left = 23
        Top = 27
        Width = 42
        Height = 13
        Alignment = taRightJustify
        Caption = #20449#29992#21345
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label6: TLabel
        Left = 8
        Top = 52
        Width = 56
        Height = 13
        Alignment = taRightJustify
        Caption = #38750#20449#29992#21345
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object lblUnit1: TLabel
        Left = 174
        Top = 27
        Width = 7
        Height = 13
        Caption = '%'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object lblUnit2: TLabel
        Left = 175
        Top = 56
        Width = 7
        Height = 13
        Caption = '%'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object dedFirstCreditCard1: TDBEdit
        Left = 68
        Top = 19
        Width = 101
        Height = 24
        DataField = 'FIRSTCREDITCARD1'
        DataSource = dsrSO120
        TabOrder = 0
      end
      object dedFirstNotCreditCard1: TDBEdit
        Left = 68
        Top = 49
        Width = 101
        Height = 24
        DataField = 'FIRSTNOTCREDITCARD1'
        DataSource = dsrSO120
        TabOrder = 1
      end
    end
    object cobPromote: TComboBox
      Left = 96
      Top = 131
      Width = 138
      Height = 21
      Enabled = False
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 3
    end
    object dedChannelViewDays: TDBEdit
      Left = 455
      Top = 53
      Width = 42
      Height = 21
      DataField = 'ChannelViewDays'
      DataSource = dsrSO120
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      MaxLength = 3
      ParentFont = False
      TabOrder = 5
    end
    object GroupBox2: TGroupBox
      Left = 534
      Top = 4
      Width = 201
      Height = 81
      Caption = '   '#39318#27425' - '#20184#27454#26041#24335'('#25910#36027#21729')'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 7
      object Label9: TLabel
        Left = 23
        Top = 27
        Width = 42
        Height = 13
        Alignment = taRightJustify
        Caption = #20449#29992#21345
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label10: TLabel
        Left = 8
        Top = 52
        Width = 56
        Height = 13
        Alignment = taRightJustify
        Caption = #38750#20449#29992#21345
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object lblUnit3: TLabel
        Left = 174
        Top = 27
        Width = 7
        Height = 13
        Caption = '%'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object lblUnit4: TLabel
        Left = 175
        Top = 56
        Width = 7
        Height = 13
        Caption = '%'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object dedFirstCreditCard2: TDBEdit
        Left = 68
        Top = 19
        Width = 101
        Height = 24
        DataField = 'FIRSTCREDITCARD2'
        DataSource = dsrSO120
        TabOrder = 0
      end
      object dedFirstNotCreditCard2: TDBEdit
        Left = 68
        Top = 49
        Width = 101
        Height = 24
        DataField = 'FIRSTNOTCREDITCARD2'
        DataSource = dsrSO120
        TabOrder = 1
      end
    end
    object GroupBox3: TGroupBox
      Left = 534
      Top = 90
      Width = 200
      Height = 81
      Caption = '   '#32396#25910' - '#20184#27454#26041#24335'('#25910#36027#21729')'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 8
      object Label13: TLabel
        Left = 23
        Top = 27
        Width = 42
        Height = 13
        Alignment = taRightJustify
        Caption = #20449#29992#21345
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label14: TLabel
        Left = 8
        Top = 50
        Width = 56
        Height = 13
        Alignment = taRightJustify
        Caption = #38750#20449#29992#21345
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object lblUnit5: TLabel
        Left = 174
        Top = 27
        Width = 7
        Height = 13
        Caption = '%'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object lblUnit6: TLabel
        Left = 175
        Top = 54
        Width = 7
        Height = 13
        Caption = '%'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object dedOtherCreditCard2: TDBEdit
        Left = 68
        Top = 19
        Width = 101
        Height = 24
        DataField = 'OTHERCREDITCARD2'
        DataSource = dsrSO120
        TabOrder = 0
      end
      object dedOtherNotCreditCard2: TDBEdit
        Left = 68
        Top = 47
        Width = 101
        Height = 24
        DataField = 'OTHERNOTCREDITCARD2'
        DataSource = dsrSO120
        TabOrder = 1
      end
    end
  end
  object pnlButton: TPanel
    Left = 0
    Top = 477
    Width = 788
    Height = 43
    Align = alBottom
    TabOrder = 1
    object btnExit: TSpeedButton
      Left = 660
      Top = 6
      Width = 67
      Height = 30
      Caption = '&O'#32080#26463
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        333333333333333333333333333333333333333FFF33FF333FFF339993370733
        999333777FF37FF377733339993000399933333777F777F77733333399970799
        93333333777F7377733333333999399933333333377737773333333333990993
        3333333333737F73333333333331013333333333333777FF3333333333910193
        333333333337773FF3333333399000993333333337377737FF33333399900099
        93333333773777377FF333399930003999333337773777F777FF339993370733
        9993337773337333777333333333333333333333333333333333333333333333
        3333333333333333333333333333333333333333333333333333}
      NumGlyphs = 2
      ParentFont = False
      OnClick = btnExitClick
    end
    object btnSave: TSpeedButton
      Left = 9
      Top = 6
      Width = 67
      Height = 30
      Caption = '&S'#20786#23384
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333330070
        7700333333337777777733333333008088003333333377F73377333333330088
        88003333333377FFFF7733333333000000003FFFFFFF77777777000000000000
        000077777777777777770FFFFFFF0FFFFFF07F3333337F3333370FFFFFFF0FFF
        FFF07F3FF3FF7FFFFFF70F00F0080CCC9CC07F773773777777770FFFFFFFF039
        99337F3FFFF3F7F777F30F0000F0F09999937F7777373777777F0FFFFFFFF999
        99997F3FF3FFF77777770F00F000003999337F773777773777F30FFFF0FF0339
        99337F3FF7F3733777F30F08F0F0337999337F7737F73F7777330FFFF0039999
        93337FFFF7737777733300000033333333337777773333333333}
      NumGlyphs = 2
      ParentFont = False
      OnClick = btnSaveClick
    end
    object btnCancel: TSpeedButton
      Left = 590
      Top = 6
      Width = 67
      Height = 30
      Caption = '&C'#21462#28040
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        3333333333FFFFF3333333333999993333333333F77777FFF333333999999999
        3333333777333777FF3333993333339993333377FF3333377FF3399993333339
        993337777FF3333377F3393999333333993337F777FF333337FF993399933333
        399377F3777FF333377F993339993333399377F33777FF33377F993333999333
        399377F333777FF3377F993333399933399377F3333777FF377F993333339993
        399377FF3333777FF7733993333339993933373FF3333777F7F3399933333399
        99333773FF3333777733339993333339933333773FFFFFF77333333999999999
        3333333777333777333333333999993333333333377777333333}
      NumGlyphs = 2
      ParentFont = False
      OnClick = btnCancelClick
    end
    object btnDelete: TSpeedButton
      Left = 211
      Top = 6
      Width = 67
      Height = 30
      Caption = '&D'#21034#38500
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        333333333333333333FF33333333333330003333333333333777333333333333
        300033FFFFFF3333377739999993333333333777777F3333333F399999933333
        3300377777733333337733333333333333003333333333333377333333333333
        3333333333333333333F333333333333330033333F33333333773333C3333333
        330033337F3333333377333CC3333333333333F77FFFFFFF3FF33CCCCCCCCCC3
        993337777777777F77F33CCCCCCCCCC399333777777777737733333CC3333333
        333333377F33333333FF3333C333333330003333733333333777333333333333
        3000333333333333377733333333333333333333333333333333}
      NumGlyphs = 2
      ParentFont = False
      OnClick = btnDeleteClick
    end
    object btnEdit: TSpeedButton
      Left = 144
      Top = 6
      Width = 67
      Height = 30
      Caption = '&M'#20462#25913
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000120B0000120B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333000000
        000033333377777777773333330FFFFFFFF03FF3FF7FF33F3FF700300000FF0F
        00F077F777773F737737E00BFBFB0FFFFFF07773333F7F3333F7E0BFBF000FFF
        F0F077F3337773F3F737E0FBFBFBF0F00FF077F3333FF7F77F37E0BFBF00000B
        0FF077F3337777737337E0FBFBFBFBF0FFF077F33FFFFFF73337E0BF0000000F
        FFF077FF777777733FF7000BFB00B0FF00F07773FF77373377373330000B0FFF
        FFF03337777373333FF7333330B0FFFF00003333373733FF777733330B0FF00F
        0FF03333737F37737F373330B00FFFFF0F033337F77F33337F733309030FFFFF
        00333377737FFFFF773333303300000003333337337777777333}
      NumGlyphs = 2
      ParentFont = False
      OnClick = btnEditClick
    end
    object btnInsert: TSpeedButton
      Left = 77
      Top = 6
      Width = 67
      Height = 30
      Caption = '&A'#26032#22686
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -15
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        33333333FF33333333FF333993333333300033377F3333333777333993333333
        300033F77FFF3333377739999993333333333777777F3333333F399999933333
        33003777777333333377333993333333330033377F3333333377333993333333
        3333333773333333333F333333333333330033333333F33333773333333C3333
        330033333337FF3333773333333CC333333333FFFFF77FFF3FF33CCCCCCCCCC3
        993337777777777F77F33CCCCCCCCCC3993337777777777377333333333CC333
        333333333337733333FF3333333C333330003333333733333777333333333333
        3000333333333333377733333333333333333333333333333333}
      NumGlyphs = 2
      ParentFont = False
      OnClick = btnInsertClick
    end
    object lblStatus: TLabel
      Left = 736
      Top = 8
      Width = 40
      Height = 25
      Alignment = taCenter
      AutoSize = False
      Caption = 'lblStatus'
      Color = clYellow
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -16
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
  end
  object pnlGrid: TPanel
    Left = 0
    Top = 185
    Width = 788
    Height = 289
    Align = alClient
    Caption = 'pnlGrid'
    TabOrder = 2
    object StatusBar1: TStatusBar
      Left = 1
      Top = 259
      Width = 786
      Height = 29
      Panels = <
        item
          Width = 50
        end>
      SimplePanel = False
    end
    object dgrMain: TDBGrid
      Left = 1
      Top = 1
      Width = 786
      Height = 258
      Align = alClient
      DataSource = dsrSO120
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      TabOrder = 1
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
    end
  end
  object dsrSO120: TDataSource
    DataSet = dtmMain1.qrySO120
    OnDataChange = dsrSO120DataChange
    Left = 120
    Top = 200
  end
  object dsrCodeCD039: TDataSource
    DataSet = dtmMain1.qryCodeCD039
    Left = 40
    Top = 8
  end
  object dsrCodeCD019: TDataSource
    DataSet = dtmMain1.qryCodeCD019
    Left = 8
    Top = 128
  end
  object dsrCodeCD032: TDataSource
    DataSet = dtmMain1.qryCodeCD032
    Left = 8
    Top = 96
  end
  object dsrCodeCD042: TDataSource
    DataSet = dtmMain1.qryCodeCD042
    Left = 8
    Top = 64
  end
end
