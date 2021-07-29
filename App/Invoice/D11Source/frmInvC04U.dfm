object frmInvC04: TfrmInvC04
  Left = 540
  Top = 383
  ActiveControl = fraSInvDate
  BorderStyle = bsDialog
  Caption = 'frmInvC04'
  ClientHeight = 234
  ClientWidth = 289
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 289
    Height = 32
    Align = alTop
    Color = clSkyBlue
    TabOrder = 0
    object Label5: TLabel
      Left = 10
      Top = 8
      Width = 153
      Height = 16
      Caption = #21508#31278#38283#31435#26041#24335#19968#35261#34920
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 32
    Width = 289
    Height = 202
    Align = alClient
    TabOrder = 1
    object Label7: TLabel
      Left = 22
      Top = 68
      Width = 52
      Height = 13
      Alignment = taRightJustify
      Caption = #38283#31435#21029#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object Label2: TLabel
      Left = 11
      Top = 22
      Width = 65
      Height = 13
      Caption = #30332#31080#26085#26399#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object cmbHowToCreate: TComboBox
      Left = 85
      Top = 63
      Width = 178
      Height = 22
      Style = csDropDownList
      ItemHeight = 14
      TabOrder = 2
      Items.Strings = (
        '0-'#20840#37096
        '1-mis'#25291#27284', '#38928#38283
        '2-mis'#25291#27284', '#24460#38283
        '3-'#29694#22580#38283#31435
        '4-'#19968#33324#38283#31435)
    end
    object btnExit: TBitBtn
      Left = 167
      Top = 154
      Width = 73
      Height = 26
      Caption = #32080#26463
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 5
      OnClick = btnExitClick
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
    end
    object btnQuery: TBitBtn
      Left = 63
      Top = 154
      Width = 73
      Height = 26
      Caption = #26597#35426
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 4
      OnClick = btnQueryExClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000120B0000120B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00555555555555
        555555555555555555555555555555555555555555FF55555555555559055555
        55555555577FF5555555555599905555555555557777F5555555555599905555
        555555557777FF5555555559999905555555555777777F555555559999990555
        5555557777777FF5555557990599905555555777757777F55555790555599055
        55557775555777FF5555555555599905555555555557777F5555555555559905
        555555555555777FF5555555555559905555555555555777FF55555555555579
        05555555555555777FF5555555555557905555555555555777FF555555555555
        5990555555555555577755555555555555555555555555555555}
      NumGlyphs = 2
    end
    inline fraEInvDate: TfraYMD
      Left = 179
      Top = 11
      Width = 82
      Height = 41
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      inherited mseYMD: TMaskEdit
        Top = 5
        Width = 80
        Height = 24
      end
    end
    inline fraSInvDate: TfraYMD
      Left = 83
      Top = 6
      Width = 90
      Height = 49
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      OnExit = fraSInvDateExit
      inherited mseYMD: TMaskEdit
        Top = 10
        Width = 80
        Height = 24
      end
    end
    object rgpIsObsolete: TRadioGroup
      Left = 23
      Top = 96
      Width = 242
      Height = 49
      Caption = '  '#26159#21542#20316#24290'  '
      Columns = 3
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ItemIndex = 1
      Items.Strings = (
        #26159
        #21542
        #20840#37096)
      ParentFont = False
      TabOrder = 3
    end
  end
end
