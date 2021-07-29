object frmInvA10: TfrmInvA10
  Left = 576
  Top = 265
  Width = 249
  Height = 181
  Caption = 'frmInvA10'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 241
    Height = 32
    Align = alTop
    Color = clSkyBlue
    TabOrder = 0
    object Label1: TLabel
      Left = 16
      Top = 7
      Width = 102
      Height = 16
      Caption = #30332#31080#37782#24115#20316#26989
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
    Width = 241
    Height = 122
    Align = alClient
    TabOrder = 1
    object Label3: TLabel
      Left = 20
      Top = 23
      Width = 117
      Height = 13
      Caption = #26597#35426#37782#24115#30332#31080#24180#24230#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object btnOK: TBitBtn
      Left = 28
      Top = 67
      Width = 81
      Height = 26
      Caption = #26597#35426
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      OnClick = btnOKClick
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
    object btnExit: TBitBtn
      Left = 140
      Top = 67
      Width = 81
      Height = 26
      Caption = #32080#26463
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 2
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
    object medQueryYear: TMaskEdit
      Left = 148
      Top = 18
      Width = 69
      Height = 21
      EditMask = '9999;1;_'
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      MaxLength = 4
      ParentFont = False
      TabOrder = 0
      Text = '    '
    end
  end
end
