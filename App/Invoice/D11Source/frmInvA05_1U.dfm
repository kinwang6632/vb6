object frmInvA05_1: TfrmInvA05_1
  Left = 479
  Top = 356
  Width = 329
  Height = 198
  Caption = #21015#21360#30332#31080#22411#24335
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 321
    Height = 170
    Align = alClient
    TabOrder = 0
    object btnExit: TBitBtn
      Left = 178
      Top = 130
      Width = 81
      Height = 26
      Caption = #32080#26463
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
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
    object btnPrintNew: TBitBtn
      Left = 56
      Top = 130
      Width = 81
      Height = 26
      Caption = #21015#21360
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      OnClick = Button1Click
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00300000000000
        00033FFFFFFFFFFFFFFF0888888888888880777777777777777F088888888888
        8880777777777777777F0000000000000000FFFFFFFFFFFFFFFF0F8F8F8F8F8F
        8F80777777777777777F08F8F8F8F8F8F9F0777777777777777F0F8F8F8F8F8F
        8F807777777777777F7F0000000000000000777777777777777F3330FFFFFFFF
        03333337F3FFFF3F7F333330F0000F0F03333337F77773737F333330FFFFFFFF
        03333337F3FF3FFF7F333330F00F000003333337F773777773333330FFFF0FF0
        33333337F3FF7F3733333330F08F0F0333333337F7737F7333333330FFFF0033
        33333337FFFF7733333333300000033333333337777773333333}
      NumGlyphs = 2
    end
    object rgpPrintType: TRadioGroup
      Left = 13
      Top = 51
      Width = 293
      Height = 62
      Caption = #30332#31080#26684#24335' '
      Columns = 3
      Font.Charset = ANSI_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ItemIndex = 0
      Items.Strings = (
        'A4'#26684#24335'('#19968')'
        #20013#19968#20992#26684#24335
        'A4'#26684#24335'('#20108')')
      ParentFont = False
      TabOrder = 0
    end
    object Panel2: TPanel
      Left = 1
      Top = 1
      Width = 319
      Height = 32
      Align = alTop
      Color = clSkyBlue
      TabOrder = 3
      object Label1: TLabel
        Left = 16
        Top = 7
        Width = 102
        Height = 16
        Caption = #21015#21360#30332#31080#22411#24335
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlue
        Font.Height = -16
        Font.Name = #26032#32048#26126#39636
        Font.Style = [fsBold]
        ParentFont = False
      end
    end
  end
end
