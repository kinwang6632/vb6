object frmInvA06: TfrmInvA06
  Left = 467
  Top = 337
  BorderStyle = bsDialog
  Caption = #26410#38283#30332#31080#36039#26009#26597#35426
  ClientHeight = 154
  ClientWidth = 357
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Label2: TLabel
    Left = 22
    Top = 59
    Width = 60
    Height = 14
    Caption = #25910#36027#26085#26399#65306
    Transparent = True
  end
  object Label1: TLabel
    Left = 182
    Top = 60
    Width = 9
    Height = 14
    Caption = '~'
  end
  object MsgLable: TLabel
    Left = 0
    Top = 131
    Width = 357
    Height = 18
    Align = alBottom
    Alignment = taCenter
    Caption = 'MsgLable'
    Font.Charset = ANSI_CHARSET
    Font.Color = clRed
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Bevel1: TBevel
    Left = 0
    Top = 149
    Width = 357
    Height = 5
    Align = alBottom
    Shape = bsSpacer
  end
  inline fraStartChargeDate: TfraYMD
    Left = 85
    Top = 37
    Width = 94
    Height = 49
    TabOrder = 0
    OnExit = fraStartChargeDateExit
    inherited mseYMD: TMaskEdit
      Top = 16
      Width = 89
      Height = 24
      Font.Charset = ANSI_CHARSET
      Font.Height = -13
      Font.Name = 'Tahoma'
      ParentFont = False
      OnChange = fraStartChargeDatemseYMDChange
    end
  end
  inline fraEndChargeDate: TfraYMD
    Left = 194
    Top = 37
    Width = 94
    Height = 49
    TabOrder = 1
    inherited mseYMD: TMaskEdit
      Left = 1
      Top = 16
      Width = 89
      Height = 24
      Font.Charset = ANSI_CHARSET
      Font.Height = -13
      Font.Name = 'Tahoma'
      ParentFont = False
      OnChange = fraStartChargeDatemseYMDChange
    end
  end
  object btnExit: TBitBtn
    Left = 196
    Top = 100
    Width = 82
    Height = 26
    Caption = #32080#26463
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
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
    Left = 89
    Top = 100
    Width = 82
    Height = 26
    Caption = #26597#35426
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    OnClick = btnQueryClick
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
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 357
    Height = 32
    Align = alTop
    Color = clSkyBlue
    TabOrder = 4
    object Label5: TLabel
      Left = 10
      Top = 8
      Width = 136
      Height = 16
      Caption = #26410#38283#30332#31080#36039#26009#26597#35426
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
end
