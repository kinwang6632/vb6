object frmMainMenu: TfrmMainMenu
  Left = 448
  Top = 231
  Width = 358
  Height = 247
  BorderIcons = []
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poScreenCenter
  WindowState = wsMaximized
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 350
    Height = 200
    Align = alClient
    BevelOuter = bvLowered
    BorderWidth = 3
    Color = clWhite
    Ctl3D = True
    ParentCtl3D = False
    TabOrder = 0
    object Panel2: TPanel
      Left = 4
      Top = 4
      Width = 342
      Height = 30
      Align = alTop
      Alignment = taRightJustify
      BevelWidth = 2
      Caption = #36628#21161#21151#33021'   '
      Color = clMoneyGreen
      Font.Charset = ANSI_CHARSET
      Font.Color = clNavy
      Font.Height = -16
      Font.Name = #32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
    end
    object Panel3: TPanel
      Left = 4
      Top = 184
      Width = 342
      Height = 12
      Align = alBottom
      BevelWidth = 2
      Color = clWhite
      TabOrder = 1
    end
    object Panel4: TPanel
      Left = 4
      Top = 176
      Width = 342
      Height = 8
      Align = alBottom
      Color = clNavy
      TabOrder = 2
    end
    object Panel5: TPanel
      Left = 4
      Top = 172
      Width = 342
      Height = 4
      Align = alBottom
      Color = clMoneyGreen
      TabOrder = 3
    end
  end
  object MainMenu1: TMainMenu
    Left = 72
    Top = 48
    object N1: TMenuItem
      Caption = #20323#29518#37329'&1'
      object btnSO8B30: TMenuItem
        Caption = #20323#29518#37329#20043#29305#27530#21443#25976'&1'
        OnClick = btnSO8B30Click
      end
      object btnSO8B10: TMenuItem
        Caption = #25277#20323#27604#20363'/'#37329#38989#35373#23450'&2'
        OnClick = btnSO8B10Click
      end
      object btnSO8B50: TMenuItem
        Caption = #20449#29992#21345#20184#36027#32396#25910#20323#37329#27512#23660#20154#21729#35373#23450'&3'
        OnClick = btnSO8B50Click
      end
      object btnSO8B40: TMenuItem
        Caption = #37782#24115'&4'
        OnClick = btnSO8B40Click
      end
      object btnSO8B20: TMenuItem
        Caption = #20323#29518#37329#35336#31639'&5'
        OnClick = btnSO8B20Click
      end
    end
    object N2: TMenuItem
      Caption = #25286#24115'&2'
      object btnSO8A10: TMenuItem
        Caption = #25286#24115#20844#24335#20195#30908'&1'
        Visible = False
        OnClick = btnSO8A10Click
      end
      object btnSO8A20: TMenuItem
        Caption = #38971#36947#21830#20195#30908'&2'
        OnClick = btnSO8A20Click
      end
      object btnSO8A30: TMenuItem
        Caption = #38971#36947#25286#24115#20844#24335#35373#23450'&3'
        Visible = False
        OnClick = btnSO8A30Click
      end
      object btnSO8A40: TMenuItem
        Caption = #25286#24115#37329#38989#35336#31639'&4'
        Visible = False
        OnClick = btnSO8A40Click
      end
      object btnSO8A50: TMenuItem
        Caption = #26126#32048#36039#26009#36681'Excel&5'
        OnClick = btnSO8A50Click
      end
      object btnSO8A60: TMenuItem
        Caption = #29256#27402#25104#26412#20998#25892#27604#20363#35373#23450'&6'
        OnClick = btnSO8A60Click
      end
    end
    object N3: TMenuItem
      Caption = #25163#38283#21934'&3'
      object btnSo8C10: TMenuItem
        Caption = #38936#21934#20316#26989'&1'
        OnClick = btnSo8C10Click
      end
      object btnSo8C20: TMenuItem
        Caption = #21934#25818#20570#24290'&2'
        OnClick = btnSo8C20Click
      end
      object btnSo8C30: TMenuItem
        Caption = #22577#34920#21015#21360'&3'
        OnClick = btnSo8C30Click
      end
      object btnSo8C40: TMenuItem
        Caption = #20462#25913#25163#38283#21934#34399'('#24050#26085#32080')&4'
        OnClick = btnSo8C40Click
      end
    end
    object N4: TMenuItem
      Caption = #20108#38542#20195#30908'&4'
      object btnSO8F10: TMenuItem
        Caption = #20108#38542#20195#30908#20998#39006#36039#26009#32173#35703'&1'
        OnClick = btnSO8F10Click
      end
      object btnSO8F20: TMenuItem
        Caption = #20998#39006#28165#21934'&2'
        OnClick = btnSO8F20Click
      end
    end
    object N10: TMenuItem
      Caption = #32080#26463'&5'
      OnClick = N10Click
    end
  end
  object ADOConnection1: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 9
    Top = 9
  end
  object qryPriv: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 48
    Top = 8
  end
  object qryCommon: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 81
    Top = 9
  end
end
