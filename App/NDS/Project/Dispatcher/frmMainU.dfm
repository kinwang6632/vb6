object frmMain: TfrmMain
  Left = 192
  Top = 158
  Width = 277
  Height = 210
  BorderIcons = []
  Caption = 'NDS Dispatcher '#20027#21151#33021#30059#38754
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 269
    Height = 164
    Align = alClient
    BevelOuter = bvSpace
    BorderStyle = bsSingle
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
    object Panel2: TPanel
      Left = 1
      Top = 1
      Width = 263
      Height = 41
      Align = alTop
      Caption = 'NDS Dispatcher'
      Color = clMoneyGreen
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -19
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
    end
  end
  object MainMenu1: TMainMenu
    Left = 72
    Top = 64
    object N1: TMenuItem
      Caption = #31995#32113#35373#23450
      object N5: TMenuItem
        Caption = #31995#32113#21443#25976#35373#23450
        OnClick = N5Click
      end
      object DB1: TMenuItem
        Caption = 'DB'#36899#32218#35373#23450
        OnClick = DB1Click
      end
    end
    object N2: TMenuItem
      Caption = #20998#27966#25351#20196
      OnClick = N2Click
    end
    object N4: TMenuItem
      Caption = #35498#26126
      object AboutSMSGateway1: TMenuItem
        Caption = 'About SMS Gateway'
        OnClick = AboutSMSGateway1Click
      end
    end
    object N3: TMenuItem
      Caption = #38626#38283
      OnClick = N3Click
    end
  end
end
