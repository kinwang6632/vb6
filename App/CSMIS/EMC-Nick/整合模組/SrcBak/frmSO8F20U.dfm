object frmSO8F20: TfrmSO8F20
  Left = 321
  Top = 155
  Width = 298
  Height = 198
  Caption = 'SO8F20'
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
  object Label1: TLabel
    Left = 212
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
  object cobTablName: TComboBox
    Left = 95
    Top = 87
    Width = 170
    Height = 21
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = #32048#26126#39636
    Font.Style = []
    ItemHeight = 13
    ParentFont = False
    TabOrder = 0
  end
  object btnClose: TButton
    Left = 166
    Top = 124
    Width = 75
    Height = 31
    Caption = #38626#38283
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    OnClick = btnCloseClick
  end
  object btnQuery: TButton
    Left = 54
    Top = 124
    Width = 75
    Height = 31
    Caption = #26597#35426
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    OnClick = btnQueryClick
  end
  object StaticText6: TStaticText
    Left = 38
    Top = 40
    Width = 52
    Height = 24
    Alignment = taRightJustify
    Caption = #20844#21496#21029
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clNavy
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
  end
  object cobCompInfo: TComboBox
    Left = 97
    Top = 40
    Width = 168
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 4
  end
  object StaticText1: TStaticText
    Left = 6
    Top = 88
    Width = 84
    Height = 24
    Alignment = taRightJustify
    Caption = #20998#39006#20195#30908#27284
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clNavy
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 5
  end
  object dsrCodeCD067A: TDataSource
    DataSet = dtmMain4.adoCodeCD067A
    Left = 8
    Top = 56
  end
end
