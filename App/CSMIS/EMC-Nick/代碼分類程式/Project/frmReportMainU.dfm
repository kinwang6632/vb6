object frmReportMain: TfrmReportMain
  Left = 321
  Top = 173
  Width = 298
  Height = 180
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
  object cobTablName: TComboBox
    Left = 79
    Top = 71
    Width = 186
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
    Left = 174
    Top = 108
    Width = 59
    Height = 28
    Caption = #38626#38283
    TabOrder = 2
    OnClick = btnCloseClick
  end
  object btnQuery: TButton
    Left = 78
    Top = 108
    Width = 59
    Height = 28
    Caption = #26597#35426
    TabOrder = 1
    OnClick = btnQueryClick
  end
  object StaticText6: TStaticText
    Left = 30
    Top = 24
    Width = 40
    Height = 20
    Caption = #20844#21496#21029
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clNavy
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
  end
  object cobCompInfo: TComboBox
    Left = 81
    Top = 24
    Width = 184
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 4
  end
  object StaticText1: TStaticText
    Left = 14
    Top = 72
    Width = 64
    Height = 20
    Caption = #20998#39006#20195#30908#27284
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clNavy
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 5
  end
  object dsrCodeCD067A: TDataSource
    DataSet = dtmMain.adoCodeCD067A
    Left = 8
    Top = 88
  end
end
