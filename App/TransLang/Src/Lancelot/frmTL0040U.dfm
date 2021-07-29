object frmTL0040: TfrmTL0040
  Left = 197
  Top = 142
  Width = 447
  Height = 310
  Caption = '譯文轉入程式 ( TL0040 )'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 16
  object Label1: TLabel
    Left = 24
    Top = 16
    Width = 401
    Height = 20
    AutoSize = False
    Caption = '※ 執行此功能前，請確定翻譯完成之譯文 ※'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clNavy
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label2: TLabel
    Left = 100
    Top = 56
    Width = 285
    Height = 20
    AutoSize = False
    Caption = '※ 已存放在RWordsInfo中 ※ '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clNavy
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label3: TLabel
    Left = 16
    Top = 232
    Width = 97
    Height = 16
    AutoSize = False
    Caption = '轉入總筆數 : '
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clTeal
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object lblCount: TLabel
    Left = 128
    Top = 232
    Width = 121
    Height = 16
    AutoSize = False
    Caption = '0'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clTeal
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object rgpLang: TRadioGroup
    Left = 16
    Top = 91
    Width = 409
    Height = 49
    Caption = ' Language '
    Columns = 4
    ItemIndex = 0
    Items.Strings = (
      'English '
      'China'
      'Japanese'
      'All')
    TabOrder = 0
  end
  object BitBtn1: TBitBtn
    Left = 88
    Top = 163
    Width = 123
    Height = 25
    Caption = '&S Start'
    TabOrder = 1
    OnClick = BitBtn1Click
  end
  object BitBtn2: TBitBtn
    Left = 240
    Top = 163
    Width = 123
    Height = 25
    Caption = '&OExit'
    TabOrder = 2
    OnClick = BitBtn2Click
  end
  object pgbStatus: TProgressBar
    Left = 0
    Top = 267
    Width = 439
    Height = 16
    Align = alBottom
    Min = 0
    Max = 100
    TabOrder = 3
  end
end
