object frmSO8F10_3: TfrmSO8F10_3
  Left = 276
  Top = 134
  Width = 307
  Height = 253
  Caption = 'frmSO8F10_3'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 299
    Height = 226
    Align = alClient
    TabOrder = 0
    object cobCompInfo: TComboBox
      Left = 107
      Top = 32
      Width = 145
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnCloseUp = cobCompInfoCloseUp
    end
    object StaticText6: TStaticText
      Left = 40
      Top = 32
      Width = 52
      Height = 24
      Caption = #20844#21496#21029
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
    end
    object btnClose: TButton
      Left = 166
      Top = 84
      Width = 75
      Height = 31
      Caption = #38626#38283
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      OnClick = btnCloseClick
    end
    object btnQuery: TButton
      Left = 62
      Top = 84
      Width = 75
      Height = 31
      Caption = #36039#26009#21516#27493
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      OnClick = btnQueryClick
    end
    object Memo1: TMemo
      Left = 8
      Top = 128
      Width = 281
      Height = 89
      Color = clInfoBk
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Lines.Strings = (
        '')
      ParentFont = False
      TabOrder = 4
    end
  end
end
