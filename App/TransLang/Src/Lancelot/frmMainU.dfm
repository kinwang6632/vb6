object frmMain: TfrmMain
  Left = 57
  Top = 159
  Width = 544
  Height = 375
  Caption = 'Trans Language'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 16
  object btnTL0020: TBitBtn
    Left = 88
    Top = 40
    Width = 145
    Height = 25
    Caption = 'Get Words Info'
    TabOrder = 0
    OnClick = btnTL0020Click
  end
  object btnExit: TBitBtn
    Left = 360
    Top = 40
    Width = 75
    Height = 25
    Caption = '&OExit'
    TabOrder = 1
    OnClick = btnExitClick
  end
  object btnTL0030: TBitBtn
    Left = 88
    Top = 136
    Width = 145
    Height = 25
    Caption = 'Trans Word Prog'
    TabOrder = 2
    OnClick = btnTL0030Click
  end
  object btnTL0040: TBitBtn
    Left = 88
    Top = 88
    Width = 145
    Height = 25
    Caption = 'Update Words'
    TabOrder = 3
    OnClick = btnTL0040Click
  end
end
