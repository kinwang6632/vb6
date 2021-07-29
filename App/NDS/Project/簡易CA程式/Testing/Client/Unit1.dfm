object Form1: TForm1
  Left = 192
  Top = 107
  Width = 340
  Height = 422
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 49
    Top = 52
    Width = 32
    Height = 13
    Caption = 'Data : '
  end
  object Label2: TLabel
    Left = 26
    Top = 16
    Width = 59
    Height = 13
    Caption = 'Remote IP : '
  end
  object Button1: TButton
    Left = 224
    Top = 50
    Width = 75
    Height = 25
    Caption = 'Send'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Edit1: TEdit
    Left = 87
    Top = 51
    Width = 121
    Height = 21
    TabOrder = 1
    Text = 'Edit1'
  end
  object Button2: TButton
    Left = 224
    Top = 16
    Width = 75
    Height = 25
    Caption = 'Connect'
    TabOrder = 2
    OnClick = Button2Click
  end
  object Edit2: TEdit
    Left = 88
    Top = 16
    Width = 121
    Height = 21
    TabOrder = 3
    Text = '127.0.0.1'
  end
  object Memo1: TMemo
    Left = 8
    Top = 80
    Width = 313
    Height = 297
    ScrollBars = ssBoth
    TabOrder = 4
  end
end
