object frmLogin: TfrmLogin
  Left = 227
  Top = 107
  Width = 372
  Height = 207
  Caption = 'Login'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  PixelsPerInch = 96
  TextHeight = 13
  object bvlLogin: TBevel
    Left = 0
    Top = 72
    Width = 364
    Height = 108
    Align = alClient
  end
  object lblUserID: TLabel
    Left = 123
    Top = 84
    Width = 39
    Height = 15
    Caption = 'User ID'
    Color = clBtnFace
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -12
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentColor = False
    ParentFont = False
  end
  object lblPassword: TLabel
    Left = 117
    Top = 116
    Width = 47
    Height = 15
    Caption = 'Password'
    Color = clBtnFace
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -12
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentColor = False
    ParentFont = False
  end
  object lblTitle1: TLabel
    Left = 0
    Top = 0
    Width = 364
    Height = 65
    Align = alTop
    Alignment = taCenter
    AutoSize = False
    Caption = '   CA Command Gateway for Nagra V2.0'
    Color = clWhite
    ParentColor = False
  end
  object lblTitle2: TLabel
    Left = 0
    Top = 65
    Width = 364
    Height = 7
    Align = alTop
    AutoSize = False
    Color = clTeal
    ParentColor = False
  end
  object edtUserID: TEdit
    Left = 171
    Top = 80
    Width = 166
    Height = 21
    TabOrder = 0
  end
  object edtPassword: TEdit
    Left = 171
    Top = 112
    Width = 166
    Height = 21
    PasswordChar = '*'
    TabOrder = 1
  end
  object btnOK: TBitBtn
    Left = 192
    Top = 144
    Width = 65
    Height = 25
    Caption = 'Confirm'
    Default = True
    TabOrder = 2
    OnClick = btnOKClick
  end
  object btnCancel: TBitBtn
    Left = 272
    Top = 144
    Width = 65
    Height = 25
    Caption = 'Cancel'
    ModalResult = 2
    TabOrder = 3
  end
end
