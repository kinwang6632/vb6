object frmLogin: TfrmLogin
  Left = 433
  Top = 219
  Width = 366
  Height = 210
  Caption = 'frmLogin'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object bvlLogin: TBevel
    Left = 0
    Top = 64
    Width = 358
    Height = 119
    Align = alClient
  end
  object lblUserID: TLabel
    Left = 97
    Top = 84
    Width = 60
    Height = 15
    Alignment = taRightJustify
    Caption = #20351#29992#32773#20195#30908
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
    Left = 121
    Top = 116
    Width = 36
    Height = 15
    Alignment = taRightJustify
    Caption = #23494#12288#30908
    Color = clBtnFace
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -12
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentColor = False
    ParentFont = False
  end
  object lblTitle: TLabel
    Left = 0
    Top = 0
    Width = 358
    Height = 57
    Align = alTop
    Alignment = taCenter
    AutoSize = False
    Color = clWhite
    ParentColor = False
  end
  object lblTitle21: TLabel
    Left = 0
    Top = 57
    Width = 358
    Height = 7
    Align = alTop
    AutoSize = False
    Color = clTeal
    ParentColor = False
  end
  object lblTitle1: TLabel
    Left = 80
    Top = 8
    Width = 193
    Height = 17
    Alignment = taCenter
    AutoSize = False
    Caption = 'lblTitle1'
    Color = clWhite
    ParentColor = False
  end
  object lblTitle2: TLabel
    Left = 80
    Top = 32
    Width = 193
    Height = 17
    Alignment = taCenter
    AutoSize = False
    Caption = 'lblTitle1'
    Color = clWhite
    ParentColor = False
  end
  object edtUserID: TEdit
    Left = 171
    Top = 80
    Width = 166
    Height = 21
    TabOrder = 0
  end
  object edtPasswd: TEdit
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
    Caption = #30906#23450'(&Y)'
    Default = True
    TabOrder = 2
    OnClick = btnOKClick
  end
  object btnCancel: TBitBtn
    Left = 272
    Top = 144
    Width = 65
    Height = 25
    Caption = #21462#28040'(&C)'
    ModalResult = 2
    TabOrder = 3
    OnClick = btnCancelClick
  end
end
