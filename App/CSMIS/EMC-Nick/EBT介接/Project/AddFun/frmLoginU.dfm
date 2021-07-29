object frmLogin: TfrmLogin
  Left = 263
  Top = 228
  Width = 289
  Height = 262
  BorderIcons = []
  Caption = #20351#29992#32773#30331#20837
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 16
  object lblTitle2: TLabel
    Left = 0
    Top = 65
    Width = 281
    Height = 7
    Align = alTop
    AutoSize = False
    Color = clTeal
    ParentColor = False
  end
  object lblTitle1: TLabel
    Left = 0
    Top = 0
    Width = 281
    Height = 65
    Align = alTop
    Alignment = taCenter
    AutoSize = False
    Caption = 'CSIS  V1.0'
    Color = clWhite
    ParentColor = False
  end
  object StaticText1: TStaticText
    Left = 40
    Top = 91
    Width = 64
    Height = 20
    Caption = #20351#29992#32773#24115#34399
    TabOrder = 0
  end
  object StaticText2: TStaticText
    Left = 40
    Top = 123
    Width = 64
    Height = 20
    Caption = #20351#29992#32773#23494#30908
    TabOrder = 1
  end
  object edtUserID: TEdit
    Left = 112
    Top = 91
    Width = 121
    Height = 24
    TabOrder = 2
  end
  object edtUserPasswd: TEdit
    Left = 112
    Top = 123
    Width = 121
    Height = 24
    PasswordChar = '*'
    TabOrder = 3
  end
  object BitBtn1: TBitBtn
    Left = 53
    Top = 187
    Width = 75
    Height = 25
    Caption = #30331#20837
    Default = True
    TabOrder = 4
    OnClick = BitBtn1Click
  end
  object BitBtn2: TBitBtn
    Left = 149
    Top = 187
    Width = 75
    Height = 25
    Caption = #21462#28040
    ModalResult = 2
    TabOrder = 5
  end
end
