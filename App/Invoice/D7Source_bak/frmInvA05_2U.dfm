object frmInvA05_2: TfrmInvA05_2
  Left = 461
  Top = 228
  BorderStyle = bsDialog
  Caption = #23384#27284#36335#24465#65306
  ClientHeight = 326
  ClientWidth = 308
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object SpeedButton1: TSpeedButton
    Left = 269
    Top = 88
    Width = 24
    Height = 21
    Caption = '...'
    OnClick = SpeedButton1Click
  end
  object edtFilePath: TEdit
    Left = 12
    Top = 88
    Width = 258
    Height = 22
    TabOrder = 0
  end
  object btnEnter: TButton
    Left = 71
    Top = 289
    Width = 75
    Height = 25
    Caption = #22519#34892
    TabOrder = 1
    OnClick = btnEnterClick
  end
  object btnExit: TButton
    Left = 167
    Top = 289
    Width = 75
    Height = 25
    Caption = #38626#38283
    TabOrder = 2
    OnClick = btnExitClick
  end
  object ProgressBar1: TProgressBar
    Left = 12
    Top = 120
    Width = 281
    Height = 17
    Step = 1
    TabOrder = 3
  end
  object rgpInvformat: TRadioGroup
    Left = 12
    Top = 40
    Width = 282
    Height = 41
    Caption = ' '#36681#27284#26684#24335' '
    Columns = 2
    Font.Charset = ANSI_CHARSET
    Font.Color = clTeal
    Font.Height = -13
    Font.Name = #26032#32048#26126#39636
    Font.Style = []
    ItemIndex = 0
    Items.Strings = (
      #20840#37096
      #19981#21547#29694#22580#38283#31435)
    ParentFont = False
    TabOrder = 4
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 308
    Height = 32
    Align = alTop
    Color = clSkyBlue
    TabOrder = 5
    object Label1: TLabel
      Left = 16
      Top = 7
      Width = 136
      Height = 16
      Caption = #36681#25991#23383#27284#23384#27284#36335#24465
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object ListBox1: TListBox
    Left = 12
    Top = 150
    Width = 282
    Height = 131
    ItemHeight = 14
    TabOrder = 6
  end
  object OpenDialog1: TOpenDialog
    Filter = #25991#23383#27284#26696'(*.txt)|*.txt|'#25152#26377#27284#26696'(*.*)|*.*'
    Left = 60
    Top = 204
  end
end
