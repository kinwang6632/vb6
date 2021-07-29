object frmSelectPath: TfrmSelectPath
  Left = 709
  Top = 278
  Width = 274
  Height = 241
  BorderIcons = []
  Caption = #36984#21462#36039#26009#22846
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object lblPath: TLabel
    Left = 16
    Top = 14
    Width = 232
    Height = 13
    Caption = 'F:\APP\Invoice\Delphi-Prog\WinVersion'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlue
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object DriveComboBox1: TDriveComboBox
    Left = 16
    Top = 40
    Width = 233
    Height = 22
    DirList = DirectoryListBox1
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
  end
  object DirectoryListBox1: TDirectoryListBox
    Left = 16
    Top = 72
    Width = 233
    Height = 97
    DirLabel = lblPath
    ItemHeight = 16
    TabOrder = 1
  end
  object btnOK: TButton
    Left = 24
    Top = 176
    Width = 91
    Height = 25
    Caption = #36984#21462
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = #26032#32048#26126#39636
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
    OnClick = btnOKClick
  end
  object btnCancel: TButton
    Left = 144
    Top = 176
    Width = 91
    Height = 25
    Caption = #21462#28040
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = #26032#32048#26126#39636
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 3
    OnClick = btnCancelClick
  end
end
