object frmCD078BA: TfrmCD078BA
  Left = 413
  Top = 303
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #19978#20659#33267'FTP [frmCD078BA]'
  ClientHeight = 97
  ClientWidth = 554
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
  object pnl1: TPanel
    Left = 0
    Top = 0
    Width = 554
    Height = 97
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object btnCompCodeStr: TButton
      Left = 11
      Top = 15
      Width = 89
      Height = 25
      Caption = #20844#21496#21029
      TabOrder = 0
      OnClick = btnCompCodeStrClick
    end
    object cxButton2: TcxButton
      Left = 102
      Top = 15
      Width = 25
      Height = 25
      TabOrder = 1
      Glyph.Data = {
        72020000424D720200000000000036000000280000000E0000000D0000000100
        1800000000003C020000C40E0000C40E00000000000000000000C8D0D4C8D0D4
        C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000
        80FFFFFF0000C8D0D4C8D0D4000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0
        D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4000080000080000080FFFF
        FFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FFFFFFC8D0D40000C8D0
        D4000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FF
        FFFFC8D0D4C8D0D40000C8D0D4C8D0D4000080000080000080FFFFFFC8D0D4C8
        D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D40000C8D0D4C8D0D4C8D0D400
        0080000080000080FFFFFFC8D0D4000080000080FFFFFFC8D0D4C8D0D4C8D0D4
        0000C8D0D4C8D0D4C8D0D4C8D0D4000080000080000080000080000080FFFFFF
        C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080
        000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4
        C8D0D4C8D0D4000080000080000080000080000080FFFFFFC8D0D4C8D0D4C8D0
        D4C8D0D40000C8D0D4C8D0D4C8D0D4000080000080000080FFFFFFC8D0D40000
        80FFFFFFC8D0D4C8D0D4C8D0D4C8D0D40000C8D0D40000800000800000800000
        80FFFFFFC8D0D4C8D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D400000000
        80000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D400008000
        0080FFFFFFC8D0D40000000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8
        D0D4C8D0D4C8D0D4C8D0D4000080000080FFFFFF0000}
      LookAndFeel.Kind = lfStandard
    end
    object edtCompCodeStr: TcxTextEdit
      Left = 131
      Top = 17
      Enabled = False
      ParentFont = False
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 2
      Width = 409
    end
    object pnl2: TPanel
      Left = 0
      Top = 56
      Width = 554
      Height = 41
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 3
      object bvl1: TBevel
        Left = 0
        Top = 0
        Width = 554
        Height = 3
        Align = alTop
        Shape = bsTopLine
      end
      object btnSave: TButton
        Left = 159
        Top = 9
        Width = 82
        Height = 26
        Caption = #30906#23450
        TabOrder = 0
        OnClick = btnSaveClick
      end
      object btnCancel: TButton
        Left = 267
        Top = 9
        Width = 82
        Height = 26
        Cancel = True
        Caption = #21462#28040
        ModalResult = 2
        TabOrder = 1
        OnClick = btnCancelClick
      end
    end
  end
  object conDataConnection: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 428
    Top = 32
  end
  object adoDataReader: TADOQuery
    Connection = conDataConnection
    Parameters = <>
    Left = 465
    Top = 32
  end
  object adoDataWriter: TADOQuery
    Connection = conDataConnection
    Parameters = <>
    Left = 505
    Top = 32
  end
  object IdFTP1: TIdFTP
    MaxLineAction = maException
    ReadTimeout = 0
    ProxySettings.ProxyType = fpcmNone
    ProxySettings.Port = 0
    Left = 396
    Top = 30
  end
end
