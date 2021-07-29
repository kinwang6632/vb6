object fmSQL: TfmSQL
  Left = 362
  Top = 201
  Width = 543
  Height = 387
  Caption = 'SQL'
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 319
    Width = 535
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      535
      41)
    object btnOK: TcxButton
      Left = 425
      Top = 8
      Width = 90
      Height = 25
      Anchors = [akRight, akBottom]
      Caption = #30906#23450
      TabOrder = 0
      OnClick = btnOKClick
      LookAndFeel.Kind = lfOffice11
      LookAndFeel.NativeStyle = False
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 535
    Height = 319
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 3
    TabOrder = 1
    object txtSQL: TSynEdit
      Left = 3
      Top = 3
      Width = 529
      Height = 313
      Align = alClient
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Courier New'
      Font.Style = []
      TabOrder = 0
      Gutter.Font.Charset = DEFAULT_CHARSET
      Gutter.Font.Color = clWindowText
      Gutter.Font.Height = -11
      Gutter.Font.Name = 'Courier New'
      Gutter.Font.Style = []
      Gutter.ShowLineNumbers = True
      Highlighter = SynSQLSyn1
      Lines.Strings = (
        'txtSQL')
    end
  end
  object SynSQLSyn1: TSynSQLSyn
    CommentAttri.Foreground = clGreen
    IdentifierAttri.Foreground = clBlue
    StringAttri.Foreground = clRed
    SQLDialect = sqlOracle
    Left = 56
    Top = 64
  end
end
