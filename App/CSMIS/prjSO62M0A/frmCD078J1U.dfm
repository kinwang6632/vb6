object frmCD078JU: TfrmCD078JU
  Left = 343
  Top = 243
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'frmCD078JU'
  ClientHeight = 161
  ClientWidth = 383
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 383
    Height = 119
    Align = alTop
    TabOrder = 0
    object Label2: TLabel
      Left = 9
      Top = 10
      Width = 48
      Height = 14
      Caption = #26381#21209#39006#21029
      Font.Charset = ANSI_CHARSET
      Font.Color = clRed
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label1: TLabel
      Left = 9
      Top = 37
      Width = 48
      Height = 14
      Caption = #29986#21697#21517#31281
      Font.Charset = ANSI_CHARSET
      Font.Color = clRed
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label3: TLabel
      Left = 9
      Top = 62
      Width = 48
      Height = 14
      Caption = #35373#20633#38918#24207
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    inline ServiceType: TDataLookup
      Left = 63
      Top = 6
      Width = 175
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 0
      inherited CodeNo: TcxTextEdit
        Properties.CharCase = ecUpperCase
        Properties.MaxLength = 1
        Properties.OnChange = ServiceTypeCodeNoPropertiesChange
      end
      inherited CodeName: TcxLookupComboBox
        Left = 61
        Properties.OnChange = ServiceTypeCodeNamePropertiesChange
        Properties.OnInitPopup = ServiceTypeCodeNamePropertiesInitPopup
        Width = 110
      end
    end
    inline ProductItem: TDataLookup
      Left = 63
      Top = 33
      Width = 300
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 1
      inherited CodeNo: TcxTextEdit
        Properties.MaxLength = 5
        Properties.OnChange = ProductItemCodeNoPropertiesChange
      end
      inherited CodeName: TcxLookupComboBox
        Left = 61
        Properties.OnChange = ProductItemCodeNamePropertiesChange
        Properties.OnInitPopup = ProductItemCodeNamePropertiesInitPopup
        Width = 235
      end
    end
    object btnInstCodeStr: TButton
      Left = 8
      Top = 87
      Width = 93
      Height = 25
      Caption = #25351#23450#27966#24037#39006#21029
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      OnClick = btnInstCodeStrClick
    end
    object cxButton1: TcxButton
      Left = 103
      Top = 87
      Width = 25
      Height = 25
      TabOrder = 3
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
    object edtInstCodeStr: TcxTextEdit
      Left = 130
      Top = 88
      Enabled = False
      ParentFont = False
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 4
      Width = 231
    end
    object edtFaciItem: TcxTextEdit
      Left = 63
      Top = 60
      ParentFont = False
      Properties.Alignment.Horz = taRightJustify
      Properties.MaxLength = 2
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 5
      Width = 69
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 119
    Width = 383
    Height = 42
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 1
    object Bevel1: TBevel
      Left = 0
      Top = 0
      Width = 383
      Height = 3
      Align = alTop
      Shape = bsTopLine
    end
    object edtDML: TcxTextEdit
      Left = 309
      Top = 9
      Enabled = False
      ParentFont = False
      Properties.Alignment.Horz = taCenter
      Properties.ReadOnly = True
      Style.Color = 16777111
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.StyleController = StyleController.EditorStyle
      Style.IsFontAssigned = True
      StyleDisabled.Color = 16777111
      StyleDisabled.TextColor = clWindowText
      TabOrder = 0
      Width = 67
    end
    object btnSave: TButton
      Left = 11
      Top = 10
      Width = 70
      Height = 26
      Action = actSave
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
    end
    object btnCancel: TButton
      Left = 87
      Top = 10
      Width = 70
      Height = 26
      Action = actCancel
      Cancel = True
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
    end
  end
  object ActionList1: TActionList
    Left = 250
    Top = 58
    object actSave: TAction
      Caption = 'F2.'#23384#27284
      ShortCut = 113
      OnExecute = actSaveExecute
    end
    object actCancel: TAction
      Caption = #21462#28040'(&X)'
      ShortCut = 32856
      OnExecute = actCancelExecute
    end
  end
  object CloneDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 200
    Top = 64
  end
end
