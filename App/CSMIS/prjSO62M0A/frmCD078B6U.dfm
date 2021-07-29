object frmCD078B6: TfrmCD078B6
  Left = 172
  Top = 123
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'frmCD078B6'
  ClientHeight = 402
  ClientWidth = 564
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  Icon.Data = {
    0000010002002020100000000000E80200002600000010101000000000002801
    00000E0300002800000020000000400000000100040000000000800200000000
    0000000000000000000000000000000000000000800000800000008080008000
    0000800080008080000080808000C0C0C0000000FF0000FF000000FFFF00FF00
    0000FF00FF00FFFF0000FFFFFF00000000000000000000000000000000000000
    00000000BBBBBBBB000000000000000000000BBBBBBBBBBBBBB0000000000000
    0000BBBBBBBBBBBBBBBB00000000000000BBBBBBBBBBBBBBBBBBBB0000000000
    0BBBBBBBBBBBBBBBBBBBBBB000000000BBBBBBBBB0000000BBBBBBBB00000000
    BBBBBBB00BBBBBBB00BBBBBB0000000BBBBBBB00BBBBBBBBB00BBBBBB00000BB
    BBBBB00BBBBBBBBBBB00BBBBBB0000BBBBBBB0BBBBBBBBBBBBB0BBBBBB0000BB
    BBBB0BBBBBBBBBBBBBBB0BBBBB000BBBBBBB0BBBBBBBBBBBBBBB0BBBBBB00BBB
    BBBB0BBBBBBBBBBBBBBB0BBBBBB00BBBBBBBBBBBBBBBBBBBBBBBBBBBBBB00BBB
    BBBBBBBBBBBBBBBBBBBBBBBBBBB00BBBBBBBBBBBBBBBBBBBBBBBBBBBBBB00BBB
    BBBBBBB00BBBBBB00BBBBBBBBBB00BBBBBBBBB0000BBBB0000BBBBBBBBB00BBB
    BBBBBB0000BBBB0000BBBBBBBBB000BBBBBBBB0000BBBB0000BBBBBBBB0000BB
    BBBBBB0000BBBB0000BBBBBBBB0000BBBBBBBB0000BBBB0000BBBBBBBB00000B
    BBBBBBB00BBBBBB00BBBBBBBB0000000BBBBBBBBBBBBBBBBBBBBBBBB00000000
    BBBBBBBBBBBBBBBBBBBBBBBB000000000BBBBBBBBBBBBBBBBBBBBBB000000000
    00BBBBBBBBBBBBBBBBBBBB00000000000000BBBBBBBBBBBBBBBB000000000000
    00000BBBBBBBBBBBBBB000000000000000000000BBBBBBBB0000000000000000
    0000000000000000000000000000FFF00FFFFF8001FFFE00007FFC00003FF800
    001FF000000FE0000007C0000003C00000038000000180000001800000010000
    0000000000000000000000000000000000000000000000000000000000008000
    00018000000180000001C0000003C0000003E0000007F000000FF800001FFC00
    003FFE00007FFF8001FFFFF00FFF280000001000000020000000010004000000
    0000C00000000000000000000000000000000000000000000000000080000080
    00000080800080000000800080008080000080808000C0C0C0000000FF0000FF
    000000FFFF00FF000000FF00FF00FFFF0000FFFFFF0000000000000000000000
    0BBBBBB00000000BBBBBBBBBB00000BBB000003BBB0000BB00BBBB03BB000BBB
    0BBBBBB0BBB00BBB0BBBBBB0BBB00BBBBBBBBBBBBBB00BBBB00BB00BBBB00BBB
    B00BB00BBBB00BBBB00BB00BBBB000BBB00BB00BBB0000BBBBBBBBBBBB00000B
    BBBBBBBBB00000000BBBBBB000000000000000000000F81F0000E0070000C003
    0000840100008801000000000000000000000000000000000000000000000000
    00008001000080010000C0030000E0070000F81F0000}
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 370
    Width = 564
    Height = 32
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      564
      32)
    object edtDML: TcxTextEdit
      Left = 603
      Top = -4
      Anchors = [akRight, akBottom]
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
      Left = 78
      Top = 3
      Width = 69
      Height = 26
      Action = actSave
      TabOrder = 1
    end
    object Button1: TButton
      Left = 158
      Top = 3
      Width = 69
      Height = 26
      Action = actCancel
      Cancel = True
      TabOrder = 2
    end
    object btnUpdate: TButton
      Left = 6
      Top = 3
      Width = 69
      Height = 26
      Action = actUpdate
      TabOrder = 3
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 0
    Width = 564
    Height = 209
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    object Label3: TLabel
      Left = 7
      Top = 58
      Width = 48
      Height = 15
      Caption = #25351#23450#35373#20633
      Font.Charset = ANSI_CHARSET
      Font.Color = clRed
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label1: TLabel
      Left = 7
      Top = 33
      Width = 48
      Height = 14
      Caption = #25910#36027#38917#30446
      Font.Charset = ANSI_CHARSET
      Font.Color = clRed
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 7
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
    object Bevel2: TBevel
      Left = 0
      Top = 208
      Width = 564
      Height = 1
      Align = alBottom
      Shape = bsBottomLine
    end
    object Label4: TLabel
      Left = 7
      Top = 86
      Width = 48
      Height = 14
      Caption = #40670#25976#36774#27861
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    inline FaciItem: TDataLookup
      Left = 58
      Top = 56
      Width = 211
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 2
      inherited CodeNo: TcxTextEdit
        Left = 1
        Properties.MaxLength = 2
        Properties.OnChange = FaciItemCodeNoPropertiesChange
      end
      inherited CodeName: TcxLookupComboBox
        Properties.OnChange = FaciItemCodeNamePropertiesChange
        Properties.OnInitPopup = FaciItemCodeNamePropertiesInitPopup
        Width = 148
      end
    end
    inline CItem: TDataLookup
      Left = 59
      Top = 30
      Width = 300
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 1
      inherited CodeNo: TcxTextEdit
        Properties.MaxLength = 5
        Properties.OnChange = CItemCodeNoPropertiesChange
        OnExit = CItemCodeNoExit
      end
      inherited CodeName: TcxLookupComboBox
        Left = 61
        Properties.OnChange = CItemCodeNamePropertiesChange
        Properties.OnInitPopup = CItemCodeNamePropertiesInitPopup
        Width = 235
      end
    end
    inline ServiceType: TDataLookup
      Left = 59
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
    object btnPrClass: TButton
      Left = 9
      Top = 111
      Width = 89
      Height = 25
      Caption = #25351#23450#27966#24037#39006#21029
      Font.Charset = ANSI_CHARSET
      Font.Color = clRed
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 3
    end
    object cxButton1: TcxButton
      Left = 100
      Top = 112
      Width = 25
      Height = 25
      TabOrder = 4
      OnClick = cxButton1Click
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
    object edtPrClassCodeStr: TcxTextEdit
      Left = 128
      Top = 112
      Enabled = False
      ParentFont = False
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 5
      Width = 409
    end
    inline DataLookup1: TDataLookup
      Left = 59
      Top = 82
      Width = 298
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 6
      inherited CodeNo: TcxTextEdit
        Properties.MaxLength = 5
      end
      inherited CodeName: TcxLookupComboBox
        Left = 61
        Width = 236
      end
    end
  end
  object Panel4: TPanel
    Left = 0
    Top = 209
    Width = 564
    Height = 161
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 2
    object Bevel1: TBevel
      Left = 0
      Top = 207
      Width = 564
      Height = 3
      Align = alBottom
      Shape = bsBottomLine
    end
  end
  object ActionList1: TActionList
    Left = 32
    Top = 258
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
    object actUpdate: TAction
      Caption = 'F11.'#20462#25913
      ShortCut = 122
      OnExecute = actUpdateExecute
    end
  end
  object dsCD078A3: TDataSource
    Left = 76
    Top = 270
  end
end
