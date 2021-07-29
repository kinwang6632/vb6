object frmCD078B5: TfrmCD078B5
  Left = 213
  Top = 196
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'frmCD078B5'
  ClientHeight = 330
  ClientWidth = 505
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
    Top = 282
    Width = 505
    Height = 48
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 1
    object Bevel1: TBevel
      Left = 0
      Top = 0
      Width = 505
      Height = 3
      Align = alTop
      Shape = bsTopLine
    end
    object edtDML: TcxTextEdit
      Left = 396
      Top = 14
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
      Top = 12
      Width = 70
      Height = 26
      Action = actSave
      TabOrder = 1
    end
    object btnCancel: TButton
      Left = 87
      Top = 12
      Width = 70
      Height = 26
      Action = actCancel
      Cancel = True
      TabOrder = 2
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 505
    Height = 282
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object Label1: TLabel
      Left = 95
      Top = 50
      Width = 48
      Height = 14
      Caption = #25910#36027#38917#30446
    end
    object Label2: TLabel
      Left = 95
      Top = 24
      Width = 48
      Height = 14
      Caption = #26381#21209#39006#21029
    end
    object Label3: TLabel
      Left = 95
      Top = 76
      Width = 48
      Height = 14
      Caption = #25351#23450#35373#20633
    end
    object Label5: TLabel
      Left = 95
      Top = 105
      Width = 48
      Height = 14
      Caption = #37559#21806#29260#20729
    end
    object Label4: TLabel
      Left = 95
      Top = 130
      Width = 48
      Height = 14
      Caption = #20778#24800#37329#38989
    end
    object Label6: TLabel
      Left = 49
      Top = 180
      Width = 96
      Height = 14
      Caption = #36949#32004#26178#20043#35336#20729#20381#25818
    end
    object Label9: TLabel
      Left = 95
      Top = 156
      Width = 48
      Height = 14
      Caption = #36949#32004#34389#20221
    end
    object chkPunish: TcxCheckBox
      Left = 149
      Top = 153
      Properties.OnChange = chkPunishPropertiesChange
      Style.StyleController = StyleController.CheckBoxStyle
      TabOrder = 5
      Width = 40
    end
    object btnInstCodeStr: TButton
      Left = 31
      Top = 206
      Width = 89
      Height = 25
      Caption = #25351#23450#27966#24037#39006#21029
      TabOrder = 7
      OnClick = btnInstCodeStrClick
    end
    inline Citem: TDataLookup
      Left = 153
      Top = 46
      Width = 309
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 1
      inherited CodeNo: TcxTextEdit
        Properties.OnChange = CitemCodeNoPropertiesChange
        OnExit = CitemCodeNoExit
      end
      inherited CodeName: TcxLookupComboBox
        Properties.OnChange = CitemCodeNamePropertiesChange
        Properties.OnInitPopup = CitemCodeNamePropertiesInitPopup
        Width = 235
      end
    end
    inline ServiceType: TDataLookup
      Left = 153
      Top = 20
      Width = 309
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 0
      inherited CodeNo: TcxTextEdit
        Properties.CharCase = ecUpperCase
        Properties.OnChange = ServiceTypeCodeNoPropertiesChange
      end
      inherited CodeName: TcxLookupComboBox
        Properties.OnChange = ServiceTypeCodeNamePropertiesChange
        Properties.OnInitPopup = ServiceTypeCodeNamePropertiesInitPopup
        Width = 234
      end
    end
    inline FaciItem: TDataLookup
      Left = 153
      Top = 72
      Width = 309
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 2
      inherited CodeNo: TcxTextEdit
        Properties.OnChange = FaciItemCodeNoPropertiesChange
      end
      inherited CodeName: TcxLookupComboBox
        Properties.OnInitPopup = FaciItemCodeNamePropertiesInitPopup
        Width = 235
      end
    end
    object cxButton1: TcxButton
      Left = 120
      Top = 206
      Width = 25
      Height = 25
      TabOrder = 8
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
    object edtAmount: TcxCurrencyEdit
      Left = 153
      Top = 99
      ParentFont = False
      Properties.DecimalPlaces = 0
      Properties.DisplayFormat = ',0;-,0'
      Properties.MaxValue = 99999999.000000000000000000
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 3
      Width = 110
    end
    object edtDiscountAmt: TcxCurrencyEdit
      Left = 153
      Top = 126
      ParentFont = False
      Properties.DecimalPlaces = 0
      Properties.DisplayFormat = ',0;-,0'
      Properties.MaxValue = 99999999.000000000000000000
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 4
      Width = 110
    end
    object cmbPenalType: TcxComboBox
      Left = 153
      Top = 177
      ParentFont = False
      Properties.DropDownListStyle = lsEditFixedList
      Properties.Items.Strings = (
        ''
        '1.'#22238#28335#29260#20729'('#20381#20840#38989')')
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 6
      Width = 264
    end
    object edtInstCodeStr: TcxTextEdit
      Left = 152
      Top = 207
      Enabled = False
      ParentFont = False
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 9
      Width = 311
    end
  end
  object ActionList1: TActionList
    Left = 176
    Top = 296
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
  object tmpCD078CDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 276
    Top = 113
  end
end
