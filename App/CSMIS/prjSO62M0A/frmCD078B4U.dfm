object frmCD078B4: TfrmCD078B4
  Left = 712
  Top = 211
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'frmCD078B4'
  ClientHeight = 274
  ClientWidth = 433
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
    Top = 233
    Width = 433
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 1
    object Bevel1: TBevel
      Left = 0
      Top = 0
      Width = 433
      Height = 3
      Align = alTop
      Shape = bsTopLine
    end
    object edtDML: TcxTextEdit
      Left = 363
      Top = 11
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
      TabOrder = 1
    end
    object btnCancel: TButton
      Left = 87
      Top = 10
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
    Width = 433
    Height = 233
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object Label1: TLabel
      Left = 28
      Top = 56
      Width = 48
      Height = 14
      Caption = #35037#27231#39006#21029
      Font.Charset = ANSI_CHARSET
      Font.Color = clRed
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 28
      Top = 29
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
    object Label3: TLabel
      Left = 28
      Top = 84
      Width = 48
      Height = 14
      Caption = #31649#27966#39006#21029
      Font.Charset = ANSI_CHARSET
      Font.Color = clRed
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label4: TLabel
      Left = 32
      Top = 143
      Width = 44
      Height = 14
      Caption = #21443' '#32771' '#34399
    end
    object Label5: TLabel
      Left = 28
      Top = 114
      Width = 48
      Height = 14
      Caption = #27966#24037#40670#25976
    end
    object Label6: TLabel
      Left = 28
      Top = 172
      Width = 48
      Height = 14
      Caption = #25286#27231#39006#21029
    end
    object Label7: TLabel
      Left = 172
      Top = 143
      Width = 60
      Height = 14
      Caption = #21487#26356#25563#38971#36947
    end
    object chkInterdepend: TcxCheckBox
      Left = 28
      Top = 202
      Caption = #26381#21209#20381#23384
      Style.StyleController = StyleController.CheckBoxStyle
      TabOrder = 6
      Width = 74
    end
    inline InstCode: TDataLookup
      Left = 84
      Top = 52
      Width = 309
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 1
      inherited CodeNo: TcxTextEdit
        Properties.OnChange = InstCodeCodeNoPropertiesChange
        OnExit = InstCodeCodeNoExit
      end
      inherited CodeName: TcxLookupComboBox
        Properties.OnChange = InstCodeCodeNamePropertiesChange
        Properties.OnInitPopup = InstCodeCodeNamePropertiesInitPopup
        Width = 235
      end
    end
    inline ServiceType: TDataLookup
      Left = 84
      Top = 25
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
        Width = 235
      end
    end
    inline GroupNo: TDataLookup
      Left = 84
      Top = 80
      Width = 309
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 2
      inherited CodeNo: TcxTextEdit
        Properties.OnChange = GroupNoCodeNoPropertiesChange
      end
      inherited CodeName: TcxLookupComboBox
        Properties.OnInitPopup = GroupNoCodeNamePropertiesInitPopup
        Width = 235
      end
    end
    object edtWorkUnit: TcxCurrencyEdit
      Left = 84
      Top = 110
      EditValue = 0.000000000000000000
      ParentFont = False
      Properties.Alignment.Horz = taRightJustify
      Properties.DisplayFormat = ',0.00;-,0.00'
      Properties.MaxValue = 999.990000000000000000
      Properties.NullString = '0'
      Properties.OnValidate = edtWorkUnitPropertiesValidate
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 3
      Width = 68
    end
    object edtRefNo: TcxTextEdit
      Left = 84
      Top = 138
      ParentFont = False
      Properties.MaxLength = 3
      Properties.OnChange = edtRefNoPropertiesChange
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 4
      Width = 53
    end
    inline PRCode: TDataLookup
      Left = 84
      Top = 167
      Width = 309
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 5
      inherited CodeNo: TcxTextEdit
        Properties.OnChange = PRCodeCodeNoPropertiesChange
      end
      inherited CodeName: TcxLookupComboBox
        Properties.OnInitPopup = PRCodeCodeNamePropertiesInitPopup
        Width = 234
      end
    end
    object btnInstCode2: TBitBtn
      Left = 238
      Top = 137
      Width = 26
      Height = 24
      Hint = #39023#31034#22810#38542#20778#24800#21443#25976#26126#32048
      ParentShowHint = False
      ShowHint = True
      TabOrder = 7
      OnClick = btnInstCode2Click
      Glyph.Data = {
        36040000424D3604000000000000360000002800000010000000100000000100
        2000000000000004000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000A27A66CCC9967EFFC5927CFFC18E7AFFBD8A78FFB98676FFB58274FFB17E
        72FFAD7A70FFA9766EFFA5726CFF845B56CC000000000000000000000000A57C
        66CCEEDACEFFF9F9F9FFF9F9F9FFFAFAFAFFFCFCFCFFFEFEFEFFFFFFFFFFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFDBC7C4FF845B56CC0000000000000000D29F
        82FFFFFFFFFF0033FFFF0033FFFF0033FFFF0033FFFF0033FFFFFDFDFDFFF38D
        00FFF48E00FFF69000FFF89200FFFFFFFFFFA5726CFF0000000000000000D7A4
        84FFFFFFFFFFF3F3F3FFF4F4F4FFF5F5F5FFF6F6F6FFF8F8F8FFF9F9F9FFF18B
        00FFF38D00FFF48E00FFF69000FFFFFFFFFFA9766EFF0000000000000000DCA9
        87FFFFFFFFFF00BFF2FF00BFF2FF00BFF2FF00BFF2FF00BFF2FFF7F7F7FFEF89
        00FFF18B00FFF38D00FFF48E00FFFFFFFFFFAE7B70FF0000000000000000E1AE
        8AFFFFFFFFFFECECECFFEDEDEDFFEEEEEEFFF1F1F1FFF3F3F3FFF5F5F5FFF7F7
        F7FFF9F9F9FFFCFCFCFFFFFFFFFFFFFFFFFFB38073FF0000000000000000E5B2
        8CFFFFFFFFFFE37D00FFE47E00FFE57F00FFE78100FFE88200FFF2F2F2FF0033
        FFFF0033FFFF0033FFFF0033FFFFFFFFFFFFB78475FF0000000000000000E9B6
        8EFFFFFFFFFFE17B00FFE37D00FFE37D00FFE57F00FFE68000FFEFEFEFFFF3F3
        F3FFF5F5F5FFF8F8F8FFFAFAFAFFFEFEFEFFBA8777FF0000000000000000EDBA
        90FFFFFFFFFFDF7900FFE07A00FFE27C00FFE37D00FFE57F00FFEEEEEEFF00A9
        DCFF00A9DCFF00A9DCFF00A9DCFFFCFCFCFFBE8B79FF0000000000000000F1BE
        92FFFFFFFFFFE0E0E0FFE3E3E3FFE5E5E5FFE7E7E7FFE9E9E9FFECECECFFEEEE
        EEFFF2F2F2FFF5F5F5FFF8F8F8FFFAFAFAFFC28F7BFF0000000000000000F5C2
        94FFFFFFFFFF0EA71CFF0FA81EFF12AB24FF14AD28FF17B02EFF18B131FF1AB3
        36FF1DB63BFF1FB83FFF21BA43FFF9F9F9FFC6937DFF0000000000000000F9C6
        96FFFFFFFFFF0BA417FF0EA71CFF10A920FF13AC26FF15AE2AFF17B02FFF19B2
        34FF1CB539FF1EB73EFF21BA43FFF9F9F9FFC9967EFF0000000000000000C29A
        76CCFDE8D5FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFEEDACEFF885F58CC00000000000000000000
        0000BE9673CCF9C696FFF5C294FFF1BE92FFEDBA90FFE9B68EFFE5B28CFFE1AE
        8AFFDDAA88FFD9A686FFD5A284FF885F58CC0000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000}
    end
  end
  object ActionList1: TActionList
    Left = 292
    Top = 132
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
end
