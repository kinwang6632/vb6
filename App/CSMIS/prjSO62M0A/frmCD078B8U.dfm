object frmCD078B8: TfrmCD078B8
  Left = 254
  Top = 286
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #20778#24800#32068#21512#20195#30908#35079#35069#33267#26412#21312' [frmCD078B8]'
  ClientHeight = 203
  ClientWidth = 593
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
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 156
    Width = 593
    Height = 47
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 1
    object Bevel1: TBevel
      Left = 0
      Top = 0
      Width = 593
      Height = 3
      Align = alTop
      Shape = bsTopLine
    end
    object btnSave: TButton
      Left = 204
      Top = 9
      Width = 82
      Height = 26
      Caption = #30906#23450
      TabOrder = 0
      OnClick = btnSaveClick
    end
    object btnCancel: TButton
      Left = 306
      Top = 9
      Width = 82
      Height = 26
      Cancel = True
      Caption = #21462#28040
      ModalResult = 2
      TabOrder = 1
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 593
    Height = 156
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object Label1: TLabel
      Left = 44
      Top = 24
      Width = 72
      Height = 14
      Caption = #20778#24800#32068#21512#20195#30908
    end
    object Label2: TLabel
      Left = 44
      Top = 122
      Width = 420
      Height = 14
      Caption = #27880#24847#65306#12308#24050#23384#22312#20195#30908#12309#20043#25152#26377#32048#38917#36039#26009#23559#26371#34987#12308#20778#24800#32068#21512#20195#30908#12309#20043#32048#38917#36039#26009#21462#20195
      Visible = False
    end
    inline BPCode: TDataLookup
      Left = 150
      Top = 20
      Width = 420
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 0
      inherited CodeNo: TcxTextEdit
        Properties.OnChange = BPCodeCodeNoPropertiesChange
        OnExit = BPCodeCodeNoExit
      end
      inherited CodeName: TcxLookupComboBox
        Left = 61
        Properties.OnChange = BPCodeCodeNamePropertiesChange
        Properties.OnInitPopup = BPCodeCodeNamePropertiesInitPopup
        Width = 356
      end
    end
    object edtNewCodeNo: TcxTextEdit
      Left = 150
      Top = 58
      ParentFont = False
      Properties.Alignment.Horz = taLeftJustify
      Properties.MaxLength = 10
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 2
      Width = 60
    end
    object radCopyNew: TcxRadioButton
      Left = 28
      Top = 60
      Width = 120
      Height = 17
      Caption = #35079#35069#33267#26032#20195#30908
      Checked = True
      TabOrder = 1
      TabStop = True
      OnClick = radCopyNewClick
      LookAndFeel.Kind = lfFlat
    end
    object radCopyExist: TcxRadioButton
      Left = 28
      Top = 89
      Width = 120
      Height = 17
      Caption = #35079#35069#33267#24050#23384#22312#20195#30908
      TabOrder = 4
      Visible = False
      OnClick = radCopyExistClick
      LookAndFeel.Kind = lfFlat
    end
    inline CopyToBPCode: TDataLookup
      Left = 150
      Top = 86
      Width = 420
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      Enabled = False
      TabOrder = 5
      Visible = False
      inherited CodeNo: TcxTextEdit
        Properties.OnChange = CopyToBPCodeCodeNoPropertiesChange
        OnExit = CopyToBPCodeCodeNoExit
      end
      inherited CodeName: TcxLookupComboBox
        Properties.OnChange = CopyToBPCodeCodeNamePropertiesChange
        Properties.OnInitPopup = CopyToBPCodeCodeNamePropertiesInitPopup
        Width = 355
      end
    end
    object edtNewCodeName: TcxTextEdit
      Left = 212
      Top = 58
      ParentFont = False
      Properties.Alignment.Horz = taLeftJustify
      Properties.MaxLength = 100
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 3
      Width = 355
    end
  end
end
