object frmCD078B9: TfrmCD078B9
  Left = 254
  Top = 286
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #20778#24800#32068#21512#20195#30908#35079#35069#33267#36328#21312' [frmCD078B9]'
  ClientHeight = 305
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
    Top = 264
    Width = 593
    Height = 41
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
      Left = 210
      Top = 9
      Width = 82
      Height = 26
      Caption = #30906#23450
      TabOrder = 0
      OnClick = btnSaveClick
    end
    object btnCancel: TButton
      Left = 318
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
    Height = 264
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object btnBPCodeStr: TButton
      Left = 29
      Top = 19
      Width = 89
      Height = 25
      Caption = #20778#24800#32068#21512#20195#30908
      TabOrder = 0
      OnClick = btnBPCodeStrClick
    end
    object cxButton1: TcxButton
      Left = 120
      Top = 19
      Width = 25
      Height = 25
      TabOrder = 1
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
    object edtBPCodeStr: TcxTextEdit
      Left = 152
      Top = 20
      Enabled = False
      ParentFont = False
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 2
      Width = 409
    end
    object btnCompCodeStr: TButton
      Left = 29
      Top = 50
      Width = 89
      Height = 25
      Caption = #20844#21496#21029
      TabOrder = 3
      OnClick = btnCompCodeStrClick
    end
    object cxButton2: TcxButton
      Left = 120
      Top = 51
      Width = 25
      Height = 25
      TabOrder = 4
      OnClick = cxButton2Click
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
      Left = 152
      Top = 52
      Enabled = False
      ParentFont = False
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 5
      Width = 409
    end
    object txtMessage: TMemo
      Left = 23
      Top = 153
      Width = 532
      Height = 84
      ReadOnly = True
      ScrollBars = ssVertical
      TabOrder = 6
    end
    object grpCopy: TGroupBox
      Left = 30
      Top = 79
      Width = 253
      Height = 62
      Caption = #35079#35069#36984#38917
      Font.Charset = ANSI_CHARSET
      Font.Color = clInfoText
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 7
      object Label1: TLabel
        Left = 12
        Top = 30
        Width = 96
        Height = 14
        Caption = #36328#21312#20195#30908#24050#23384#22312#26178
        Font.Charset = ANSI_CHARSET
        Font.Color = clNavy
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
      end
      object radCopyNew: TcxRadioButton
        Left = 121
        Top = 18
        Width = 100
        Height = 17
        Caption = #19981#35206#33995
        Checked = True
        TabOrder = 0
        TabStop = True
      end
      object radCopyExist: TcxRadioButton
        Left = 11
        Top = 62
        Width = 250
        Height = 13
        Caption = #35079#35069#33267#36328#21312#65288#36328#21312#20195#30908#19981#23384#22312#26178#65292#19981#26032#22686#65289
        TabOrder = 1
        Visible = False
      end
      object radCopyOver: TcxRadioButton
        Left = 121
        Top = 38
        Width = 114
        Height = 17
        Caption = #21034#38500#24460#37325#26032#35206#33995
        TabOrder = 2
      end
    end
  end
  object DataConnection: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 428
    Top = 32
  end
  object DataWriter: TADOQuery
    Connection = DataConnection
    Parameters = <>
    Left = 505
    Top = 32
  end
  object DataReader: TADOQuery
    Connection = DBController.DataConnection
    Parameters = <>
    Left = 465
    Top = 32
  end
end
