object fmLogin: TfmLogin
  Left = 483
  Top = 389
  ActiveControl = txtUserId
  BorderStyle = bsDialog
  Caption = #30331#20837
  ClientHeight = 181
  ClientWidth = 402
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object Panel1: TPanel
    Left = 0
    Top = 136
    Width = 402
    Height = 45
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    object Bevel1: TBevel
      Left = 0
      Top = 0
      Width = 402
      Height = 3
      Align = alTop
      Shape = bsTopLine
    end
    object btnOk: TcxButton
      Left = 185
      Top = 8
      Width = 90
      Height = 28
      Caption = #30331#20837
      Default = True
      TabOrder = 0
      OnClick = btnOkClick
      Glyph.Data = {
        36040000424D3604000000000000360000002800000010000000100000000100
        2000000000000004000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        00000000000000000000008B000E008200B5038805E8008A005A000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        00000000000000A3000D008E00C513AC27FF15B32BFF06940EFB009300590000
        0000000000000000000000000000000000000000000000000000000000000000
        000000B1000C009400C216AE2EFF17B231FF14AD2AFF13B129FF06920BFB0095
        00560000000000000000000000000000000000000000000000000000000000A9
        000A009A00C019B233FF1CB63AFF18B334FF34BC4DFF17B030FF13B02AFF0691
        0BFA009500520000000000000000000000000000000000000000000000000CA3
        16B21FB73EFF20BA42FF1FB840FF11A824FF018E01E962C771FF19B232FF13B1
        2BFF06910BF90095004E000000000000000000000000000000000000000028B3
        38E942C865FF1FBB46FF15B32CFF00A400990092000F009300C761C871FF19B1
        31FF13B22CFF06900BF80095004A000000000000000000000000000000000EB9
        246543C456FC36BC4EFF00AF089C000000010000000000AA0010009200CC63CA
        72FF17B131FF14B12CFF069109F7009B00480000000000000000000000000000
        000010D2242E0DBE204E0097000200000000000000000000000000A800120093
        01D063CA72FF17B131FF14B32DFF069009F6008F004400000000000000000000
        00000000000000000000000000000000000000000000000000000000000000AD
        0015009400D464CB75FF15B131FF16B330FF028904E800000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        000000AB0017019703D866CD78FF2ABC46FF008A02D300000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        000000000000009A001B009401CB049306D10084002F00000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000}
      LookAndFeel.Kind = lfFlat
    end
    object btnCancel: TcxButton
      Left = 287
      Top = 8
      Width = 90
      Height = 28
      Cancel = True
      Caption = #21462#28040
      ModalResult = 2
      TabOrder = 1
      Glyph.Data = {
        36040000424D3604000000000000360000002800000010000000100000000100
        2000000000000004000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        A4140206B0CB030AAEE10000A67B000098070000000000000000000000000000
        00000000A78500009A8B0000A72A000000000000000000000000000000000000
        A98D1844F6FF194DF8FF1031D2FF0102ABD00000B62400000000000000010000
        B1930928D7FF092ED7FF0313B3FF0000AC8F0000000000000000000000000103
        B3D02451F9FF1F52FFFF1D4FFFFF1744E8FF040BB0EC0000B03A0000AC970D2E
        DDFF1142F9FF0D3DF5FF0B3BF0FF041ABCFF0000A55D00000000000000000000
        AE421832DBFF285BFFFF2456FFFF2253FFFF1B4BF1FF050DB1FB0F30DDFF164A
        FEFF1344F9FF1041F6FF0E3EF6FF0A3CF0FF00009FBF00000000000000000000
        BE511F37DDFF3A6FFFFF2C5EFFFF295AFFFF2657FFFF2052FCFF1C4FFFFF194A
        FDFF1646FAFF1445FAFF0F3DF2FF020AB1E70000A84300000000000000000000
        00010000C866121DC8F03D6AFBFF3567FFFF2C5DFFFF2859FFFF2253FFFF1D4E
        FFFF1A4DFFFF123DEDFF0002ACCB0000BA1A0000000000000000000000000000
        0000000000000000CC1C0000B6BA2E4EE7FF3668FFFF2E5EFFFF2859FFFF2254
        FFFF163DEAFF0000A8BE0000AB0B000000000000000000000000000000000000
        000000000000000000000000BF4D253FDFFF3B6DFFFF3464FFFF2E5EFFFF2759
        FFFF1B46EAFF0001ACCF0000A908000000000000000000000000000000000000
        000000000000000000000203C8C14B7CFFFF4170FFFF3B6BFFFF396CFFFF2D5E
        FFFF2558FFFF1336D7FF0000B693000000000000000000000000000000000000
        0000000000000000D92C263CDBFE5080FFFF4575FFFF3662FAFF0D14C3FE3C6D
        FFFF2A5BFFFF2053FDFF0B1DC4FD0000C0470000000000000000000000000000
        0000000000000000CB8B527CFAFF5081FFFF4B7DFFFF0B13C9EA0000BB420E15
        C7E9386AFFFF2456FFFF1A4AF2FF0207B3E30000B50F00000000000000000000
        000000000000131CDDE06A9CFFFF5788FFFF2B46E7FF0000CD5C000000000000
        CD320F1BCBF23065FFFF1F51FFFF1439DDFF0000B19C00000000000000000000
        0000000000000000DE913A52E4FE5782FBFF0101D0C200000000000000000000
        00000000CC471426D6FA265AFFFF0F2EE3FF0103B8CF00000000000000000000
        000000000000000000000000CF3C0000C0A40000CE3000000000000000000000
        0000000000000000C4630001B8BC0000B5490000770200000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000}
      LookAndFeel.Kind = lfFlat
    end
  end
  object ScrollBox1: TScrollBox
    Left = 0
    Top = 0
    Width = 402
    Height = 136
    Align = alClient
    BevelInner = bvNone
    BorderStyle = bsNone
    TabOrder = 1
    object Label1: TLabel
      Left = 47
      Top = 25
      Width = 50
      Height = 16
      Caption = #31995#32113#21488':'
    end
    object Label2: TLabel
      Left = 32
      Top = 55
      Width = 65
      Height = 16
      Caption = #30331#20837#24115#34399':'
    end
    object Label3: TLabel
      Left = 32
      Top = 86
      Width = 65
      Height = 16
      Caption = #30331#20837#23494#30908':'
    end
    object cmbSo: TcxComboBox
      Left = 103
      Top = 22
      ParentFont = False
      Properties.DropDownListStyle = lsFixedList
      Style.StyleController = StyleModule.cxEditStyle
      TabOrder = 0
      Width = 274
    end
    object txtUserId: TcxTextEdit
      Left = 103
      Top = 52
      ParentFont = False
      Style.StyleController = StyleModule.cxEditStyle
      TabOrder = 1
      Width = 181
    end
    object txtPassword: TcxTextEdit
      Left = 103
      Top = 84
      ParentFont = False
      Properties.EchoMode = eemPassword
      Style.StyleController = StyleModule.cxEditStyle
      TabOrder = 2
      Width = 181
    end
  end
end
