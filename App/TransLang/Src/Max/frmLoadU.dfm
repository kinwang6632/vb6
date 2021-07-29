object frmLoad: TfrmLoad
  Left = 286
  Top = 93
  Width = 377
  Height = 447
  BorderIcons = [biSystemMenu, biMinimize]
  Caption = 'Load project'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Bevel1: TBevel
    Left = 0
    Top = 378
    Width = 369
    Height = 1
    Align = alClient
  end
  object Panel1: TPanel
    Left = 0
    Top = 379
    Width = 369
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    object BitBtn1: TBitBtn
      Tag = 1
      Left = 124
      Top = 5
      Width = 80
      Height = 33
      Caption = 'Load All'
      Enabled = False
      TabOrder = 0
      OnClick = BitBtn1Click
    end
    object BitBtn2: TBitBtn
      Tag = 2
      Left = 207
      Top = 5
      Width = 80
      Height = 33
      Caption = 'Load Selection'
      Enabled = False
      TabOrder = 1
      OnClick = BitBtn1Click
    end
    object BitBtn3: TBitBtn
      Left = 289
      Top = 5
      Width = 80
      Height = 33
      Caption = 'Exit'
      TabOrder = 2
      OnClick = BitBtn1Click
    end
  end
  object FileListBox1: TFileListBox
    Left = 0
    Top = 145
    Width = 369
    Height = 96
    Align = alTop
    ImeName = '中文 (繁體) - 新注音'
    ItemHeight = 16
    Mask = '*.vbp'
    ShowGlyphs = True
    TabOrder = 1
    OnChange = FileListBox1Change
    OnDblClick = SpeedButton1Click
  end
  object DirectoryListBox1: TDirectoryListBox
    Left = 0
    Top = 49
    Width = 369
    Height = 96
    Align = alTop
    FileList = FileListBox1
    ImeName = '中文 (繁體) - 新注音'
    ItemHeight = 16
    TabOrder = 2
    OnClick = DirectoryListBox1Click
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 369
    Height = 49
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 3
    object Label3: TLabel
      Left = 88
      Top = 28
      Width = 136
      Height = 13
      Caption = 'C:\App\TransLang\Src\Max'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 12
      Top = 28
      Width = 75
      Height = 13
      Caption = 'Search Path:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label1: TLabel
      Left = 12
      Top = 8
      Width = 74
      Height = 13
      Caption = 'Search Disk:'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object DriveComboBox1: TDriveComboBox
      Left = 88
      Top = 5
      Width = 281
      Height = 19
      DirList = DirectoryListBox1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ImeName = '中文 (繁體) - 新注音'
      ParentFont = False
      TabOrder = 0
      TextCase = tcUpperCase
    end
  end
  object ListBox1: TListBox
    Left = 0
    Top = 282
    Width = 369
    Height = 96
    Align = alTop
    ImeName = '中文 (繁體) - 新注音'
    ItemHeight = 13
    MultiSelect = True
    TabOrder = 4
    OnClick = ListBox1Click
    OnDblClick = SpeedButton2Click
  end
  object Panel3: TPanel
    Left = 0
    Top = 241
    Width = 369
    Height = 41
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 5
    object SpeedButton2: TSpeedButton
      Left = 34
      Top = 9
      Width = 23
      Height = 22
      Hint = 'Del Selection from the listed below'
      Glyph.Data = {
        8E050000424D8E05000000000000360000002800000017000000130000000100
        18000000000058050000C40E0000C40E00000000000000000000CECECECECECE
        CECECECECECECECECECECECECECECECECECECECECECECECECECECECECECECECE
        CECECECECECECECECECECECECECECECECECECECECECECECECECECECECECECE00
        0000CECECECECECECECECECECECECECECECECECECECECECECECECECECECECECE
        CECECECECECECECECECECECECECECECECECECECECECECECECECECECECECECECE
        CECECECECECECE000000CECECECECECECECECECECECECECECECECECECECECECE
        CECECECECECECECECECECECECECECECECECECECECECECECECECECECECECECECE
        CECECECECECECECECECECECECECECE000000CECECECECECECECECECECECE0000
        0000000000000000000000000000000000000000000000000000000000000000
        0000CECECECECECECECECECECECECECECECECECECECECE000000CECECECECECE
        CECECECECECE0000000000000080800080800080800080800080800080800080
        80008080008080008080000000CECECECECECECECECECECECECECECECECECE00
        0000CECECECECECECECECECECECE00000000FFFF000000008080008080008080
        008080008080008080008080008080008080008080000000CECECECECECECECE
        CECECECECECECE000000CECECECECECECECECECECECE000000FFFFFF86868600
        0000008080008080008080008080008080008080008080008080008080008080
        000000CECECECECECECECECECECECE000000CECECECECECECECECECECECE0000
        0000FFFF86868600000000808000808000808000808000808000808000808000
        8080008080008080008080000000CECECECECECECECECE000000CECECECECECE
        CECECECECECE000000FFFFFF868686FFFFFF0000000000000000000000000080
        80008080008080008080008080008080008080000000CECECECECECECECECE00
        0000CECECECECECECECECECECECE00000000FFFF868686FFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFF000000000000000000000000000000000000000000000000CECE
        CECECECECECECE000000CECECECECECECECECECECECE000000FFFFFF868686FF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000CECECE
        CECECECECECECECECECECECECECECE000000CECECECECECECECECECECECE0000
        0000FFFF868686FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFF000000CECECECECECECECECECECECECECECECECECE000000CECECECECECE
        CECECECECECECECECE000000868686FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFFFFFFFFFF000000CECECECECECECECECECECECECECECECECECE00
        0000CECECECECECECECECECECECECECECECECECE868686FFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFF868686868686868686868686000000CECECECECECECECECECECE
        CECECECECECECE000000CECECECECECECECECECECECECECECECECECE868686FF
        FFFFFFFFFFFFFFFFFFFFFFFFFFFF868686000000000000000000000000000000
        000000CECECECECECECECECECECECE000000CECECECECECECECECECECECECECE
        CECECECE868686FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF86868680800080800080
        8000808000808000808000000000CECECECECECECECECE000000CECECECECECE
        CECECECECECECECECECECECE8686868686868686868686868686868686868686
        86FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00000000CECECECECECECECECE00
        0000CECECECECECECECECECECECECECECECECECECECECECECECECECECECECECE
        CECECECECECECECECE868686868686868686868686868686868686CECECECECE
        CECECECECECECE000000CECECECECECECECECECECECECECECECECECECECECECE
        CECECECECECECECECECECECECECECECECECECECECECECECECECECECECECECECE
        CECECECECECECECECECECECECECECE000000}
      Layout = blGlyphTop
      ParentShowHint = False
      ShowHint = True
      OnClick = SpeedButton2Click
    end
    object SpeedButton1: TSpeedButton
      Left = 3
      Top = 9
      Width = 23
      Height = 22
      Hint = 'Add selection from the above-listed'
      Glyph.Data = {
        D6050000424DD605000000000000360000002800000017000000140000000100
        180000000000A0050000C40E0000C40E00000000000000000000C8D0D4C8D0D4
        C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0
        D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D400
        0000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4
        C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0
        D4C8D0D4C8D0D4000000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8
        D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4
        C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000000C8D0D4C8D0D4C8D0D4C8D0D40000
        0000000000000000000000000000000000000000000000000000000000000000
        0000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000000C8D0D4C8D0D4
        C8D0D4C8D0D40000000000000084840084840084840084840084840084840084
        84008484008484008484000000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D400
        0000C8D0D4C8D0D4C8D0D4C8D0D400000000FFFF000000008484008484008484
        008484008484008484008484008484008484008484000000C8D0D4C8D0D4C8D0
        D4C8D0D4C8D0D4000000C8D0D4C8D0D4C8D0D4C8D0D4000000FFFFFF84848400
        0000008484008484008484008484008484008484008484008484008484008484
        000000C8D0D4C8D0D4C8D0D4C8D0D4000000C8D0D4C8D0D4C8D0D4C8D0D40000
        0000FFFF84848400000000848400848400848400848400848400848400848400
        8484008484008484008484000000C8D0D4C8D0D4C8D0D4000000C8D0D4C8D0D4
        C8D0D4C8D0D4000000FFFFFF848484FFFFFF0000000000000000000000000084
        84008484008484008484008484008484008484000000C8D0D4C8D0D4C8D0D400
        0000C8D0D4C8D0D4C8D0D4C8D0D400000000FFFF848484FFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFF000000000000000000000000000000000000000000000000C8D0
        D4C8D0D4C8D0D4000000C8D0D4C8D0D4C8D0D4C8D0D4000000FFFFFF848484FF
        FFFFC6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6C6FFFFFFC6C6C6000000C8D0D4
        C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000000C8D0D4C8D0D4C8D0D4C8D0D40000
        0000FFFF848484FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00
        0000000000000000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000000C8D0D4C8D0D4
        C8D0D4C8D0D4C8D0D4000000848484FFFFFFC6C6C6C6C6C6C6C6C6C6C6C6FFFF
        FFFFFFFF84848400FF0000FF00000000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D400
        0000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4848484FFFFFFFFFFFFFFFFFF
        FFFFFFFFFFFFFFFFFF84848484848400FF0000FF00000000000000C8D0D4C8D0
        D4C8D0D4C8D0D4000000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4848484FF
        FFFFC6C6C6C6C6C6C6C6C6C6C6C684848400FF0000FF0000FF0000FF0000FF00
        00FF00000000C8D0D4C8D0D4C8D0D4000000C8D0D4C8D0D4C8D0D4C8D0D4C8D0
        D4C8D0D4848484FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF84848400FF0000FF0000
        FF0000FF0000FF0000FF00000000C8D0D4C8D0D4C8D0D4000000C8D0D4C8D0D4
        C8D0D4C8D0D4C8D0D4C8D0D48484848484848484848484848484848484848484
        8484848484848400FF0000FF00000000000000C8D0D4C8D0D4C8D0D4C8D0D400
        0000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4
        C8D0D4C8D0D4C8D0D4C8D0D484848400FF0000FF00000000C8D0D4C8D0D4C8D0
        D4C8D0D4C8D0D4000000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8
        D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4848484848484C8D0D4
        C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000000C8D0D4C8D0D4C8D0D4C8D0D4C8D0
        D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8
        D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000000}
      Layout = blGlyphTop
      ParentShowHint = False
      ShowHint = True
      OnClick = SpeedButton1Click
    end
    object SpeedButton3: TSpeedButton
      Left = 67
      Top = 9
      Width = 23
      Height = 22
      Hint = 'Add all from the above-listed'
      Glyph.Data = {
        76060000424D7606000000000000360400002800000018000000180000000100
        0800000000004002000000000000000000000001000000010000FFFFFF00FFFF
        FF00C9C9C900C2C2C2008F8FC300BDBDBD00B2B2B200ACACAC00AFAFA7008C8C
        BC008989B7008484A4009A9A9A00968888008A8A8A0089848400878787008978
        78008972720000FFFF0000CDCD0000C7C70000B8B80000ABAB0000A3A300009F
        9F0000939300008D8D00008383007F7FB5006F6F96000000FF000000E1000000
        BC000000B500000096007D7D7D007B7B7600007F7F000074740037373F002D2B
        2B002B2727002524240020200800311E1E002B191900002222001D1D1C001515
        15001A1A000014140000001D1D00001414000C0C0C00000A0A00000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        00000000000000000000000000000000000000000000FFFFFF00010101010101
        0101010101010101010101010101010101010101010101010101010101010101
        0101010101010101010101010101010101010101010101010101010101010101
        0101010101010101010101010101010101010101010101010101010101010101
        0101010101010101010101010101010101010101010138383838383838383838
        383801010101010101010101010138381C1C1C1C1C1C1C1C1C1C380101010101
        0101010101013813381C1C1C1C1C1C1C1C1C1C38010101010101010101013800
        0F381C1C1C1C1C1C1C1C1C1C3801010101010101010138130F381C1C1C1C1C1C
        1C1C1C1C1C38010101010101010138000F00383838381C1C1C1C1C1C1C380101
        01010101010138130F0000000000383838383838383801010101010101013800
        0F00020202020202000238010101010101010101010138130F00000000000000
        003838380101010101010101010101380F000202020200000F13133801010101
        01010101010101010F0000000000000F0F131338380101010101010101010101
        0F00020202020F13131313131338010101010101010101010F00000000000F13
        131313131338010101010101010101010F0F0F0F0F0F0F0F0F13133838010101
        010101010101010101010101010101010F131338010101010101010101010101
        0101010101010101010F0F010101010101010101010101010101010101010101
        0101010101010101010101010101010101010101010101010101010101010101
        0101010101010101010101010101010101010101010101010101}
      Layout = blGlyphTop
      ParentShowHint = False
      ShowHint = True
      OnClick = SpeedButton3Click
    end
    object SpeedButton4: TSpeedButton
      Left = 98
      Top = 9
      Width = 23
      Height = 22
      Hint = 'Del all from listed below'
      Glyph.Data = {
        76060000424D7606000000000000360400002800000018000000180000000100
        0800000000004002000000000000000000000001000000010000FFFFFF00F2F2
        F200C9C9C900C2C2C2008F8FC300BDBDBD00B2B2B200ACACAC00AFAFA7008C8C
        BC008989B7008484A4009A9A9A00968888008A8A8A0089848400FFFFFF008978
        78008972720000FFFF0000CDCD0000C7C70000B8B80000ABAB0000A3A300009F
        9F0000939300008D8D00008383007F7FB5006F6F96000000FF000000E1000000
        BC000000B500000096007D7D7D007B7B7600007F7F000074740037373F002D2B
        2B002B2727002524240020200800311E1E002B191900002222001D1D1C001515
        15001A1A000014140000001D1D00001414000C0C0C00000A0A00000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        00000000000000000000000000000000000000000000FFFFFF00101010101010
        1010101010101010101010101010101010101010101010101010101010101010
        1010101010101010101010101010101010101010101010101010101010101010
        1010101010101010101010101010101010101010101010101010101010101010
        1010101010101010101010101010101010101010101010101010101010101010
        101010101010101010101010101038382F343434343434342F34101010101010
        1010101010103838161919191919191918263810101010101010101010103814
        34171A1A1A1A1A1A1A1B26371010101010101010101038100D38271A1A1A1A1A
        1A1A1A1737101010101010101010381312381C1819191A1A1A1A1A1A1C381010
        10101010101038100F10353434371719181918181C3810101010101010103815
        1110101010102E2D382E38383838101010101010101038100E10030101010100
        1002381010101010101010101010381411101010101010101003381010101010
        10101010101010380E100C030502101010063810101010101010101010101010
        0E10101010100607072438101010101010101010101010100F100C0305032533
        2C3238383810101010101010101010100E10101010100B222222222123381010
        10101010101010100607070707081D1F1F1F1F1F203810101010101010101010
        10101010101010040A0A0A091E10101010101010101010101010101010101010
        1010101010101010101010101010101010101010101010101010101010101010
        1010101010101010101010101010101010101010101010101010}
      Layout = blGlyphTop
      ParentShowHint = False
      ShowHint = True
      OnClick = SpeedButton4Click
    end
  end
end
