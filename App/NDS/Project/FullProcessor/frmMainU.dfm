object frmMain: TfrmMain
  Left = 322
  Top = 171
  Width = 798
  Height = 522
  BorderIcons = []
  Caption = 'CA Command Processor'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poScreenCenter
  WindowState = wsMaximized
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object StatusBar1: TStatusBar
    Left = 0
    Top = 452
    Width = 790
    Height = 24
    Panels = <>
  end
  object pnlTitle: TPanel
    Left = 0
    Top = 0
    Width = 790
    Height = 41
    Align = alTop
    Caption = #21363#26178#30435#25511
    Color = clSkyBlue
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    object chbAction: TCheckBox
      Left = 16
      Top = 13
      Width = 98
      Height = 17
      Caption = #21855#21205#31243#24335
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      OnClick = chbActionClick
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 41
    Width = 33
    Height = 411
    Align = alLeft
    BevelOuter = bvLowered
    Color = clBackground
    Font.Charset = CHINESEBIG5_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = '@'#27161#26999#39636
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
    object Label1: TLabel
      Left = 5
      Top = 56
      Width = 22
      Height = 19
      Caption = 'CA'
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clFuchsia
      Font.Height = -19
      Font.Name = #32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label2: TLabel
      Left = 5
      Top = 96
      Width = 21
      Height = 19
      Caption = #25351
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clFuchsia
      Font.Height = -19
      Font.Name = #32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label3: TLabel
      Left = 5
      Top = 144
      Width = 21
      Height = 19
      Caption = #20196
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clFuchsia
      Font.Height = -19
      Font.Name = #32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label4: TLabel
      Left = 5
      Top = 192
      Width = 21
      Height = 19
      Caption = #20132
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clFuchsia
      Font.Height = -19
      Font.Name = #32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label5: TLabel
      Left = 5
      Top = 235
      Width = 21
      Height = 19
      Caption = #25563
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clFuchsia
      Font.Height = -19
      Font.Name = #32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label6: TLabel
      Left = 5
      Top = 285
      Width = 21
      Height = 19
      Caption = #27169
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clFuchsia
      Font.Height = -19
      Font.Name = #32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Label7: TLabel
      Left = 5
      Top = 328
      Width = 21
      Height = 19
      Caption = #32068
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clFuchsia
      Font.Height = -19
      Font.Name = #32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object Panel2: TPanel
    Left = 33
    Top = 41
    Width = 757
    Height = 411
    Align = alClient
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 3
    object Splitter1: TSplitter
      Left = 2
      Top = 330
      Width = 753
      Height = 4
      Cursor = crVSplit
      Align = alBottom
    end
    object PageControl1: TPageControl
      Left = 2
      Top = 2
      Width = 753
      Height = 328
      ActivePage = TabSheet1
      Align = alClient
      TabOrder = 0
      object TabSheet1: TTabSheet
        Caption = #20659#36865#32102' CA '#20043#25351#20196#34389#29702#29376#24907
        object livStatus: TListView
          Left = 0
          Top = 0
          Width = 745
          Height = 297
          Align = alClient
          Color = clInfoBk
          Columns = <
            item
              Caption = #20132#26131#24207#34399
              Width = 85
            end
            item
              Caption = #20844#21496#21029
              Width = 80
            end
            item
              Caption = #25351#20196#24207#34399
              Width = 85
            end
            item
              Caption = #25351#20196#20195#34399
              Width = 60
            end
            item
              Caption = #20659#36865
              Width = 40
            end
            item
              Caption = #22238#25033
              Width = 40
            end
            item
              Caption = #22238#25033#35338#24687#20195#30908
              Width = 90
            end
            item
              Caption = #22238#25033#35338#24687
              Width = 90
            end
            item
              Caption = #25805#20316#32773
              Width = 60
            end
            item
              Caption = #26178#38291
              Width = 130
            end>
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clNavy
          Font.Height = -13
          Font.Name = 'MS Sans Serif'
          Font.Style = []
          GridLines = True
          HideSelection = False
          HotTrack = True
          ParentFont = False
          SmallImages = imlStatus
          TabOrder = 0
          ViewStyle = vsReport
        end
      end
      object tabTesting: TTabSheet
        Caption = #28204#35430
        ImageIndex = 1
        object BitBtn2: TBitBtn
          Left = 208
          Top = 8
          Width = 105
          Height = 25
          Caption = 'get checksum'
          TabOrder = 0
          OnClick = BitBtn2Click
        end
        object Edit1: TEdit
          Left = 72
          Top = 24
          Width = 121
          Height = 24
          TabOrder = 1
          Text = '3000000002'
        end
        object StaticText1: TStaticText
          Left = 16
          Top = 24
          Width = 52
          Height = 20
          Caption = 'Card ID '
          TabOrder = 2
        end
        object BitBtn1: TBitBtn
          Left = 320
          Top = 8
          Width = 120
          Height = 25
          Caption = 'get signature(NDS)'
          TabOrder = 3
          OnClick = BitBtn1Click
        end
        object Button1: TButton
          Left = 440
          Top = 8
          Width = 129
          Height = 25
          Caption = #20659#36865#28204#35430#25351#20196
          TabOrder = 4
          OnClick = Button1Click
        end
        object Memo2: TMemo
          Left = 40
          Top = 104
          Width = 889
          Height = 63
          ScrollBars = ssVertical
          TabOrder = 5
        end
        object Memo1: TMemo
          Left = 40
          Top = 184
          Width = 297
          Height = 177
          ScrollBars = ssVertical
          TabOrder = 6
        end
        object Memo4: TMemo
          Left = 344
          Top = 184
          Width = 457
          Height = 175
          Lines.Strings = (
            'Memo4')
          ScrollBars = ssVertical
          TabOrder = 7
        end
        object Memo3: TMemo
          Left = 32
          Top = 376
          Width = 897
          Height = 129
          ScrollBars = ssBoth
          TabOrder = 8
        end
        object Edit2: TEdit
          Left = 72
          Top = 64
          Width = 121
          Height = 24
          TabOrder = 9
          Text = '1'
        end
        object StaticText2: TStaticText
          Left = 16
          Top = 64
          Width = 54
          Height = 20
          Caption = 'From ID '
          TabOrder = 10
        end
        object Button2: TButton
          Left = 208
          Top = 40
          Width = 105
          Height = 25
          Caption = 'get signatrue key'
          TabOrder = 11
          OnClick = Button2Click
        end
        object BitBtn3: TBitBtn
          Left = 320
          Top = 40
          Width = 120
          Height = 25
          Caption = 'get signature(MD5)'
          TabOrder = 12
          OnClick = BitBtn3Click
        end
        object BitBtn4: TBitBtn
          Left = 440
          Top = 40
          Width = 129
          Height = 25
          Caption = 'handle response'
          TabOrder = 13
          OnClick = BitBtn4Click
        end
        object Button3: TButton
          Left = 568
          Top = 8
          Width = 145
          Height = 25
          Caption = #39511#35657'response data'
          TabOrder = 14
          OnClick = Button3Click
        end
        object Button4: TButton
          Left = 568
          Top = 40
          Width = 75
          Height = 25
          Caption = 'Button4'
          TabOrder = 15
          OnClick = Button4Click
        end
        object Button5: TButton
          Left = 208
          Top = 72
          Width = 105
          Height = 25
          Caption = 'HandleXML'
          TabOrder = 16
          OnClick = Button5Click
        end
      end
    end
    object memErrorLog: TMemo
      Left = 2
      Top = 334
      Width = 753
      Height = 75
      Align = alBottom
      Color = clInfoBk
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      ScrollBars = ssVertical
      TabOrder = 1
    end
  end
  object MainMenu1: TMainMenu
    Left = 160
    Top = 8
    object N1: TMenuItem
      Caption = #21151#33021#34920'&1'
      object N2: TMenuItem
        Caption = #31995#32113#21443#25976#35373#23450' &1'
        OnClick = N2Click
      end
      object DB31: TMenuItem
        Caption = 'DB'#36899#32218#35373#23450' &2'
        OnClick = DB31Click
      end
      object N8: TMenuItem
        Caption = #36899#32218#36039#35338' &3'
        OnClick = N8Click
      end
      object N3: TMenuItem
        Caption = '-'
      end
      object N4: TMenuItem
        Caption = #32080#26463#31243#24335' &9'
        OnClick = N4Click
      end
    end
    object N5: TMenuItem
      Caption = #35498#26126'&2'
      object N6: TMenuItem
        Caption = '-'
      end
      object N7: TMenuItem
        Caption = 'About SMS Gateway'
        OnClick = N7Click
      end
    end
  end
  object imlStatus999: TImageList
    Left = 984
    Top = 8
  end
  object TcpClient1: TTcpClient
    Left = 128
    Top = 8
  end
  object Timer1: TTimer
    Enabled = False
    OnTimer = Timer1Timer
    Left = 96
    Top = 8
  end
  object imlStatus: TImageList
    Left = 192
    Top = 8
    Bitmap = {
      494C010103000400040010001000FFFFFFFFFF10FFFFFFFFFFFFFFFF424D3600
      0000000000003600000028000000400000001000000001002000000000000010
      0000000000000000000000000000000000004242420018181800000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000181818009CCEFF00424242001818
      1800000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000008484840084848400000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      CE00000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000CE00000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000018181800009CCE00319CFF009CCE
      FF00424242001818180000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000848484008484840084848400848484000000000000000000000000000000
      00000000000000000000000000000000000000000000000000000000CE000000
      FF000000CE000000000000000000000000000000000000000000000000000000
      000000000000000000000000FF00000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000018181800009CCE00009C
      CE00319CFF0063CEFF0042424200181818000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000008400000084000084848400848484000000000000000000000000000000
      00000000000000000000000000000000000000000000000000000000FF006363
      FF000000FF000000CE0000000000000000000000000000000000000000000000
      0000000000000000CE0000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000000000000000000018181800009C
      CE00009CCE00009CCE00319CFF0031CEFF004242420018181800000000000000
      0000000000000000000000000000000000000000000000000000000000000084
      0000008400000084000000840000848484008484840000000000000000000000
      00000000000000000000000000000000000000000000000000000000FF006363
      FF000000FF000000CE0000000000000000000000000000000000000000000000
      00000000CE000000FF0000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000001818
      1800009CCE00319CCE00319CCE00319CCE0031CEFF00319CCE00424242001818
      1800000000000000000000000000000000000000000000000000848484000084
      0000008400000084000000840000848484008484840084848400000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      FF006363FF000000FF000000CE00000000000000000000000000000000000000
      CE000000FF000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000018181800319CCE00319CFF00319CFF0031CEFF00319CFF0031CEFF0031CE
      FF00181818000000000000000000000000000000000000000000008400000084
      0000008400000084000000840000008400008484840084848400000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000FF006363FF000000FF000000CE0000000000000000000000CE000000
      FF00000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000181818004242420031CEFF0031CEFF0031CEFF0031CEFF00319CCE0031CE
      FF00181818000000000000000000000000000000000000840000008400000084
      0000008400000000000000840000008400000084000084848400848484000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000006363FF000000FF000000FF000000CE000000CE000000FF000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000001818
      1800319CCE0031CEFF0031CEFF0063CEFF0063CEFF00319CFF00424242001818
      1800000000000000000000000000000000000000000000000000008400000084
      0000008400000000000000840000008400000084000084848400848484008484
      8400000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000FF000000FF000000CE00000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000001818
      1800319CCE00319CFF0063CEFF0063CEFF0063CEFF0063CEFF0031CEFF004242
      4200000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000008400000084000000840000848484008484
      8400848484000000000000000000000000000000000000000000000000000000
      000000000000000000000000CE000000FF000000FF000000FF000000CE000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00001818180042424200319CCE0031CEFF0063CEFF0063CEFF0063CEFF009CCE
      FF00424242000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000084000000840000008400008484
      8400848484008484840000000000000000000000000000000000000000000000
      0000000000000000CE000000FF000000FF0000000000000000000000FF000000
      CE00000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000001818180042424200319CCE0031CEFF0063CEFF0063CE
      FF0063CEFF004242420000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000840000008400000084
      0000848484008484840084848400000000000000000000000000000000000000
      CE000000CE000000FF000000FF006363FF000000000000000000000000000000
      FF000000CE000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000001818180042424200319CCE0031CE
      FF0063CEFF009CCEFF0042424200000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000008400000084
      00000084000000000000848484008484840000000000000000000000CE000000
      FF000000FF000000FF000000FF00000000000000000000000000000000000000
      00000000FF000000CE0000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000181818004242
      4200009CCE0063CEFF009CCEFF00424242000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000084
      00000084000000840000000000000000000000000000000000000000FF006363
      FF006363FF000000FF0000000000000000000000000000000000000000000000
      000000000000000000000000CE00000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000181818004242420063CEFF00424242000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000084000000840000000000000000000000000000000000000000
      FF000000FF000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000018181800424242000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000424D3E000000000000003E000000
      2800000040000000100000000100010000000000800000000000000000000000
      000000000000000000000000FFFFFF003FFFFFFFFFFF00000FFFF9FFEFFD0000
      03FFF0FFC7FD000080FFF0FFC3FB0000C03FE07FC3F30000E00FC03FE1E70000
      F007C03FF0CF0000F007841FF81F0000E00FC40FFE3F0000E00FFE07FC1F0000
      F007FF03F8CF0000FC03FF81E0E70000FF01FFC4C1F30000FFC0FFE3C3FD0000
      FFF0FFF9E7FF0000FFFCFFFFFFFF000000000000000000000000000000000000
      000000000000}
  end
  object DataSource1: TDataSource
    DataSet = dtmMain.cdsXmlData
    Left = 152
    Top = 48
  end
  object tcpServer: TIdTCPServer
    Bindings = <
      item
        Port = 23
      end>
    CommandHandlers = <>
    DefaultPort = 0
    Greeting.NumericCode = 0
    MaxConnectionReply.NumericCode = 0
    OnConnect = tcpServerConnect
    OnExecute = tcpServerExecute
    ReplyExceptionCode = 0
    ReplyTexts = <>
    ReplyUnknownCommand.NumericCode = 0
    ThreadMgr = IdThreadMgrDefault1
    Left = 288
    Top = 8
  end
  object IdThreadMgrDefault1: TIdThreadMgrDefault
    Left = 320
    Top = 8
  end
end
