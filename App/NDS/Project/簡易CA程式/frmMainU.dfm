object frmMain: TfrmMain
  Left = 340
  Top = 0
  Width = 290
  Height = 547
  BorderIcons = [biMinimize]
  Caption = #31777#26131'CA'#31243#24335
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
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 282
    Height = 501
    Align = alClient
    BevelInner = bvSpace
    BevelOuter = bvLowered
    TabOrder = 0
    object BitBtn1: TBitBtn
      Left = 5
      Top = 473
      Width = 268
      Height = 24
      Caption = #32080#26463
      TabOrder = 13
      OnClick = BitBtn1Click
    end
    object BitBtn2: TBitBtn
      Left = 6
      Top = 414
      Width = 75
      Height = 24
      Caption = #38283#27231
      TabOrder = 7
      OnClick = BitBtn2Click
    end
    object StaticText1: TStaticText
      Left = 25
      Top = 154
      Width = 85
      Height = 20
      Caption = 'Subscriber ID'
      TabOrder = 14
    end
    object StaticText2: TStaticText
      Left = 45
      Top = 179
      Width = 64
      Height = 20
      Caption = #26234#24935#21345#32232#34399
      TabOrder = 15
    end
    object edtSubscriberID: TEdit
      Left = 115
      Top = 150
      Width = 121
      Height = 24
      Hint = #27231#19978#30418#32232#34399
      Color = clInfoBk
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnExit = edtSubscriberIDExit
      OnKeyPress = edtSubscriberIDKeyPress
    end
    object edtIccNo: TEdit
      Left = 115
      Top = 176
      Width = 121
      Height = 24
      Hint = 'Icc No'
      Color = clInfoBk
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
    end
    object StaticText3: TStaticText
      Left = 42
      Top = 259
      Width = 36
      Height = 24
      Caption = #38971#36947
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 16
    end
    object BitBtn3: TBitBtn
      Left = 86
      Top = 414
      Width = 75
      Height = 24
      Caption = #38364#27231
      TabOrder = 8
      OnClick = BitBtn3Click
    end
    object BitBtn4: TBitBtn
      Left = 6
      Top = 442
      Width = 75
      Height = 24
      Caption = #38283#38971#36947
      TabOrder = 10
      OnClick = BitBtn4Click
    end
    object BitBtn5: TBitBtn
      Left = 86
      Top = 442
      Width = 75
      Height = 24
      Caption = #38364#38971#36947
      TabOrder = 11
      OnClick = BitBtn5Click
    end
    object Panel2: TPanel
      Left = 2
      Top = 2
      Width = 278
      Height = 95
      Align = alTop
      BevelOuter = bvLowered
      TabOrder = 17
      object StaticText5: TStaticText
        Left = 43
        Top = 9
        Width = 64
        Height = 20
        Caption = #20351#29992#32773#21517#31281
        TabOrder = 3
      end
      object StaticText6: TStaticText
        Left = 42
        Top = 36
        Width = 64
        Height = 20
        Caption = #20351#29992#32773#23494#30908
        TabOrder = 4
      end
      object edtUserID: TEdit
        Left = 110
        Top = 6
        Width = 121
        Height = 24
        Color = clInfoBk
        TabOrder = 0
      end
      object edtPasswd: TEdit
        Left = 110
        Top = 34
        Width = 121
        Height = 24
        Color = clInfoBk
        PasswordChar = '*'
        TabOrder = 1
      end
      object btnLogin: TBitBtn
        Left = 42
        Top = 63
        Width = 193
        Height = 25
        Caption = #30331#20837'/'#36899#25509#33267'CA'
        TabOrder = 2
        OnClick = btnLoginClick
      end
    end
    object memCh: TMemo
      Left = 43
      Top = 283
      Width = 193
      Height = 88
      Color = clInfoBk
      ReadOnly = True
      ScrollBars = ssVertical
      TabOrder = 6
    end
    object btnChooseCh: TBitBtn
      Left = 115
      Top = 258
      Width = 121
      Height = 21
      Caption = #40670#36984#38971#36947
      TabOrder = 5
      OnClick = btnChooseChClick
    end
    object BitBtn7: TBitBtn
      Left = 165
      Top = 414
      Width = 108
      Height = 24
      Caption = #35373#23450#35242#23376#23494#30908
      TabOrder = 9
      OnClick = BitBtn7Click
    end
    object chbShowResult: TCheckBox
      Left = 24
      Top = 130
      Width = 169
      Height = 17
      Caption = #39023#31034#38283#27231#32080#26524#35222#31383
      TabOrder = 0
    end
    object Panel3: TPanel
      Left = 2
      Top = 97
      Width = 278
      Height = 32
      Align = alTop
      BevelInner = bvLowered
      BevelOuter = bvLowered
      Color = clSkyBlue
      Font.Charset = ANSI_CHARSET
      Font.Color = clNavy
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 18
    end
    object StaticText7: TStaticText
      Left = 36
      Top = 203
      Width = 71
      Height = 20
      Hint = '('#26368#22810'8'#20491'bytes. '#25976#23383','#23383#20803#22343#21487')'
      Caption = 'RegionKey'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 19
    end
    object StaticText8: TStaticText
      Left = 37
      Top = 232
      Width = 70
      Height = 20
      Caption = 'Bouquet ID'
      TabOrder = 20
    end
    object edtRegionKey: TEdit
      Left = 115
      Top = 202
      Width = 121
      Height = 24
      Hint = '('#26368#22810'8'#20491'bytes. '#25976#23383','#23383#20803#22343#21487')'
      Color = clInfoBk
      MaxLength = 8
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
    end
    object edtBouquetID: TEdit
      Left = 115
      Top = 229
      Width = 121
      Height = 24
      Color = clInfoBk
      MaxLength = 4
      TabOrder = 4
      OnExit = edtBouquetIDExit
      OnKeyPress = edtBouquetIDKeyPress
    end
    object BitBtn8: TBitBtn
      Left = 166
      Top = 441
      Width = 107
      Height = 25
      Caption = #35373#23450' Bouquet ID'
      TabOrder = 12
      OnClick = BitBtn8Click
    end
    object StaticText4: TStaticText
      Left = 10
      Top = 380
      Width = 129
      Height = 20
      Caption = 'Master Subscriber ID'
      TabOrder = 21
    end
    object edtMSubscriberID: TEdit
      Left = 142
      Top = 378
      Width = 121
      Height = 24
      Hint = #27231#19978#30418#32232#34399
      Color = clInfoBk
      ParentShowHint = False
      ShowHint = True
      TabOrder = 22
      OnExit = edtSubscriberIDExit
      OnKeyPress = edtSubscriberIDKeyPress
    end
  end
  object MainMenu1: TMainMenu
    Left = 8
    Top = 64
    object N1: TMenuItem
      Caption = #20854#20182#21151#33021'(&O)'
      object N2: TMenuItem
        Caption = #20316#26989#27511#31243#34920'(&1)'
        OnClick = N2Click
      end
      object N4: TMenuItem
        Caption = #36039#26009#21516#27493'(&2)'
        OnClick = N4Click
      end
      object N3: TMenuItem
        Caption = #21443#25976#35373#23450'(&3)'
        OnClick = N3Click
      end
    end
  end
  object IdTCPClient1: TIdTCPClient
    RecvBufferSize = 8192
    Port = 0
    Left = 192
    Top = 120
  end
  object Timer1: TTimer
    Interval = 500
    OnTimer = Timer1Timer
    Left = 160
    Top = 120
  end
end
