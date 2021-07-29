object frmMain: TfrmMain
  Left = 339
  Top = 41
  Width = 290
  Height = 570
  BorderIcons = [biMinimize]
  Caption = #31777#26131#38283#38364#27231#31243#24335
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
    Height = 524
    Align = alClient
    BevelInner = bvSpace
    BevelOuter = bvLowered
    TabOrder = 0
    object BitBtn1: TBitBtn
      Left = 7
      Top = 451
      Width = 266
      Height = 25
      Caption = #32080#26463
      TabOrder = 8
      OnClick = BitBtn1Click
    end
    object BitBtn2: TBitBtn
      Left = 6
      Top = 371
      Width = 75
      Height = 25
      Caption = #38283#27231
      TabOrder = 3
      OnClick = BitBtn2Click
    end
    object StaticText1: TStaticText
      Left = 25
      Top = 187
      Width = 85
      Height = 20
      Caption = 'Subscriber ID'
      TabOrder = 9
    end
    object StaticText2: TStaticText
      Left = 45
      Top = 220
      Width = 64
      Height = 20
      Caption = #26234#24935#21345#32232#34399
      TabOrder = 10
    end
    object edtSubscriberID: TEdit
      Left = 115
      Top = 183
      Width = 121
      Height = 24
      Hint = #27231#19978#30418#32232#34399
      Color = clInfoBk
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
    end
    object edtIccNo: TEdit
      Left = 115
      Top = 217
      Width = 121
      Height = 24
      Hint = 'Icc No'
      Color = clInfoBk
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
    end
    object StaticText3: TStaticText
      Left = 42
      Top = 250
      Width = 36
      Height = 24
      Caption = #38971#36947
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 11
    end
    object BitBtn3: TBitBtn
      Left = 86
      Top = 371
      Width = 75
      Height = 25
      Caption = #38364#27231
      TabOrder = 4
      OnClick = BitBtn3Click
    end
    object BitBtn4: TBitBtn
      Left = 6
      Top = 411
      Width = 75
      Height = 25
      Caption = #38283#38971#36947
      TabOrder = 6
      OnClick = BitBtn4Click
    end
    object BitBtn5: TBitBtn
      Left = 86
      Top = 411
      Width = 75
      Height = 25
      Caption = #38364#38971#36947
      TabOrder = 7
      OnClick = BitBtn5Click
    end
    object Panel2: TPanel
      Left = 2
      Top = 2
      Width = 278
      Height = 143
      Align = alTop
      BevelOuter = bvLowered
      TabOrder = 12
      object StaticText4: TStaticText
        Left = 42
        Top = 19
        Width = 64
        Height = 20
        Caption = #36039#26009#24235#21029#21517
        TabOrder = 4
      end
      object StaticText5: TStaticText
        Left = 43
        Top = 51
        Width = 64
        Height = 20
        Caption = #20351#29992#32773#21517#31281
        TabOrder = 5
      end
      object StaticText6: TStaticText
        Left = 42
        Top = 86
        Width = 64
        Height = 20
        Caption = #20351#29992#32773#23494#30908
        TabOrder = 6
      end
      object edtAlias: TEdit
        Left = 110
        Top = 16
        Width = 121
        Height = 24
        Color = clInfoBk
        TabOrder = 0
      end
      object edtUserID: TEdit
        Left = 110
        Top = 48
        Width = 121
        Height = 24
        Color = clInfoBk
        TabOrder = 1
      end
      object edtPasswd: TEdit
        Left = 110
        Top = 84
        Width = 121
        Height = 24
        Color = clInfoBk
        PasswordChar = '*'
        TabOrder = 2
      end
      object BitBtn6: TBitBtn
        Left = 42
        Top = 113
        Width = 193
        Height = 25
        Caption = #30331#20837
        TabOrder = 3
        OnClick = BitBtn6Click
      end
    end
    object memCh: TMemo
      Left = 43
      Top = 274
      Width = 193
      Height = 89
      Color = clInfoBk
      ReadOnly = True
      ScrollBars = ssVertical
      TabOrder = 13
    end
    object btnChooseCh: TBitBtn
      Left = 115
      Top = 249
      Width = 121
      Height = 22
      Caption = #40670#36984#38971#36947
      TabOrder = 2
      OnClick = btnChooseChClick
    end
    object BitBtn7: TBitBtn
      Left = 165
      Top = 371
      Width = 108
      Height = 25
      Caption = #35373#23450#35242#23376#23494#30908
      TabOrder = 5
      OnClick = BitBtn7Click
    end
    object chbShowResult: TCheckBox
      Left = 24
      Top = 160
      Width = 169
      Height = 17
      Caption = #39023#31034#38283#27231#32080#26524#35222#31383
      TabOrder = 14
    end
    object Panel3: TPanel
      Left = 2
      Top = 481
      Width = 278
      Height = 41
      Align = alBottom
      BevelOuter = bvNone
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 15
    end
  end
  object ADOConnection1: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 8
    Top = 8
  end
  object ADOCommand1: TADOCommand
    Connection = ADOConnection1
    Parameters = <>
    Left = 40
    Top = 8
  end
  object qryChInfo: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 72
    Top = 8
  end
  object qryActionTimes: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 112
    Top = 8
  end
  object MainMenu1: TMainMenu
    Left = 24
    Top = 40
    object N1: TMenuItem
      Caption = #20854#20182#21151#33021'(&O)'
      object N2: TMenuItem
        Caption = #20316#26989#27511#31243#34920'(&1)'
        OnClick = N2Click
      end
      object N3: TMenuItem
        Caption = #21443#25976#35373#23450'(&2)'
        OnClick = N3Click
      end
    end
  end
  object qryReport1: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from NDS002')
    Left = 152
    Top = 8
    object qryReport1ICC_NO: TStringField
      FieldName = 'ICC_NO'
      Size = 12
    end
    object qryReport1SUBSCRIBER_ID: TStringField
      FieldName = 'SUBSCRIBER_ID'
      Size = 16
    end
    object qryReport1PRODUCTID: TStringField
      FieldName = 'PRODUCTID'
      Size = 1024
    end
    object qryReport1ACTION: TStringField
      FieldName = 'ACTION'
    end
    object qryReport1AUTHORSTARTDATE: TDateTimeField
      FieldName = 'AUTHORSTARTDATE'
    end
    object qryReport1AUTHORSTOPDATE: TDateTimeField
      FieldName = 'AUTHORSTOPDATE'
    end
    object qryReport1PINCODE: TStringField
      FieldName = 'PINCODE'
      Size = 4
    end
    object qryReport1OPERATOR: TStringField
      FieldName = 'OPERATOR'
    end
    object qryReport1UPDTIME: TDateTimeField
      FieldName = 'UPDTIME'
    end
  end
  object qryResult: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 248
    Top = 8
  end
  object qryGetSeqNo: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 248
    Top = 56
  end
end
