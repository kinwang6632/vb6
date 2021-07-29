object frmMain: TfrmMain
  Left = 748
  Top = 200
  Width = 290
  Height = 506
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
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 282
    Height = 460
    Align = alClient
    BevelInner = bvSpace
    BevelOuter = bvLowered
    TabOrder = 0
    object btnExit: TBitBtn
      Left = 7
      Top = 427
      Width = 266
      Height = 25
      Caption = #32080#26463
      TabOrder = 8
      OnClick = btnExitClick
    end
    object btnPair: TBitBtn
      Left = 6
      Top = 347
      Width = 75
      Height = 25
      Caption = #38283#27231
      TabOrder = 3
      OnClick = btnPairClick
    end
    object StaticText1: TStaticText
      Left = 57
      Top = 163
      Width = 54
      Height = 20
      Caption = 'STB NO'
      TabOrder = 9
    end
    object StaticText2: TStaticText
      Left = 62
      Top = 196
      Width = 48
      Height = 20
      Caption = 'ICC NO'
      TabOrder = 10
    end
    object edtStbNo: TEdit
      Left = 115
      Top = 159
      Width = 121
      Height = 24
      Color = clInfoBk
      TabOrder = 0
    end
    object edtIccNo: TEdit
      Left = 115
      Top = 193
      Width = 121
      Height = 24
      Color = clInfoBk
      TabOrder = 1
    end
    object Lab_4: TStaticText
      Left = 42
      Top = 226
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
    object btnDepair: TBitBtn
      Left = 86
      Top = 347
      Width = 75
      Height = 25
      Caption = #38364#27231
      TabOrder = 4
      OnClick = btnDepairClick
    end
    object btnOpenCh: TBitBtn
      Left = 6
      Top = 387
      Width = 75
      Height = 25
      Caption = #38283#38971#36947'(7'#22825')'
      TabOrder = 6
      OnClick = btnOpenChClick
    end
    object btnCloseCh: TBitBtn
      Left = 86
      Top = 387
      Width = 75
      Height = 25
      Caption = #38364#38971#36947
      TabOrder = 7
      OnClick = btnCloseChClick
    end
    object Panel2: TPanel
      Left = 2
      Top = 2
      Width = 278
      Height = 143
      Align = alTop
      BevelOuter = bvLowered
      TabOrder = 12
      object Lab_1: TStaticText
        Left = 42
        Top = 19
        Width = 64
        Height = 20
        Caption = #36039#26009#24235#21029#21517
        TabOrder = 4
      end
      object Lab_2: TStaticText
        Left = 19
        Top = 51
        Width = 88
        Height = 20
        Caption = #36575#39636#20351#29992#32773#21517#31281
        TabOrder = 5
      end
      object Lab_3: TStaticText
        Left = 18
        Top = 86
        Width = 88
        Height = 20
        Caption = #36575#39636#20351#29992#32773#23494#30908
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
      object btnLogin: TBitBtn
        Left = 42
        Top = 113
        Width = 193
        Height = 25
        Caption = #30331#20837
        TabOrder = 3
        OnClick = btnLoginClick
      end
    end
    object memCh: TMemo
      Left = 43
      Top = 250
      Width = 193
      Height = 89
      Color = clInfoBk
      ReadOnly = True
      ScrollBars = ssVertical
      TabOrder = 13
    end
    object btnChooseCh: TBitBtn
      Left = 115
      Top = 225
      Width = 121
      Height = 22
      Caption = #40670#36984#38971#36947
      TabOrder = 2
      OnClick = btnChooseChClick
    end
    object btnPinCode: TBitBtn
      Left = 165
      Top = 347
      Width = 108
      Height = 25
      Caption = #35373#23450'PinCode'
      TabOrder = 5
      OnClick = btnPinCodeClick
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
    Left = 8
    Top = 64
  end
  object qryChInfo: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 8
    Top = 96
  end
  object qryActionTimes: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 8
    Top = 144
  end
  object MainMenu1: TMainMenu
    Left = 8
    Top = 224
    object mitFun1Header: TMenuItem
      Caption = #21151#33021'(&F)'
      object mitFun1: TMenuItem
        Caption = #20316#26989#27511#31243#34920'(&1)'
        OnClick = mitFun1Click
      end
      object mitFun2: TMenuItem
        Caption = #20351#29992#32773#36039#26009#32173#35703'(&2)'
        OnClick = mitFun2Click
      end
    end
  end
  object qryReport1: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from CaWareHouseLog')
    Left = 8
    Top = 184
    object qryReport1ICC_NO: TStringField
      FieldName = 'ICC_NO'
      Size = 12
    end
    object qryReport1STB_NO: TStringField
      FieldName = 'STB_NO'
      Size = 16
    end
    object qryReport1PRODUCTID: TStringField
      FieldName = 'PRODUCTID'
      Size = 3
    end
    object qryReport1ACTION: TStringField
      DisplayWidth = 12
      FieldName = 'ACTION'
      Size = 12
    end
    object qryReport1AUTHORSTARTDATE: TDateTimeField
      FieldName = 'AUTHORSTARTDATE'
    end
    object qryReport1AUTHORSTOPDATE: TDateTimeField
      FieldName = 'AUTHORSTOPDATE'
    end
    object qryReport1OPERATOR: TStringField
      FieldName = 'OPERATOR'
    end
    object qryReport1UPDTIME: TDateTimeField
      FieldName = 'UPDTIME'
    end
  end
end
