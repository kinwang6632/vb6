object frmSimpleSGS: TfrmSimpleSGS
  Left = 101
  Top = 210
  Width = 864
  Height = 339
  Caption = 'frmSimpleSGS'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 856
    Height = 33
    Align = alTop
    TabOrder = 0
    object cbActive: TCheckBox
      Left = 14
      Top = 9
      Width = 97
      Height = 17
      Caption = '&Activate'
      TabOrder = 0
      OnClick = acActivateExecute
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 33
    Width = 856
    Height = 279
    Align = alClient
    Caption = 'Panel2'
    TabOrder = 1
    object Memo1: TMemo
      Left = 1
      Top = 81
      Width = 854
      Height = 176
      Align = alTop
      Color = clSkyBlue
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
    end
    object Panel3: TPanel
      Left = 1
      Top = 1
      Width = 854
      Height = 80
      Align = alTop
      TabOrder = 1
      object Label1: TLabel
        Left = 24
        Top = 8
        Width = 32
        Height = 13
        Alignment = taRightJustify
        Caption = 'Label1'
      end
      object Label2: TLabel
        Left = 24
        Top = 40
        Width = 32
        Height = 13
        Alignment = taRightJustify
        Caption = 'Label1'
      end
      object Label3: TLabel
        Left = 437
        Top = 8
        Width = 32
        Height = 13
        Alignment = taRightJustify
        Caption = 'Label1'
      end
      object cobCompCode: TComboBox
        Left = 72
        Top = 8
        Width = 145
        Height = 21
        ItemHeight = 13
        TabOrder = 0
        Text = 'cobCompCode'
      end
      object cobQueryType: TComboBox
        Left = 72
        Top = 40
        Width = 145
        Height = 21
        ItemHeight = 13
        TabOrder = 1
        Text = 'cobQueryType'
      end
      object chkNow: TCheckBox
        Left = 272
        Top = 8
        Width = 97
        Height = 17
        Caption = #31435#21363#26597#35426
        TabOrder = 2
        OnClick = chkNowClick
      end
      inline fraYMD1: TfraYMD
        Left = 496
        Top = 0
        Width = 81
        Height = 35
        TabOrder = 3
      end
      object btnSend: TButton
        Left = 680
        Top = 8
        Width = 75
        Height = 25
        Caption = #20659#36865
        TabOrder = 4
        OnClick = btnSendClick
      end
      object btnExit: TButton
        Left = 768
        Top = 8
        Width = 75
        Height = 25
        Caption = #32080#26463
        TabOrder = 5
        OnClick = btnExitClick
      end
    end
    object lbLog: TListBox
      Left = 1
      Top = 257
      Width = 854
      Height = 21
      Align = alClient
      ItemHeight = 13
      TabOrder = 2
    end
  end
  object HTTPServer: TIdHTTPServer
    Bindings = <>
    OnCommandGet = HTTPServerCommandGet
    AutoStartSession = True
    SessionTimeOut = 1200000
    Left = 136
    Top = 8
  end
  object alGeneral: TActionList
    Left = 168
    Top = 8
    object acActivate: TAction
      Caption = '&Activate'
    end
  end
  object HTTP: TIdHTTP
    RecvBufferSize = 1024
    HandleRedirects = True
    ProtocolVersion = pv1_0
    Request.Accept = 'text/html, */*'
    Request.ContentLength = 0
    Request.ContentRangeEnd = 0
    Request.ContentRangeStart = 0
    Request.Password = '1'
    Request.ProxyPort = 8080
    Request.ProxyServer = '212.56.15.196'
    Request.ProxyUsername = '1'
    Request.UserAgent = 'Mozilla/3.0 (compatible; Indy Library)'
    Left = 200
    Top = 8
  end
  object setXMLDocument: TXMLDocument
    Left = 232
    Top = 9
    DOMVendorDesc = 'MSXML'
  end
end
