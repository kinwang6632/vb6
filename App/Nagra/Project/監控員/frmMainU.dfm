object frmMain: TfrmMain
  Left = 344
  Top = 227
  Width = 556
  Height = 437
  BorderIcons = []
  Caption = 'CA '#30435#25511#21729
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 548
    Height = 410
    Align = alClient
    BevelInner = bvLowered
    BevelOuter = bvNone
    Color = clSkyBlue
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clTeal
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    object memInfo: TMemo
      Left = 1
      Top = 153
      Width = 546
      Height = 237
      Align = alClient
      Color = clSkyBlue
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      ScrollBars = ssVertical
      TabOrder = 0
    end
    object Panel2: TPanel
      Left = 1
      Top = 1
      Width = 546
      Height = 40
      Align = alTop
      Caption = 'CA '#30435#25511#21729
      Color = clInfoBk
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
    end
    object Panel3: TPanel
      Left = 1
      Top = 113
      Width = 546
      Height = 40
      Align = alTop
      BevelInner = bvRaised
      BevelOuter = bvNone
      BorderStyle = bsSingle
      TabOrder = 2
      object CheckBox1: TCheckBox
        Left = 24
        Top = 10
        Width = 121
        Height = 17
        Caption = #32178#36335#30435#25511
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -16
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        Visible = False
      end
      object CheckBox2: TCheckBox
        Left = 166
        Top = 10
        Width = 131
        Height = 17
        Caption = 'CA'#22519#34892#32210#30435#25511
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -16
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnClick = CheckBox2Click
      end
      object chbAcceptClient: TCheckBox
        Left = 328
        Top = 10
        Width = 150
        Height = 17
        Caption = #25509#21463' client '#31471#25511#21046
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -16
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
      end
    end
    object Panel4: TPanel
      Left = 1
      Top = 41
      Width = 546
      Height = 72
      Align = alTop
      BevelInner = bvRaised
      BevelOuter = bvLowered
      TabOrder = 3
      object SpeedButton1: TSpeedButton
        Left = 359
        Top = 8
        Width = 25
        Height = 24
        Caption = '...'
        OnClick = SpeedButton1Click
      end
      object BitBtn3: TBitBtn
        Left = 389
        Top = 8
        Width = 105
        Height = 25
        Caption = #37325#35373#30435#25511#21443#25976
        TabOrder = 1
        OnClick = BitBtn3Click
      end
      object BitBtn2: TBitBtn
        Left = 297
        Top = 40
        Width = 113
        Height = 25
        Caption = #36681#20986#30435#25511#36039#26009
        TabOrder = 3
        OnClick = BitBtn2Click
      end
      object BitBtn1: TBitBtn
        Left = 419
        Top = 40
        Width = 75
        Height = 25
        Caption = #32080#26463#31243#24335
        TabOrder = 4
        OnClick = BitBtn1Click
      end
      object edtIniFileName: TEdit
        Left = 136
        Top = 6
        Width = 218
        Height = 28
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -16
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
      object StaticText1: TStaticText
        Left = 39
        Top = 8
        Width = 96
        Height = 24
        Caption = #30435#25511#21443#25976#27284' : '
        TabOrder = 5
      end
      object BitBtn4: TBitBtn
        Left = 174
        Top = 40
        Width = 113
        Height = 25
        Caption = #28165#38500#30435#25511#36039#26009
        TabOrder = 2
        OnClick = BitBtn4Click
      end
    end
    object stb: TStatusBar
      Left = 1
      Top = 390
      Width = 546
      Height = 19
      Panels = <
        item
          Width = 200
        end>
    end
  end
  object SaveDialog1: TSaveDialog
    Left = 392
    Top = 8
  end
  object timAutoInvoke: TTimer
    Enabled = False
    OnTimer = timAutoInvokeTimer
    Left = 288
    Top = 8
  end
  object ServerSocket: TServerSocket
    Active = False
    Port = 1024
    ServerType = stNonBlocking
    OnAccept = ServerSocketAccept
    OnClientDisconnect = ServerSocketClientDisconnect
    OnClientRead = ServerSocketClientRead
    Left = 187
    Top = 70
  end
  object OpenDialog1: TOpenDialog
    Filter = '*.ini|*.ini'
    Left = 432
    Top = 8
  end
  object cdsParam: TClientDataSet
    Aggregates = <>
    FieldDefs = <>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 504
    Top = 48
    object cdsParamsServerIP: TStringField
      FieldName = 'sServerIP'
      Size = 15
    end
    object cdsParamnSPortNo: TIntegerField
      FieldName = 'nSPortNo'
    end
    object cdsParamnRPortNo: TIntegerField
      FieldName = 'nRPortNo'
    end
    object cdsParamsSysName: TStringField
      FieldName = 'sSysName'
      Size = 40
    end
    object cdsParamsSysVersion: TStringField
      FieldName = 'sSysVersion'
      Size = 5
    end
    object cdsParamsMisCommandPath: TStringField
      FieldName = 'sLogPath'
      Size = 60
    end
    object cdsParamnTimeOut: TIntegerField
      FieldName = 'nTimeOut'
    end
    object cdsParamnMaxCommandCount: TIntegerField
      FieldName = 'nMaxCommandCount'
    end
    object cdsParambCommandLog: TBooleanField
      FieldName = 'bCommandLog'
    end
    object cdsParambErrorLog: TBooleanField
      FieldName = 'bErrorLog'
    end
    object cdsParamnResponseLog: TBooleanField
      FieldName = 'nResponseLog'
    end
    object cdsParamdUptTime: TDateTimeField
      FieldName = 'dUptTime'
    end
    object cdsParamsUptName: TStringField
      FieldName = 'sUptName'
    end
    object cdsParambSetZipCode: TBooleanField
      FieldName = 'bSetZipCode'
    end
    object cdsParamnCreditLimit: TIntegerField
      FieldName = 'nCreditLimit'
    end
    object cdsParamnSourceID: TIntegerField
      FieldName = 'nSourceID'
    end
    object cdsParamnDestID: TIntegerField
      FieldName = 'nDestID'
    end
    object cdsParamnMopPPID: TIntegerField
      FieldName = 'nMopPPID'
    end
    object cdsParamsBroadcastSDate: TStringField
      FieldName = 'sBroadcastSDate'
      EditMask = '!9999/99/99;1;_'
      Size = 10
    end
    object cdsParamsBroadcastEDate: TStringField
      FieldName = 'sBroadcastEDate'
      EditMask = '!9999/99/99;1;_'
      Size = 10
    end
    object cdsParamnCmdRefreshRate: TIntegerField
      FieldName = 'nCmdRefreshRate1'
    end
    object cdsParamnCmdRefreshRate2: TIntegerField
      FieldName = 'nCmdRefreshRate2'
    end
    object cdsParamnCmdResentTimes: TIntegerField
      FieldName = 'nCmdResentTimes'
    end
    object cdsParambShowUI: TBooleanField
      FieldName = 'bShowUI'
    end
    object cdsParambAssignBroadcastDate: TBooleanField
      FieldName = 'bAssignBroadcastDate'
    end
    object cdsParamsSHotTime: TStringField
      FieldName = 'sSHotTime'
      Size = 4
    end
    object cdsParamsEHotTime: TStringField
      FieldName = 'sEHotTime'
      Size = 4
    end
  end
end
