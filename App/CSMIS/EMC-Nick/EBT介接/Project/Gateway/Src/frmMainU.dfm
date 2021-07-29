object frmMain: TfrmMain
  Left = 423
  Top = 227
  Width = 721
  Height = 567
  BorderIcons = [biMinimize, biMaximize]
  Caption = 'frmMain'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poScreenCenter
  WindowState = wsMaximized
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object pnlProgStatus: TPanel
    Left = 0
    Top = 0
    Width = 713
    Height = 41
    Align = alTop
    BevelInner = bvLowered
    BevelOuter = bvNone
    Caption = #31995#32113#23578#26410#21855#21205
    Color = clSkyBlue
    Font.Charset = ANSI_CHARSET
    Font.Color = clRed
    Font.Height = -24
    Font.Name = #27161#26999#39636
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
    object chbActivate: TCheckBox
      Left = 17
      Top = 11
      Width = 86
      Height = 20
      Caption = #21855#21205
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -27
      Font.Name = #27161#26999#39636
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
      OnClick = chbActivateClick
    end
  end
  object memErrorLog: TMemo
    Left = 0
    Top = 380
    Width = 713
    Height = 122
    Align = alBottom
    Color = clInfoBk
    Font.Charset = ANSI_CHARSET
    Font.Color = clRed
    Font.Height = -13
    Font.Name = 'Courier New'
    Font.Style = []
    ParentFont = False
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 1
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 502
    Width = 713
    Height = 19
    Panels = <>
  end
  object pgcLV: TPageControl
    Left = 0
    Top = 41
    Width = 713
    Height = 339
    ActivePage = TabSheet1
    Align = alClient
    TabOrder = 3
    object TabSheet1: TTabSheet
      Caption = 'EBT'#23436#24037#36039#26009'(EBT->EMC)'
      object livStatus006: TListView
        Left = 0
        Top = 0
        Width = 705
        Height = 308
        Align = alClient
        Color = clInfoBk
        Columns = <
          item
            Caption = #24207#34399
            Width = 40
          end
          item
            Caption = #36039#26009#39006#21029
            Width = 90
          end
          item
            Caption = 'CATV'#23458#32232
            Width = 80
          end
          item
            Caption = 'EBT CATV ID'
            Width = 80
          end
          item
            Caption = #21512#32004#32232#34399
            Width = 120
          end
          item
            Caption = #23458#25142#32232#34399
            Width = 60
          end
          item
            Caption = #21512#32004#29983#25928#26085#26399
            Width = 70
          end
          item
            Caption = #21512#32004#25130#27490#26085#26399
            Width = 70
          end
          item
            Caption = #23458#25142#20013#25991#21517#31281
            Width = 80
          end
          item
            Caption = #23458#25142#32879#32097#38651#35441
            Width = 70
          end
          item
            Caption = #23458#25142#34892#21205#38651#35441
            Width = 70
          end
          item
            Caption = #20844#21496#36000#36012#20154#20013#25991#22995#21517
            Width = 120
          end
          item
            Caption = '('#20844#21496')'#32879#32097#20154#32879#32097#38651#35441
            Width = 100
          end
          item
            Caption = #25216#34899#32879#32097#20154#20013#25991#22995#21517
            Width = 100
          end
          item
            Caption = #25216#34899#32879#32097#20154#38651#35441
            Width = 100
          end
          item
            Caption = #25216#34899#32879#32097#20154#34892#21205#38651#35441
            Width = 100
          end
          item
            Caption = #25216#34899#32879#32097#20154'e-mail'
            Width = 100
          end
          item
            Caption = #35037#27231#22320#22336
            Width = 100
          end
          item
            Caption = #25142#31821#22320#22336
            Width = 100
          end
          item
            Caption = #24115#21934#22320#22336
            Width = 100
          end
          item
            Caption = #21512#32004#29376#24907
            Width = 100
          end
          item
            Caption = 'EBT'#32371#36027#36913#26399
            Width = 100
          end
          item
            Caption = #26381#21209#21029
            Width = 100
          end
          item
            Caption = #20195#29702#20154#22995#21517
            Width = 100
          end
          item
            Caption = #20195#29702#20154#38651#35441
            Width = 100
          end
          item
            Caption = #20195#29702#20154#36523#20221#35657#34399
            Width = 100
          end
          item
            Caption = #20195#29702#20154#22320#22336
            Width = 100
          end
          item
            Caption = #20491#20154#36523#20221#35657#34399
            Width = 100
          end
          item
            Caption = #36000#36012#20154#36523#20221#35657#34399' '
            Width = 100
          end
          item
            Caption = #38750#29151#21033#27861#20154#35657#34399
            Width = 100
          end
          item
            Caption = #20844#21496#35657#29031#34399#30908
            Width = 100
          end
          item
            Caption = #34389#29702#35387#35352
            Width = 100
          end
          item
            Caption = #37679#35492#20195#30908
            Width = 100
          end
          item
            Caption = #37679#35492#35338#24687
            Width = 100
          end
          item
            Caption = 'EBT'#20633#35387
            Width = 100
          end
          item
            Caption = 'EMC'#20633#35387
            Width = 100
          end
          item
            Caption = 'Catv'#23458#25142
            Width = 100
          end
          item
            Caption = #30064#21205#26178#38291
            Width = 100
          end
          item
            Caption = #30064#21205#20154#21729
            Width = 100
          end
          item
            Caption = #21514#29260#34399#30908
            Width = 100
          end>
        GridLines = True
        TabOrder = 0
        ViewStyle = vsReport
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'EBT'#30064#21205#36039#26009'(EBT->EMC)'
      ImageIndex = 1
      object livStatus007: TListView
        Left = 0
        Top = 0
        Width = 705
        Height = 308
        Align = alClient
        Color = clSkyBlue
        Columns = <
          item
            Caption = #24207#34399
            Width = 40
          end
          item
            Caption = #36039#26009#39006#21029
            Width = 90
          end
          item
            Caption = 'CATV'#23458#32232
            Width = 80
          end
          item
            Caption = 'EBT CATV ID'
            Width = 80
          end
          item
            Caption = #21512#32004#32232#34399
            Width = 120
          end
          item
            Caption = #30064#21205#24207#34399
          end
          item
            Caption = #23458#25142#32232#34399
            Width = 60
          end
          item
            Caption = #21512#32004#29983#25928#26085#26399
            Width = 70
          end
          item
            Caption = #21512#32004#25130#27490#26085#26399
            Width = 70
          end
          item
            Caption = #23458#25142#20013#25991#21517#31281
            Width = 80
          end
          item
            Caption = #23458#25142#32879#32097#38651#35441
            Width = 70
          end
          item
            Caption = #23458#25142#34892#21205#38651#35441
            Width = 70
          end
          item
            Caption = #20844#21496#36000#36012#20154#20013#25991#22995#21517
            Width = 120
          end
          item
            Caption = '('#20844#21496')'#32879#32097#20154#32879#32097#38651#35441
            Width = 100
          end
          item
            Caption = #25216#34899#32879#32097#20154#20013#25991#22995#21517
            Width = 100
          end
          item
            Caption = #25216#34899#32879#32097#20154#38651#35441
            Width = 100
          end
          item
            Caption = #25216#34899#32879#32097#20154#34892#21205#38651#35441
            Width = 100
          end
          item
            Caption = #25216#34899#32879#32097#20154'e-mail'
            Width = 100
          end
          item
            Caption = #35037#27231#22320#22336
            Width = 100
          end
          item
            Caption = #25142#31821#22320#22336
            Width = 100
          end
          item
            Caption = #24115#21934#22320#22336
            Width = 100
          end
          item
            Caption = #21512#32004#29376#24907
            Width = 100
          end
          item
            Caption = 'EBT'#32371#36027#36913#26399
            Width = 100
          end
          item
            Caption = #26381#21209#21029
            Width = 100
          end
          item
            Caption = #20195#29702#20154#22995#21517
            Width = 100
          end
          item
            Caption = #20195#29702#20154#38651#35441
            Width = 100
          end
          item
            Caption = #20195#29702#20154#36523#20221#35657#34399
            Width = 100
          end
          item
            Caption = #20195#29702#20154#22320#22336
            Width = 100
          end
          item
            Caption = #20491#20154#36523#20221#35657#34399
            Width = 100
          end
          item
            Caption = #36000#36012#20154#36523#20221#35657#34399' '
            Width = 100
          end
          item
            Caption = #38750#29151#21033#27861#20154#35657#34399
            Width = 100
          end
          item
            Caption = #20844#21496#35657#29031#34399#30908
            Width = 100
          end
          item
            Caption = #30064#21205#38917#30446
            Width = 100
          end
          item
            Caption = #33290'EMC'#23458#32232
            Width = 100
          end
          item
            Caption = #26159#21542#36328#31995#21488#31227#27231
            Width = 100
          end
          item
            Caption = #34389#29702#35387#35352
            Width = 100
          end
          item
            Caption = #37679#35492#20195#30908
            Width = 100
          end
          item
            Caption = #37679#35492#35338#24687
            Width = 100
          end
          item
            Caption = 'EBT'#20633#35387
            Width = 100
          end
          item
            Caption = 'Catv'#23458#25142
            Width = 100
          end
          item
            Caption = #30064#21205#26178#38291
            Width = 100
          end
          item
            Caption = #30064#21205#20154#21729
            Width = 100
          end
          item
            Caption = #21514#29260#34399#30908
            Width = 100
          end>
        GridLines = True
        TabOrder = 0
        ViewStyle = vsReport
      end
    end
    object TabSheet3: TTabSheet
      Caption = 'EMC'#23458#25142#30064#21205#36039#26009'(EMC->EBT)'
      ImageIndex = 2
      object livStatusSO151: TListView
        Left = 0
        Top = 0
        Width = 705
        Height = 308
        Align = alClient
        Color = clMedGray
        Columns = <
          item
            Caption = #24207#34399
            Width = 40
          end
          item
            Caption = #36039#26009#24207#34399
            Width = 80
          end
          item
            Caption = 'EMC'#20844#21496#21029
            Width = 90
          end
          item
            Caption = 'EMC'#23458#32232
            Width = 80
          end
          item
            Caption = #26032#23458#25142#29376#24907
            Width = 100
          end
          item
            Caption = #33290#23458#25142#29376#24907
            Width = 100
          end
          item
            Caption = 'EBT'#23458#32232
            Width = 80
          end
          item
            Caption = 'EBT'#21512#32004#32232#34399
            Width = 100
          end
          item
            Caption = #26381#21209#39006#21029
            Width = 70
          end
          item
            Caption = #34389#29702#35387#35352
          end
          item
            Caption = #30064#21205#26178#38291
            Width = 100
          end
          item
            Caption = #30064#21205#20154#21729
            Width = 100
          end>
        GridLines = True
        TabOrder = 0
        ViewStyle = vsReport
      end
    end
    object TabSheet4: TTabSheet
      Caption = 'EMC'#23458#32232#36070#20104#23436#25104#36039#26009'(EMC->EBT)'
      ImageIndex = 3
      object livStatusSO153: TListView
        Left = 0
        Top = 0
        Width = 705
        Height = 308
        Align = alClient
        Color = clTeal
        Columns = <
          item
            Caption = #24207#34399
            Width = 40
          end
          item
            Caption = #36039#26009#39006#21029
            Width = 90
          end
          item
            Caption = 'CATV'#23458#32232
            Width = 80
          end
          item
            Caption = 'EBT CATV ID'
            Width = 80
          end
          item
            Caption = #21512#32004#32232#34399
            Width = 120
          end
          item
            Caption = #23458#25142#32232#34399
            Width = 60
          end
          item
            Caption = #21512#32004#29983#25928#26085#26399
            Width = 70
          end
          item
            Caption = #21512#32004#25130#27490#26085#26399
            Width = 70
          end
          item
            Caption = #23458#25142#20013#25991#21517#31281
            Width = 80
          end
          item
            Caption = #23458#25142#32879#32097#38651#35441
            Width = 70
          end
          item
            Caption = #23458#25142#34892#21205#38651#35441
            Width = 70
          end
          item
            Caption = #20844#21496#36000#36012#20154#20013#25991#22995#21517
            Width = 120
          end
          item
            Caption = '('#20844#21496')'#32879#32097#20154#32879#32097#38651#35441
            Width = 100
          end
          item
            Caption = #25216#34899#32879#32097#20154#20013#25991#22995#21517
            Width = 100
          end
          item
            Caption = #25216#34899#32879#32097#20154#38651#35441
            Width = 100
          end
          item
            Caption = #25216#34899#32879#32097#20154#34892#21205#38651#35441
            Width = 100
          end
          item
            Caption = #25216#34899#32879#32097#20154'e-mail'
            Width = 100
          end
          item
            Caption = #35037#27231#22320#22336
            Width = 100
          end
          item
            Caption = #25142#31821#22320#22336
            Width = 100
          end
          item
            Caption = #24115#21934#22320#22336
            Width = 100
          end
          item
            Caption = #21512#32004#29376#24907
            Width = 100
          end
          item
            Caption = 'EBT'#32371#36027#36913#26399
            Width = 100
          end
          item
            Caption = #26381#21209#21029
            Width = 100
          end
          item
            Caption = #20195#29702#20154#22995#21517
            Width = 100
          end
          item
            Caption = #20195#29702#20154#38651#35441
            Width = 100
          end
          item
            Caption = #20195#29702#20154#36523#20221#35657#34399
            Width = 100
          end
          item
            Caption = #20195#29702#20154#22320#22336
            Width = 100
          end
          item
            Caption = #20491#20154#36523#20221#35657#34399
            Width = 100
          end
          item
            Caption = #36000#36012#20154#36523#20221#35657#34399' '
            Width = 100
          end
          item
            Caption = #38750#29151#21033#27861#20154#35657#34399
            Width = 100
          end
          item
            Caption = #20844#21496#35657#29031#34399#30908
            Width = 100
          end
          item
            Caption = #34389#29702#35387#35352
            Width = 100
          end
          item
            Caption = #37679#35492#20195#30908
            Width = 100
          end
          item
            Caption = #37679#35492#35338#24687
            Width = 100
          end
          item
            Caption = 'EBT'#20633#35387
            Width = 100
          end
          item
            Caption = 'EMC'#20633#35387
            Width = 100
          end
          item
            Caption = #30064#21205#26178#38291
            Width = 100
          end
          item
            Caption = #30064#21205#20154#21729
            Width = 100
          end>
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clYellow
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        GridLines = True
        ParentFont = False
        TabOrder = 0
        ViewStyle = vsReport
      end
    end
  end
  object MainMenu1: TMainMenu
    Left = 236
    Top = 8
    object N11: TMenuItem
      Caption = '&1. '#20854#20182#21151#33021
      object A1: TMenuItem
        Caption = '&A. '#21443#25976#35373#23450
        OnClick = A1Click
      end
      object B1: TMenuItem
        Caption = '&B. '#20351#29992#32773#32173#35703
        OnClick = B1Click
      end
      object N1: TMenuItem
        Caption = '-'
      end
      object X1: TMenuItem
        Caption = '&X. '#32080#26463#31243#24335
        OnClick = X1Click
      end
    end
  end
  object timPeriodCount: TTimer
    OnTimer = timPeriodCountTimer
    Left = 276
    Top = 8
  end
  object ADOQuery1: TADOQuery
    Parameters = <>
    Left = 320
    Top = 8
  end
  object timProgStatus: TTimer
    Enabled = False
    Interval = 5000
    OnTimer = timProgStatusTimer
    Left = 364
    Top = 8
  end
end
