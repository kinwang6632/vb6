object fmMain: TfmMain
  Left = 374
  Top = 269
  Width = 618
  Height = 538
  Caption = #38283#21338#31185#25216' - SMS Command Gateway 2005'
  Color = clBtnFace
  Font.Charset = ARABIC_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object LedPanel: TPanel
    Left = 100
    Top = 104
    Width = 46
    Height = 22
    BevelOuter = bvNone
    TabOrder = 1
    Visible = False
    object Image1: TImage
      Left = 0
      Top = 0
      Width = 18
      Height = 18
    end
    object Image2: TImage
      Left = 21
      Top = 0
      Width = 18
      Height = 18
    end
  end
  object dxDockHost: TdxDockSite
    Left = 0
    Top = 85
    Width = 610
    Height = 400
    Align = alClient
    DockType = 0
    OriginalWidth = 610
    OriginalHeight = 400
    object dxLayoutDockSite2: TdxLayoutDockSite
      Left = 250
      Top = 0
      Width = 360
      Height = 400
      DockType = 1
      OriginalWidth = 300
      OriginalHeight = 200
      object dxLayoutDockSite3: TdxLayoutDockSite
        Left = 0
        Top = 0
        Width = 360
        Height = 228
        DockType = 1
        OriginalWidth = 300
        OriginalHeight = 200
        object dxLayoutDockSite1: TdxLayoutDockSite
          Left = 0
          Top = 0
          Width = 360
          Height = 228
          DockType = 1
          OriginalWidth = 300
          OriginalHeight = 200
        end
        object dxTabContainerDockSite3: TdxTabContainerDockSite
          Left = 0
          Top = 0
          Width = 360
          Height = 228
          ActiveChildIndex = 0
          AllowFloating = True
          AutoHide = False
          CaptionButtons = []
          DockType = 1
          OriginalWidth = 185
          OriginalHeight = 124
          object dxDockPanel3: TdxDockPanel
            Left = 0
            Top = 0
            Width = 356
            Height = 178
            AllowFloating = True
            AutoHide = False
            Caption = 'dxDockPanel3'
            CaptionButtons = []
            ImageIndex = 2
            DockType = 1
            OriginalWidth = 185
            OriginalHeight = 124
            object ControlSendTree: TcxTreeList
              Left = 0
              Top = 0
              Width = 356
              Height = 178
              Styles.Content = StyleModule.cxStyle1
              Align = alClient
              Bands = <>
              BufferedPaint = False
              Images = StyleModule.TreeImageList
              OptionsData.Editing = False
              OptionsData.Deleting = False
              OptionsSelection.CellSelect = False
              StateImages = StyleModule.TreeImageList
              TabOrder = 0
            end
          end
          object dxDockPanel4: TdxDockPanel
            Left = 0
            Top = 0
            Width = 356
            Height = 178
            AllowFloating = True
            AutoHide = False
            Caption = 'dxDockPanel4'
            CaptionButtons = []
            ImageIndex = 3
            DockType = 1
            OriginalWidth = 185
            OriginalHeight = 140
          end
        end
      end
      object dxTabContainerDockSite2: TdxTabContainerDockSite
        Left = 0
        Top = 228
        Width = 360
        Height = 172
        ActiveChildIndex = 0
        AllowFloating = True
        AutoHide = False
        DockType = 5
        OriginalWidth = 185
        OriginalHeight = 172
        object dxDockPanel5: TdxDockPanel
          Left = 0
          Top = 0
          Width = 356
          Height = 122
          AllowFloating = False
          AllowDock = [dtLeft, dtTop, dtRight, dtBottom]
          AutoHide = False
          Caption = 'dxDockPanel5'
          ImageIndex = 1
          DockType = 1
          OriginalWidth = 185
          OriginalHeight = 172
          object cxPageControl1: TcxPageControl
            Left = 0
            Top = 0
            Width = 356
            Height = 122
            Align = alClient
            TabOrder = 0
            TabPosition = tpBottom
            ClientRectBottom = 122
            ClientRectRight = 356
            ClientRectTop = 0
          end
          object StartupMsgList: TcxListView
            Left = 0
            Top = 0
            Width = 356
            Height = 122
            Align = alClient
            ColumnClick = False
            Columns = <
              item
                Width = 1000
              end>
            ParentFont = False
            ReadOnly = True
            RowSelect = True
            ShowColumnHeaders = False
            SmallImages = StyleModule.MsgImageList
            Style.StyleController = StyleModule.CommonStyle
            StyleDisabled.StyleController = StyleModule.CommonStyle
            StyleFocused.StyleController = StyleModule.CommonStyle
            StyleHot.StyleController = StyleModule.CommonStyle
            TabOrder = 1
            ViewStyle = vsReport
          end
        end
        object dxDockPanel6: TdxDockPanel
          Left = 0
          Top = 0
          Width = 356
          Height = 122
          AllowFloating = True
          AutoHide = False
          Caption = 'dxDockPanel6'
          ImageIndex = 1
          DockType = 1
          OriginalWidth = 175
          OriginalHeight = 172
          object SoDbMsgList: TcxListView
            Left = 0
            Top = 0
            Width = 356
            Height = 122
            Align = alClient
            ColumnClick = False
            Columns = <
              item
                Width = 1000
              end>
            ParentFont = False
            ReadOnly = True
            RowSelect = True
            ShowColumnHeaders = False
            SmallImages = StyleModule.MsgImageList
            Style.StyleController = StyleModule.CommonStyle
            StyleDisabled.StyleController = StyleModule.CommonStyle
            StyleFocused.StyleController = StyleModule.CommonStyle
            StyleHot.StyleController = StyleModule.CommonStyle
            TabOrder = 0
            ViewStyle = vsReport
          end
        end
        object dxDockPanel7: TdxDockPanel
          Left = 0
          Top = 0
          Width = 356
          Height = 122
          AllowFloating = True
          AutoHide = False
          Caption = 'dxDockPanel7'
          ImageIndex = 1
          DockType = 1
          OriginalWidth = 185
          OriginalHeight = 172
          object ControlMsgList: TcxListView
            Left = 0
            Top = 0
            Width = 356
            Height = 122
            Align = alClient
            ColumnClick = False
            Columns = <
              item
                Width = 1000
              end>
            ParentFont = False
            ReadOnly = True
            RowSelect = True
            ShowColumnHeaders = False
            SmallImages = StyleModule.MsgImageList
            Style.StyleController = StyleModule.CommonStyle
            StyleDisabled.StyleController = StyleModule.CommonStyle
            StyleFocused.StyleController = StyleModule.CommonStyle
            StyleHot.StyleController = StyleModule.CommonStyle
            TabOrder = 0
            ViewStyle = vsReport
          end
        end
      end
    end
    object dxTabContainerDockSite1: TdxTabContainerDockSite
      Left = 0
      Top = 0
      Width = 250
      Height = 400
      ActiveChildIndex = 0
      AllowFloating = False
      AutoHide = False
      DockType = 2
      OriginalWidth = 250
      OriginalHeight = 140
      object dxDockPanel1: TdxDockPanel
        Left = 0
        Top = 0
        Width = 246
        Height = 350
        AllowFloating = True
        AutoHide = False
        Caption = 'dxDockPanel1'
        ImageIndex = 0
        DockType = 1
        OriginalWidth = 250
        OriginalHeight = 309
        object SoTree: TcxTreeList
          Left = 0
          Top = 0
          Width = 246
          Height = 350
          Styles.Content = StyleModule.cxStyle1
          Align = alClient
          Bands = <
            item
            end>
          BufferedPaint = False
          Images = StyleModule.TreeImageList
          LookAndFeel.Kind = lfOffice11
          OptionsBehavior.Sorting = False
          OptionsData.Editing = False
          OptionsData.Deleting = False
          OptionsSelection.CellSelect = False
          OptionsView.CellEndEllipsis = True
          OptionsView.CategorizedColumn = SoTreecxTreeListColumn2
          OptionsView.PaintStyle = tlpsCategorized
          StateImages = StyleModule.TreeImageList
          TabOrder = 0
          object SoTreecxTreeListColumn1: TcxTreeListColumn
            Caption.Text = #20195#30908
            DataBinding.ValueType = 'String'
            Width = 110
            Position.ColIndex = 0
            Position.RowIndex = 0
            Position.BandIndex = 0
          end
          object SoTreecxTreeListColumn2: TcxTreeListColumn
            Caption.Text = #31995#32113#21488
            DataBinding.ValueType = 'String'
            Width = 75
            Position.ColIndex = 1
            Position.RowIndex = 0
            Position.BandIndex = 0
          end
          object SoTreecxTreeListColumn3: TcxTreeListColumn
            Caption.Text = #32047#35336#25351#20196
            DataBinding.ValueType = 'String'
            Width = 56
            Position.ColIndex = 3
            Position.RowIndex = 0
            Position.BandIndex = 0
          end
        end
      end
      object dxDockPanel2: TdxDockPanel
        Left = 0
        Top = 0
        Width = 246
        Height = 350
        AllowFloating = True
        AutoHide = False
        Caption = 'dxDockPanel2'
        DockType = 1
        OriginalWidth = 250
        OriginalHeight = 309
      end
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 49
    Width = 610
    Height = 36
    Align = alTop
    BevelOuter = bvNone
    BorderWidth = 2
    TabOrder = 6
    object Panel2: TPanel
      Left = 2
      Top = 2
      Width = 606
      Height = 32
      Align = alClient
      BevelOuter = bvNone
      Color = clAppWorkSpace
      TabOrder = 0
      object lblConfigFileName: TcxLabel
        Left = 11
        Top = 7
        Width = 113
        Height = 20
        Caption = 'lblConfigFileName'
        ParentFont = False
        Style.StyleController = StyleModule.TitleCaptionStyle
        StyleDisabled.StyleController = StyleModule.TitleCaptionStyle
        StyleFocused.StyleController = StyleModule.TitleCaptionStyle
        StyleHot.StyleController = StyleModule.TitleCaptionStyle
      end
    end
  end
  object dxBarManager: TdxBarManager
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    Bars = <
      item
        AllowClose = False
        AllowCustomizing = False
        AllowQuickCustomizing = False
        AllowReset = False
        BorderStyle = bbsNone
        Caption = #29376#24907#21015
        DockedDockingStyle = dsBottom
        DockedLeft = 0
        DockedTop = 0
        DockingStyle = dsBottom
        FloatLeft = 788
        FloatTop = 369
        FloatClientWidth = 70
        FloatClientHeight = 70
        ItemLinks = <
          item
            Item = dxComputer
            Visible = True
          end
          item
            Item = dxServer
            Visible = True
          end>
        Name = 'StatusBar'
        NotDocking = [dsNone]
        OneOnRow = True
        Row = 0
        UseOwnFont = False
        Visible = True
        WholeRow = True
      end
      item
        AllowClose = False
        AllowCustomizing = False
        AllowQuickCustomizing = False
        AllowReset = False
        Caption = #21151#33021#34920
        DockedDockingStyle = dsTop
        DockedLeft = 0
        DockedTop = 0
        DockingStyle = dsTop
        FloatLeft = 464
        FloatTop = 357
        FloatClientWidth = 23
        FloatClientHeight = 22
        IsMainMenu = True
        ItemLinks = <
          item
            Item = dxFiles
            Visible = True
          end
          item
            Item = dxAbout
            Visible = True
          end>
        MultiLine = True
        Name = 'MenuBar'
        NotDocking = [dsNone, dsLeft, dsRight, dsBottom]
        OneOnRow = True
        Row = 0
        UseOwnFont = False
        Visible = True
        WholeRow = True
      end
      item
        AllowClose = False
        AllowCustomizing = False
        AllowQuickCustomizing = False
        AllowReset = False
        Caption = #24037#20855#21015
        DockedDockingStyle = dsTop
        DockedLeft = 0
        DockedTop = 23
        DockingStyle = dsTop
        FloatLeft = 464
        FloatTop = 357
        FloatClientWidth = 23
        FloatClientHeight = 22
        ItemLinks = <
          item
            Item = dxPlay
            Visible = True
          end
          item
            Item = dxPause
            Visible = True
          end
          item
            Item = dxStop
            Visible = True
          end
          item
            Item = dxReload
            Visible = True
          end>
        Name = 'ToolBar'
        NotDocking = [dsLeft, dsRight, dsBottom]
        OneOnRow = True
        Row = 1
        UseOwnFont = False
        Visible = True
        WholeRow = True
      end>
    Categories.Strings = (
      'Menus'
      'Status'
      'Tools')
    Categories.ItemsVisibles = (
      2
      2
      2)
    Categories.Visibles = (
      True
      True
      True)
    HotImages = StyleModule.BarImageList
    Images = StyleModule.BarImageList
    LargeImages = StyleModule.BarImageList
    PopupMenuLinks = <>
    Style = bmsOffice11
    UseSystemFont = True
    Left = 68
    Top = 340
    DockControlHeights = (
      0
      0
      49
      26)
    object dxFiles: TdxBarSubItem
      Caption = #27284#26696
      Category = 0
      Visible = ivAlways
      ItemLinks = <>
    end
    object dxAbout: TdxBarSubItem
      Caption = #38364#26044
      Category = 0
      Visible = ivAlways
      ItemLinks = <>
    end
    object dxPlay: TdxBarButton
      Caption = #21855#21205
      Category = 2
      Hint = #21855#21205
      Visible = ivAlways
      ImageIndex = 0
      OnClick = dxPlayClick
    end
    object dxPause: TdxBarButton
      Caption = #26283#20572
      Category = 2
      Enabled = False
      Hint = #26283#20572
      Visible = ivAlways
      ImageIndex = 1
    end
    object dxStop: TdxBarButton
      Caption = #20572#27490
      Category = 2
      Enabled = False
      Hint = #20572#27490
      Visible = ivAlways
      ImageIndex = 2
      OnClick = dxStopClick
    end
    object dxReload: TdxBarButton
      Caption = #37325#26032#21855#21205
      Category = 2
      Enabled = False
      Hint = #37325#26032#21855#21205
      Visible = ivAlways
      ImageIndex = 3
    end
    object dxComputer: TdxBarStatic
      Caption = #26412#27231
      Category = 1
      Hint = #26412#27231
      Visible = ivAlways
      BorderStyle = sbsLowered
      ImageIndex = 4
    end
    object dxServer: TdxBarStatic
      Caption = #20282#26381#22120
      Category = 1
      Hint = #20282#26381#22120
      Visible = ivAlways
      BorderStyle = sbsLowered
      ImageIndex = 5
      LeftIndent = 5
    end
    object dxLED: TdxBarControlContainerItem
      Caption = #29376#24907
      Category = 1
      Hint = #29376#24907
      Visible = ivAlways
      Control = LedPanel
    end
    object PlayCmdGroup: TdxBarGroup
      Items = (
        'dxPlay')
    end
    object StopCmdGroup: TdxBarGroup
      Enabled = False
      Items = (
        'dxPause'
        'dxStop'
        'dxReload')
    end
  end
  object dxDockingManager1: TdxDockingManager
    Color = clBtnFace
    DefaultHorizContainerSiteProperties.Dockable = True
    DefaultHorizContainerSiteProperties.ImageIndex = -1
    DefaultVertContainerSiteProperties.Dockable = True
    DefaultVertContainerSiteProperties.ImageIndex = -1
    DefaultTabContainerSiteProperties.Dockable = True
    DefaultTabContainerSiteProperties.ImageIndex = -1
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = []
    Images = StyleModule.DockPanelImageList
    Options = [doActivateAfterDocking, doDblClickDocking, doFloatingOnTop, doTabContainerHasCaption, doTabContainerCanAutoHide, doSideContainerCanClose, doSideContainerCanAutoHide, doTabContainerCanInSideContainer]
    ViewStyle = vsOffice11
    Left = 36
    Top = 340
  end
  object AutoExecTimer: TTimer
    Enabled = False
    Interval = 2200
    OnTimer = AutoExecTimerTimer
    Left = 101
    Top = 340
  end
  object StartupTimer: TTimer
    Enabled = False
    Interval = 1200
    OnTimer = StartupTimerTimer
    Left = 135
    Top = 340
  end
  object ControlSendSocket: TIdTCPClient
    MaxLineAction = maException
    Port = 0
    Left = 171
    Top = 339
  end
end
