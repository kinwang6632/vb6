object fmMain: TfmMain
  Left = 257
  Top = 112
  Width = 763
  Height = 608
  Caption = 'SMS '#31680#30446#34920#21295#25972
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object dxDockMain: TdxDockSite
    Left = 0
    Top = 52
    Width = 755
    Height = 500
    Align = alClient
    DockType = 0
    OriginalWidth = 755
    OriginalHeight = 500
    object dxLayoutDockSite1: TdxLayoutDockSite
      Left = 120
      Top = 0
      Width = 635
      Height = 500
      DockType = 1
      OriginalWidth = 300
      OriginalHeight = 200
      object dxLayoutDockSite4: TdxLayoutDockSite
        Left = 0
        Top = 0
        Width = 635
        Height = 362
        DockType = 1
        OriginalWidth = 300
        OriginalHeight = 200
        object dxLayoutDockSite2: TdxLayoutDockSite
          Left = 0
          Top = 0
          Width = 635
          Height = 362
          DockType = 1
          OriginalWidth = 300
          OriginalHeight = 200
        end
        object dxDockClient: TdxDockPanel
          Left = 0
          Top = 0
          Width = 635
          Height = 362
          AllowFloating = False
          AutoHide = False
          Caption = #21295#25972#27511#31243
          CaptionButtons = [cbHide]
          ShowCaption = False
          DockType = 1
          OriginalWidth = 185
          OriginalHeight = 140
          object Bevel1: TBevel
            Left = 0
            Top = 31
            Width = 631
            Height = 4
            Align = alTop
            Shape = bsSpacer
          end
          object RzTitlePanel: TRzPanel
            Left = 0
            Top = 0
            Width = 631
            Height = 31
            Align = alTop
            BorderOuter = fsFlat
            Color = 8551805
            FlatColor = clGray
            TabOrder = 0
            DesignSize = (
              631
              31)
            object PageMarkImage: TImage
              Left = 593
              Top = 4
              Width = 24
              Height = 24
              Anchors = [akTop, akRight]
              Stretch = True
            end
            object lblTitle: TLabel
              Left = 16
              Top = 8
              Width = 41
              Height = 16
              Caption = 'lblTitle'
              Font.Charset = DEFAULT_CHARSET
              Font.Color = clWhite
              Font.Height = -13
              Font.Name = 'Tahoma'
              Font.Style = [fsBold]
              ParentFont = False
            end
          end
          object MainPage: TcxPageControl
            Left = 0
            Top = 35
            Width = 631
            Height = 323
            ActivePage = cxTabAction
            Align = alClient
            TabOrder = 1
            TabPosition = tpBottom
            OnChange = MainPageChange
            ClientRectBottom = 298
            ClientRectRight = 631
            ClientRectTop = 0
            object cxTabAction: TcxTabSheet
              Caption = 'cxTabAction'
              ImageIndex = 0
              object ActionGrid: TcxGrid
                Left = 0
                Top = 0
                Width = 631
                Height = 298
                Align = alClient
                TabOrder = 0
                object gvAction: TcxGridDBTableView
                  NavigatorButtons.ConfirmDelete = False
                  OnCustomDrawCell = gvActionCustomDrawCell
                  DataController.DataModeController.SyncMode = False
                  DataController.DataSource = dsAction
                  DataController.KeyFieldNames = 'KeyId'
                  DataController.Summary.DefaultGroupSummaryItems = <>
                  DataController.Summary.FooterSummaryItems = <>
                  DataController.Summary.SummaryGroups = <>
                  OptionsCustomize.ColumnFiltering = False
                  OptionsCustomize.ColumnSorting = False
                  OptionsData.Deleting = False
                  OptionsData.Editing = False
                  OptionsData.Inserting = False
                  OptionsSelection.CellSelect = False
                  OptionsView.CellEndEllipsis = True
                  OptionsView.GroupByBox = False
                  OptionsView.HeaderEndEllipsis = True
                  OptionsView.Indicator = True
                  object ColumnKeyId: TcxGridDBColumn
                    DataBinding.FieldName = 'KeyId'
                    Visible = False
                  end
                  object ColumnActDate: TcxGridDBColumn
                    Caption = #26085#26399
                    DataBinding.FieldName = 'ActDate'
                    HeaderAlignmentHorz = taCenter
                    Width = 92
                  end
                  object ColumnActSource: TcxGridDBColumn
                    Caption = #39006#22411
                    DataBinding.FieldName = 'ActSource'
                    PropertiesClassName = 'TcxImageComboBoxProperties'
                    Properties.Images = StyleModule.ActionImageList
                    Properties.Items = <
                      item
                        Description = 'Wrapper'
                        ImageIndex = 6
                        Value = 'Wrapper'
                      end
                      item
                        Description = 'Ppv'
                        ImageIndex = 7
                        Value = 'Ppv'
                      end
                      item
                        Description = 'Sub'
                        ImageIndex = 7
                        Value = 'Sub'
                      end
                      item
                        Description = 'AsRun'
                        ImageIndex = 8
                        Value = 'AsRun'
                      end
                      item
                        Description = 'Merge'
                        ImageIndex = 14
                        Value = 'Merge'
                      end>
                    Width = 90
                  end
                  object ColumnActFileName: TcxGridDBColumn
                    Caption = #27284#26696#21517#31281
                    DataBinding.FieldName = 'ActFileName'
                    Width = 220
                  end
                  object ColumnActFileSize: TcxGridDBColumn
                    Caption = #27284#26696#22823#23567
                    DataBinding.FieldName = 'ActFileSize'
                    HeaderAlignmentHorz = taRightJustify
                    Width = 90
                  end
                  object ColumnActFilePath: TcxGridDBColumn
                    Caption = #36335#24465
                    DataBinding.FieldName = 'ActFilePath'
                    Visible = False
                    Width = 230
                  end
                  object ColumnActProgress: TcxGridDBColumn
                    Caption = #36914#24230
                    DataBinding.FieldName = 'ActProgress'
                    Visible = False
                    Width = 125
                  end
                  object ColumnActTime: TcxGridDBColumn
                    Caption = #34389#29702#26178#38291
                    DataBinding.FieldName = 'ActTime'
                    HeaderAlignmentHorz = taCenter
                    Width = 90
                  end
                  object ColumnActCost: TcxGridDBColumn
                    Caption = #33457#36027#26178#38291'('#31186')'
                    DataBinding.FieldName = 'ActCost'
                    OnGetDisplayText = ColumnActCostGetDisplayText
                    HeaderAlignmentHorz = taRightJustify
                    Width = 93
                  end
                  object ColumnActStatus: TcxGridDBColumn
                    Caption = #34389#29702#29376#24907
                    DataBinding.FieldName = 'ActStatus'
                    PropertiesClassName = 'TcxImageComboBoxProperties'
                    Properties.DefaultImageIndex = 0
                    Properties.Images = StyleModule.ActionImageList
                    Properties.Items = <
                      item
                        Description = #26410#30693
                        ImageIndex = 0
                        Value = ''
                      end
                      item
                        Description = #34389#29702#20013
                        ImageIndex = 1
                        Value = 'P'
                      end
                      item
                        Description = #23436#25104
                        ImageIndex = 2
                        Value = 'C'
                      end
                      item
                        Description = #37679#35492
                        ImageIndex = 3
                        Value = 'E'
                      end>
                    HeaderAlignmentHorz = taCenter
                    Width = 88
                  end
                  object ColumnActErrMsg: TcxGridDBColumn
                    Caption = #37679#35492#35338#24687
                    DataBinding.FieldName = 'ActErrMsg'
                    Width = 550
                  end
                end
                object ActionGridLevel1: TcxGridLevel
                  GridView = gvAction
                end
              end
            end
            object cxTabList: TcxTabSheet
              Caption = 'cxTabList'
              ImageIndex = 1
            end
          end
        end
      end
      object dxDockMsg: TdxDockPanel
        Left = 0
        Top = 362
        Width = 635
        Height = 138
        AllowFloating = False
        AutoHide = False
        Caption = #35338#24687#35222#31383
        CaptionButtons = [cbMaximize, cbHide]
        DockType = 5
        OriginalWidth = 516
        OriginalHeight = 138
        object MsgList: TcxListView
          Left = 0
          Top = 0
          Width = 631
          Height = 113
          Align = alClient
          ColumnClick = False
          Columns = <
            item
              AutoSize = True
              Caption = #35338#24687
            end>
          HideSelection = False
          ParentFont = False
          ReadOnly = True
          RowSelect = True
          ShowColumnHeaders = False
          SmallImages = StyleModule.MsgImageList
          Style.StyleController = StyleModule.MsgStyle
          TabOrder = 0
          ViewStyle = vsReport
        end
      end
    end
    object dxDockNavBar: TdxDockPanel
      Left = 0
      Top = 0
      Width = 120
      Height = 500
      AllowFloating = False
      AutoHide = False
      Caption = #36984#38917#35222#31383
      CaptionButtons = [cbMaximize, cbHide]
      DockType = 2
      OriginalWidth = 120
      OriginalHeight = 140
      object RzPanelNavBar: TRzPanel
        Left = 0
        Top = 0
        Width = 116
        Height = 475
        Align = alClient
        BorderOuter = fsFlat
        FlatColor = clGray
        TabOrder = 0
        object dxNavBar: TdxNavBar
          Left = 1
          Top = 1
          Width = 114
          Height = 473
          Align = alClient
          ActiveGroupIndex = 0
          TabOrder = 0
          View = 3
          OptionsImage.LargeImages = StyleModule.NavBarImageList
          OptionsStyle.DefaultStyles.GroupBackground.BackColor = 13160661
          OptionsStyle.DefaultStyles.GroupBackground.BackColor2 = 8551805
          OptionsStyle.DefaultStyles.GroupBackground.GradientMode = gmForwardDiagonal
          OptionsStyle.DefaultStyles.GroupBackground.Font.Charset = DEFAULT_CHARSET
          OptionsStyle.DefaultStyles.GroupBackground.Font.Color = clWindowText
          OptionsStyle.DefaultStyles.GroupBackground.Font.Height = -11
          OptionsStyle.DefaultStyles.GroupBackground.Font.Name = 'MS Sans Serif'
          OptionsStyle.DefaultStyles.GroupBackground.Font.Style = []
          OptionsStyle.DefaultStyles.GroupBackground.HAlignment = haLeft
          OptionsStyle.DefaultStyles.Item.BackColor = clWhite
          OptionsStyle.DefaultStyles.Item.BackColor2 = clWhite
          OptionsStyle.DefaultStyles.Item.Font.Charset = ANSI_CHARSET
          OptionsStyle.DefaultStyles.Item.Font.Color = clBtnText
          OptionsStyle.DefaultStyles.Item.Font.Height = -12
          OptionsStyle.DefaultStyles.Item.Font.Name = 'Tahoma'
          OptionsStyle.DefaultStyles.Item.Font.Style = []
          OptionsStyle.DefaultStyles.ItemDisabled.BackColor = clWhite
          OptionsStyle.DefaultStyles.ItemDisabled.BackColor2 = clWhite
          OptionsStyle.DefaultStyles.ItemDisabled.Font.Charset = ANSI_CHARSET
          OptionsStyle.DefaultStyles.ItemDisabled.Font.Color = clBtnShadow
          OptionsStyle.DefaultStyles.ItemDisabled.Font.Height = -12
          OptionsStyle.DefaultStyles.ItemDisabled.Font.Name = 'Tahoma'
          OptionsStyle.DefaultStyles.ItemDisabled.Font.Style = []
          OptionsStyle.DefaultStyles.ItemHotTracked.BackColor = clWhite
          OptionsStyle.DefaultStyles.ItemHotTracked.BackColor2 = clWhite
          OptionsStyle.DefaultStyles.ItemHotTracked.Font.Charset = ANSI_CHARSET
          OptionsStyle.DefaultStyles.ItemHotTracked.Font.Color = clBtnText
          OptionsStyle.DefaultStyles.ItemHotTracked.Font.Height = -12
          OptionsStyle.DefaultStyles.ItemHotTracked.Font.Name = 'Tahoma'
          OptionsStyle.DefaultStyles.ItemHotTracked.Font.Style = []
          OptionsStyle.DefaultStyles.ItemPressed.BackColor = clWhite
          OptionsStyle.DefaultStyles.ItemPressed.BackColor2 = clWhite
          OptionsStyle.DefaultStyles.ItemPressed.Font.Charset = ANSI_CHARSET
          OptionsStyle.DefaultStyles.ItemPressed.Font.Color = clBtnText
          OptionsStyle.DefaultStyles.ItemPressed.Font.Height = -12
          OptionsStyle.DefaultStyles.ItemPressed.Font.Name = 'Tahoma'
          OptionsStyle.DefaultStyles.ItemPressed.Font.Style = []
          OptionsView.Common.ShowGroupCaptions = False
          object dxNavBarGroup1: TdxNavBarGroup
            Caption = 'dxNavBarGroup1'
            LinksUseSmallImages = False
            SelectedLinkIndex = -1
            TopVisibleLinkIndex = 0
            Links = <
              item
                Item = dxAction
              end
              item
                Item = dxList
              end
              item
                Item = dxOption
              end>
          end
          object dxAction: TdxNavBarItem
            Caption = #21295#25972#27511#31243
            LargeImageIndex = 0
            OnClick = dxActionClick
          end
          object dxList: TdxNavBarItem
            Caption = #31680#30446#34920
            LargeImageIndex = 1
            Visible = False
            OnClick = dxListClick
          end
          object dxOption: TdxNavBarItem
            Caption = #36984#38917#35373#23450
            LargeImageIndex = 2
            OnClick = dxOptionClick
          end
        end
      end
    end
  end
  object dxBarManager: TdxBarManager
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = #24494#36575#27491#40657#39636
    Font.Style = []
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
    ImageOptions.HotImages = StyleModule.BarImageList
    ImageOptions.Images = StyleModule.BarImageList
    PopupMenuLinks = <>
    Style = bmsOffice11
    UseSystemFont = True
    Left = 20
    Top = 292
    DockControlHeights = (
      0
      0
      52
      24)
    object dxBarManagerBar1: TdxBar
      AllowClose = False
      AllowCustomizing = False
      AllowQuickCustomizing = False
      AllowReset = False
      BorderStyle = bbsNone
      Caption = #29376#24907#21015
      CaptionButtons = <>
      DockedDockingStyle = dsBottom
      DockedLeft = 0
      DockedTop = 0
      DockingStyle = dsBottom
      FloatLeft = 788
      FloatTop = 369
      FloatClientWidth = 70
      FloatClientHeight = 70
      ItemLinks = <>
      NotDocking = [dsNone]
      OldName = 'StatusBar'
      OneOnRow = True
      Row = 0
      UseOwnFont = False
      Visible = True
      WholeRow = True
    end
    object dxBarManagerBar2: TdxBar
      AllowClose = False
      AllowCustomizing = False
      AllowQuickCustomizing = False
      AllowReset = False
      Caption = #21151#33021#34920
      CaptionButtons = <>
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
          Visible = True
          ItemName = 'dxAbout'
        end>
      MultiLine = True
      NotDocking = [dsNone, dsLeft, dsRight, dsBottom]
      OldName = 'MenuBar'
      OneOnRow = True
      Row = 0
      UseOwnFont = False
      Visible = True
      WholeRow = True
    end
    object dxBarManagerBar3: TdxBar
      AllowClose = False
      AllowCustomizing = False
      AllowQuickCustomizing = False
      AllowReset = False
      Caption = #24037#20855#21015
      CaptionButtons = <>
      DockedDockingStyle = dsTop
      DockedLeft = 0
      DockedTop = 26
      DockingStyle = dsTop
      FloatLeft = 464
      FloatTop = 357
      FloatClientWidth = 23
      FloatClientHeight = 22
      ItemLinks = <
        item
          Visible = True
          ItemName = 'dxPlay'
        end
        item
          Visible = True
          ItemName = 'dxStop'
        end
        item
          BeginGroup = True
          Visible = True
          ItemName = 'dxAutoScroll'
        end>
      NotDocking = [dsLeft, dsRight, dsBottom]
      OldName = 'ToolBar'
      OneOnRow = True
      Row = 1
      UseOwnFont = False
      Visible = True
      WholeRow = True
    end
    object dxAbout: TdxBarButton
      Caption = #38364#26044
      Category = 0
      Hint = #38364#26044
      Visible = ivAlways
    end
    object dxPlay: TdxBarButton
      Action = actPlay
      Category = 2
    end
    object dxStop: TdxBarButton
      Action = actStop
      Category = 2
    end
    object dxAutoScroll: TdxBarButton
      Action = actAutoScroll
      Category = 2
      ButtonStyle = bsChecked
    end
    object PlayCmdGroup: TdxBarGroup
      Items = (
        'dxPlay')
    end
    object StopCmdGroup: TdxBarGroup
      Enabled = False
      Items = (
        'dxStop')
    end
    object ControlSendGroup: TdxBarGroup
      Items = ()
    end
  end
  object dxDockingManager: TdxDockingManager
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
    Options = [doActivateAfterDocking, doDblClickDocking, doFloatingOnTop, doTabContainerHasCaption, doTabContainerCanAutoHide, doSideContainerCanClose, doSideContainerCanAutoHide, doTabContainerCanInSideContainer]
    ViewStyle = vsOffice11
    Left = 20
    Top = 328
  end
  object StartupTimer: TTimer
    Enabled = False
    Interval = 500
    OnTimer = StartupTimerTimer
    Left = 20
    Top = 365
  end
  object BarActionList: TActionList
    Images = StyleModule.BarImageList
    Left = 20
    Top = 404
    object actPlay: TAction
      Caption = #21855#21205
      ImageIndex = 0
      OnExecute = actPlayExecute
    end
    object actStop: TAction
      Caption = #20572#27490
      ImageIndex = 2
      OnExecute = actStopExecute
    end
    object actAutoScroll: TAction
      Caption = 'actAutoScroll'
      ImageIndex = 11
      OnExecute = actAutoScrollExecute
    end
  end
  object dsAction: TDataSource
    DataSet = LogModule.ActionDataSet
    Left = 520
    Top = 329
  end
end
