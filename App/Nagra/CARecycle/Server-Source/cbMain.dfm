object fmMain: TfmMain
  Left = 316
  Top = 180
  Width = 632
  Height = 479
  Caption = #38283#21338#31185#25216'-CA Recycle Processer'
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object dxDockHost: TdxDockSite
    Left = 0
    Top = 52
    Width = 624
    Height = 376
    Align = alClient
    DockType = 0
    OriginalWidth = 624
    OriginalHeight = 376
    object dxLayoutDockSite2: TdxLayoutDockSite
      Left = 194
      Top = 0
      Width = 430
      Height = 376
      DockType = 1
      OriginalWidth = 300
      OriginalHeight = 200
      object dxLayoutDockSite3: TdxLayoutDockSite
        Left = 0
        Top = 0
        Width = 430
        Height = 226
        DockType = 1
        OriginalWidth = 300
        OriginalHeight = 200
        object dxLayoutDockSite1: TdxLayoutDockSite
          Left = 0
          Top = 0
          Width = 430
          Height = 226
          DockType = 1
          OriginalWidth = 300
          OriginalHeight = 200
        end
        object dxDockPanel2: TdxDockPanel
          Left = 0
          Top = 0
          Width = 430
          Height = 226
          AllowFloating = True
          AutoHide = False
          Caption = #34389#29702#29376#24907
          CaptionButtons = []
          ImageIndex = 2
          DockType = 1
          OriginalWidth = 185
          OriginalHeight = 124
          object RecycleGrid: TcxGrid
            Left = 0
            Top = 0
            Width = 426
            Height = 202
            Align = alClient
            TabOrder = 0
            LookAndFeel.Kind = lfOffice11
            LookAndFeel.NativeStyle = False
            object gvRecycle: TcxGridDBBandedTableView
              NavigatorButtons.ConfirmDelete = False
              OnCustomDrawCell = gvRecycleCustomDrawCell
              DataController.DataModeController.SyncMode = False
              DataController.DataSource = dsRecycle
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
              OptionsView.RowSeparatorWidth = 3
              Bands = <
                item
                  Width = 82
                end
                item
                  Width = 198
                end
                item
                  Width = 119
                end
                item
                  Caption = #26410#21855#29992#22238#25765
                  Width = 195
                end
                item
                  Caption = #21855#29992#22238#25765
                  Width = 326
                end
                item
                end>
              object gvRecycleColumn1: TcxGridDBBandedColumn
                Caption = #31995#32113#21488
                DataBinding.FieldName = 'COMPNAME'
                PropertiesClassName = 'TcxTextEditProperties'
                Properties.Alignment.Horz = taCenter
                Properties.Alignment.Vert = taVCenter
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 143
                Position.BandIndex = 0
                Position.ColIndex = 0
                Position.LineCount = 2
                Position.RowIndex = 0
              end
              object gvRecycleColumn2: TcxGridDBBandedColumn
                Caption = 'ICC '#21345#34399
                DataBinding.FieldName = 'ICCNO'
                Width = 104
                Position.BandIndex = 1
                Position.ColIndex = 0
                Position.RowIndex = 0
              end
              object gvRecycleColumn3: TcxGridDBBandedColumn
                Caption = 'STB '#24207#34399
                DataBinding.FieldName = 'STBNO'
                Width = 104
                Position.BandIndex = 1
                Position.ColIndex = 0
                Position.RowIndex = 1
              end
              object gvRecycleColumn4: TcxGridDBBandedColumn
                Caption = #36039#26009#36664#20837#26085
                DataBinding.FieldName = 'RCSTARTDATE'
                Width = 94
                Position.BandIndex = 1
                Position.ColIndex = 1
                Position.RowIndex = 0
              end
              object gvRecycleColumn5: TcxGridDBBandedColumn
                Caption = #20316#26989#20154#21729
                DataBinding.FieldName = 'OPERATOR'
                Width = 94
                Position.BandIndex = 1
                Position.ColIndex = 1
                Position.RowIndex = 1
              end
              object gvRecycleColumn6: TcxGridDBBandedColumn
                Caption = #40670#25976#28165#38500#26085
                DataBinding.FieldName = 'LASTDOWNLOADTIME'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Position.BandIndex = 2
                Position.ColIndex = 0
                Position.RowIndex = 1
              end
              object gvRecycleColumn7: TcxGridDBBandedColumn
                Tag = 1
                Caption = #37197#23565
                DataBinding.FieldName = 'CMD52_1'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 60
                Position.BandIndex = 3
                Position.ColIndex = 0
                Position.RowIndex = 0
              end
              object gvRecycleColumn8: TcxGridDBBandedColumn
                Tag = 1
                Caption = #40670#25976
                DataBinding.FieldName = 'CMD8'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 60
                Position.BandIndex = 3
                Position.ColIndex = 0
                Position.RowIndex = 1
              end
              object gvRecycleColumn9: TcxGridDBBandedColumn
                Tag = 1
                Caption = #25910#35222#35352#37636
                DataBinding.FieldName = 'CMD97_96'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 60
                Position.BandIndex = 3
                Position.ColIndex = 1
                Position.RowIndex = 0
              end
              object gvRecycleColumn10: TcxGridDBBandedColumn
                Tag = 1
                Caption = #37109#36958#21312#34399
                DataBinding.FieldName = 'CMD48'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 60
                Position.BandIndex = 3
                Position.ColIndex = 2
                Position.RowIndex = 0
              end
              object gvRecycleColumn11: TcxGridDBBandedColumn
                Tag = 1
                Caption = #38364#38971#36947
                DataBinding.FieldName = 'CMD7'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 60
                Position.BandIndex = 3
                Position.ColIndex = 1
                Position.RowIndex = 1
              end
              object gvRecycleColumn12: TcxGridDBBandedColumn
                Tag = 1
                Caption = #35299#37197#23565
                DataBinding.FieldName = 'CMD52_2'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 60
                Position.BandIndex = 3
                Position.ColIndex = 2
                Position.RowIndex = 1
              end
              object gvRecycleColumn13: TcxGridDBBandedColumn
                Caption = #26159#21542#22238#25765
                DataBinding.FieldName = 'STBAUTOCB'
                PropertiesClassName = 'TcxCheckBoxProperties'
                Properties.NullStyle = nssUnchecked
                Properties.ValueChecked = '1'
                Properties.ValueUnchecked = '0'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Position.BandIndex = 2
                Position.ColIndex = 0
                Position.RowIndex = 0
              end
              object gvRecycleColumn14: TcxGridDBBandedColumn
                Tag = 1
                Caption = #37197#23565
                DataBinding.FieldName = 'CallCMD52_1'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 87
                Position.BandIndex = 4
                Position.ColIndex = 0
                Position.RowIndex = 0
              end
              object gvRecycleColumn15: TcxGridDBBandedColumn
                Tag = 1
                Caption = #33258#21205#22238#25765'('#38364')'
                DataBinding.FieldName = 'CallCMD62'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 88
                Position.BandIndex = 4
                Position.ColIndex = 0
                Position.RowIndex = 1
              end
              object gvRecycleColumn16: TcxGridDBBandedColumn
                Tag = 1
                Caption = #28165#38500#40670#25976
                DataBinding.FieldName = 'CallCMD8'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 78
                Position.BandIndex = 4
                Position.ColIndex = 1
                Position.RowIndex = 0
              end
              object gvRecycleColumn17: TcxGridDBBandedColumn
                Tag = 1
                Caption = #31435#21363#22238#25765
                DataBinding.FieldName = 'CallCMD60'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 79
                Position.BandIndex = 4
                Position.ColIndex = 1
                Position.RowIndex = 1
              end
              object gvRecycleColumn18: TcxGridDBBandedColumn
                Tag = 1
                Caption = #38364#38971#36947
                DataBinding.FieldName = 'CallCMD7'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 82
                Position.BandIndex = 4
                Position.ColIndex = 2
                Position.RowIndex = 0
              end
              object gvRecycleColumn19: TcxGridDBBandedColumn
                Tag = 1
                Caption = #37109#36958#21312#34399
                DataBinding.FieldName = 'CallCMD48'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 82
                Position.BandIndex = 4
                Position.ColIndex = 2
                Position.RowIndex = 1
              end
              object gvRecycleColumn20: TcxGridDBBandedColumn
                Tag = 1
                Caption = #31561#20505#22238#20659
                DataBinding.FieldName = 'CallCMD99'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 79
                Position.BandIndex = 4
                Position.ColIndex = 3
                Position.RowIndex = 0
              end
              object gvRecycleColumn21: TcxGridDBBandedColumn
                Tag = 1
                Caption = #35299#37197#23565
                DataBinding.FieldName = 'CallCMD52_2'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 80
                Position.BandIndex = 4
                Position.ColIndex = 3
                Position.RowIndex = 1
              end
              object gvRecycleColumn22: TcxGridDBBandedColumn
                Caption = #37679#35492#35338#24687
                DataBinding.FieldName = 'ERRMSG'
                PropertiesClassName = 'TcxTextEditProperties'
                Properties.Alignment.Vert = taVCenter
                HeaderAlignmentVert = vaCenter
                Width = 280
                Position.BandIndex = 5
                Position.ColIndex = 0
                Position.LineCount = 2
                Position.RowIndex = 0
              end
            end
            object glRecycle: TcxGridLevel
              GridView = gvRecycle
            end
          end
        end
      end
      object dxDockPanel5: TdxDockPanel
        Left = 0
        Top = 226
        Width = 430
        Height = 150
        AllowFloating = True
        AutoHide = False
        Caption = #35338#24687#35222#31383
        CaptionButtons = []
        ImageIndex = 1
        DockType = 5
        OriginalWidth = 185
        OriginalHeight = 150
        object MsgList: TcxListView
          Left = 0
          Top = 0
          Width = 426
          Height = 126
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
          Style.Font.Charset = DEFAULT_CHARSET
          Style.Font.Color = clWindowText
          Style.Font.Height = -12
          Style.Font.Name = 'Verdana'
          Style.Font.Style = []
          Style.StyleController = StyleModule.MsgStyle
          Style.IsFontAssigned = True
          TabOrder = 0
          ViewStyle = vsReport
        end
      end
    end
    object dxDockPanel1: TdxDockPanel
      Left = 0
      Top = 0
      Width = 194
      Height = 376
      AllowFloating = True
      AutoHide = False
      Caption = #31995#32113#21488
      ImageIndex = 0
      DockType = 2
      OriginalWidth = 194
      OriginalHeight = 140
      object SoTree: TcxTreeList
        Left = 0
        Top = 0
        Width = 190
        Height = 352
        Align = alClient
        Bands = <
          item
          end>
        BufferedPaint = True
        Images = StyleModule.SoTreeImageList
        LookAndFeel.Kind = lfOffice11
        OptionsBehavior.Sorting = False
        OptionsData.Editing = False
        OptionsData.Deleting = False
        OptionsSelection.CellSelect = False
        OptionsView.CellEndEllipsis = True
        OptionsView.CategorizedColumn = SoTreeColumn2
        OptionsView.PaintStyle = tlpsCategorized
        StateImages = StyleModule.SoTreeImageList
        TabOrder = 0
        object SoTreeColumn1: TcxTreeListColumn
          Caption.Text = #20195#30908
          DataBinding.ValueType = 'String'
          Width = 107
          Position.ColIndex = 0
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
        object SoTreeColumn2: TcxTreeListColumn
          Caption.Text = #31995#32113#21488
          DataBinding.ValueType = 'String'
          Width = 73
          Position.ColIndex = 1
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
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
        ItemLinks = <>
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
            Item = dxMenuItem1
            Visible = True
          end
          item
            Item = dxMenuItem2
            Visible = True
          end
          item
            Item = dxMenuItem3
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
        DockedTop = 26
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
            Item = dxStop
            Visible = True
          end
          item
            Item = dxAutoScroll
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
    Images = StyleModule.BarImageList
    PopupMenuLinks = <>
    Style = bmsOffice11
    UseSystemFont = True
    Left = 60
    Top = 196
    DockControlHeights = (
      0
      0
      52
      24)
    object dxMenuItem1: TdxBarSubItem
      Caption = #31995#32113
      Category = 0
      Visible = ivAlways
      ItemLinks = <
        item
          Item = dxMenuItem11
          Visible = True
        end>
    end
    object dxMenuItem2: TdxBarSubItem
      Caption = #27298#35222
      Category = 0
      Visible = ivAlways
      ItemLinks = <
        item
          Item = dxMenuItem21
          Visible = True
        end
        item
          BeginGroup = True
          Item = dxMenuItem22
          Visible = True
        end>
    end
    object dxPlay: TdxBarButton
      Caption = #21855#21205
      Category = 2
      Hint = #38283#22987#22519#34892#25351#20196#20659#36865#21450#25509#25910
      Visible = ivAlways
      ImageIndex = 0
      OnClick = dxPlayClick
    end
    object dxStop: TdxBarButton
      Caption = #20572#27490
      Category = 2
      Enabled = False
      Hint = #20572#27490#22519#34892#25351#20196#20659#36865#21450#25509#25910
      Visible = ivAlways
      ImageIndex = 2
      OnClick = dxStopClick
    end
    object dxAutoScroll: TdxBarButton
      Caption = #25351#20196#25458#21205
      Category = 2
      Hint = #33258#21205#25458#21205#25351#20196
      Visible = ivAlways
      ButtonStyle = bsChecked
      ImageIndex = 11
    end
    object dxMenuItem11: TdxBarButton
      Caption = #36984#38917
      Category = 0
      Hint = #36984#38917
      Visible = ivAlways
      ImageIndex = 7
      OnClick = dxMenuItem11Click
    end
    object dxMenuItem3: TdxBarButton
      Caption = #38364#26044
      Category = 0
      Hint = #38364#26044
      Visible = ivAlways
    end
    object dxMenuItem22: TdxBarButton
      Caption = #31995#32113#21488#29376#24907
      Category = 0
      Hint = #31995#32113#21488#29376#24907
      Visible = ivAlways
    end
    object dxMenuItem21: TdxBarButton
      Caption = #22238#24489#33267#38928#35373#20301#32622
      Category = 0
      Hint = #22238#24489#33267#38928#35373#20301#32622
      Visible = ivAlways
    end
    object PlayCmdGroup: TdxBarGroup
      Items = (
        'dxPlay'
        'dxMenuItem11')
    end
    object StopCmdGroup: TdxBarGroup
      Enabled = False
      Items = (
        'dxStop')
    end
    object ControlSendGroup: TdxBarGroup
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
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    Images = StyleModule.DockPanelImageList
    Options = [doActivateAfterDocking, doDblClickDocking, doFloatingOnTop, doTabContainerHasCaption, doTabContainerCanAutoHide, doSideContainerCanClose, doSideContainerCanAutoHide, doTabContainerCanInSideContainer]
    ViewStyle = vsOffice11
    Left = 24
    Top = 196
  end
  object LoadConfigTimer: TTimer
    Enabled = False
    OnTimer = LoadConfigTimerTimer
    Left = 100
    Top = 196
  end
  object cdsRecycle: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'SEQNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ICCNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'STBNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'STB_FLAG'
        DataType = ftInteger
      end
      item
        Name = 'STBAUTOCB'
        DataType = ftInteger
      end
      item
        Name = 'COMPCODE'
        DataType = ftInteger
      end
      item
        Name = 'COMPNAME'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'RCSTARTDATE'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'RCENDDATE'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'OPERATOR'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'UPDTIME'
        DataType = ftDateTime
      end
      item
        Name = 'LASTDOWNLOADTIME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CMDFLAG'
        DataType = ftInteger
      end
      item
        Name = 'RECYCLETEXT'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'TRANSNUM'
        DataType = ftString
        Size = 512
      end
      item
        Name = 'CONFIRM'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ERRMSG'
        DataType = ftString
        Size = 512
      end
      item
        Name = 'CMD52_1'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CMD8'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CMD97_96'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CMD7'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CMD48'
        DataType = ftString
        Size = 2
      end
      item
        Name = 'CMD52_2'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CallCMD52_1'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CallCMD62'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CallCMD8'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CallCMD60'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CallCMD7'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CallCMD48'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CallCMD99'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CallCMD52_2'
        DataType = ftString
        Size = 1
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 436
    Top = 212
  end
  object dsRecycle: TDataSource
    DataSet = cdsRecycle
    Left = 402
    Top = 212
  end
end
