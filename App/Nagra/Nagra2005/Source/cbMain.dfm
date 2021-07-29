object fmMain: TfmMain
  Left = 287
  Top = 166
  Width = 677
  Height = 551
  Caption = #38283#21338#31185#25216' - SMS Command Gateway'
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Verdana'
  Font.Style = []
  OldCreateOrder = False
  Scaled = False
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object dxDockHost: TdxDockSite
    Left = 0
    Top = 86
    Width = 669
    Height = 405
    Align = alClient
    DockType = 0
    OriginalWidth = 669
    OriginalHeight = 405
    object dxLayoutDockSite2: TdxLayoutDockSite
      Left = 260
      Top = 0
      Width = 409
      Height = 405
      DockType = 1
      OriginalWidth = 300
      OriginalHeight = 200
      object dxLayoutDockSite3: TdxLayoutDockSite
        Left = 0
        Top = 0
        Width = 409
        Height = 255
        DockType = 1
        OriginalWidth = 300
        OriginalHeight = 200
        object dxLayoutDockSite1: TdxLayoutDockSite
          Left = 0
          Top = 0
          Width = 409
          Height = 255
          DockType = 1
          OriginalWidth = 300
          OriginalHeight = 200
        end
        object dxTabContainerDockSite3: TdxTabContainerDockSite
          Left = 0
          Top = 0
          Width = 409
          Height = 255
          ActiveChildIndex = 0
          AllowFloating = True
          AutoHide = False
          CaptionButtons = []
          DockType = 1
          OriginalWidth = 185
          OriginalHeight = 124
          object dxDockPanel2: TdxDockPanel
            Left = 0
            Top = 0
            Width = 405
            Height = 207
            AllowFloating = True
            AutoHide = False
            Caption = 'dxDockPanel2'
            CaptionButtons = []
            ImageIndex = 2
            OnActivate = dxDockPanel2Activate
            OnResize = dxDockPanelResize
            DockType = 1
            OriginalWidth = 175
            OriginalHeight = 88
            object ControlSendTree: TcxTreeList
              Left = 0
              Top = 0
              Width = 405
              Height = 207
              Align = alClient
              Bands = <
                item
                end>
              BufferedPaint = True
              Images = StyleModule.CmdTreeImageList
              LookAndFeel.Kind = lfOffice11
              OptionsBehavior.Sorting = False
              OptionsSelection.CellSelect = False
              OptionsView.CellEndEllipsis = True
              OptionsView.GridLines = tlglBoth
              TabOrder = 0
              OnCustomDrawCell = ControlSendTreeCustomDrawCell
              object ControlSendTreeColumn1: TcxTreeListColumn
                Caption.Text = 'COMPNAME'
                DataBinding.ValueType = 'String'
                Width = 110
                Position.ColIndex = 0
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object ControlSendTreeColumn2: TcxTreeListColumn
                Caption.Text = 'SEQNO'
                DataBinding.ValueType = 'String'
                Width = 70
                Position.ColIndex = 1
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object ControlSendTreeColumn3: TcxTreeListColumn
                Caption.Text = 'ICCNO'
                DataBinding.ValueType = 'String'
                Width = 85
                Position.ColIndex = 2
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object ControlSendTreeColumn4: TcxTreeListColumn
                Caption.Text = 'CMDID'
                DataBinding.ValueType = 'String'
                Width = 230
                Position.ColIndex = 3
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object ControlSendTreeColumn5: TcxTreeListColumn
                PropertiesClassName = 'TcxImageComboBoxProperties'
                Properties.DefaultImageIndex = 4
                Properties.Images = StyleModule.CmdTreeImageList
                Properties.Items = <
                  item
                    ImageIndex = 5
                    Value = 'C'
                  end
                  item
                    ImageIndex = 6
                    Value = 'E'
                  end
                  item
                    ImageIndex = 5
                    Value = 'P'
                  end>
                Caption.Text = 'SEND'
                DataBinding.ValueType = 'String'
                Width = 40
                Position.ColIndex = 4
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object ControlSendTreeColumn6: TcxTreeListColumn
                PropertiesClassName = 'TcxImageComboBoxProperties'
                Properties.DefaultImageIndex = 4
                Properties.Images = StyleModule.CmdTreeImageList
                Properties.Items = <
                  item
                    ImageIndex = 5
                    Value = 'C'
                  end
                  item
                    ImageIndex = 6
                    Value = 'E'
                  end>
                Caption.Text = 'RECV'
                DataBinding.ValueType = 'String'
                Width = 40
                Position.ColIndex = 5
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object ControlSendTreeColumn7: TcxTreeListColumn
                Caption.Text = 'OPERATOR'
                DataBinding.ValueType = 'String'
                Position.ColIndex = 6
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object ControlSendTreeColumn8: TcxTreeListColumn
                Caption.Text = 'ERRORCODE'
                DataBinding.ValueType = 'String'
                Width = 150
                Position.ColIndex = 7
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object ControlSendTreeColumn9: TcxTreeListColumn
                Caption.Text = 'ERRORMSG'
                DataBinding.ValueType = 'String'
                Width = 150
                Position.ColIndex = 8
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object ControlSendTreeColumn10: TcxTreeListColumn
                Caption.Text = 'UPDTIME'
                DataBinding.ValueType = 'String'
                Width = 130
                Position.ColIndex = 9
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
            end
          end
          object dxDockPanel3: TdxDockPanel
            Left = 0
            Top = 0
            Width = 405
            Height = 207
            AllowFloating = True
            AutoHide = False
            Caption = 'dxDockPanel3'
            CaptionButtons = []
            ImageIndex = 3
            DockType = 1
            OriginalWidth = 175
            OriginalHeight = 88
            object ConsoleTree: TcxTreeList
              Left = 0
              Top = 0
              Width = 405
              Height = 207
              Align = alClient
              Bands = <
                item
                end>
              BufferedPaint = True
              Images = StyleModule.CmdTreeImageList
              LookAndFeel.Kind = lfOffice11
              OptionsBehavior.Sorting = False
              OptionsData.Editing = False
              OptionsData.Deleting = False
              OptionsSelection.CellSelect = False
              OptionsSelection.MultiSelect = True
              OptionsView.CellEndEllipsis = True
              OptionsView.TreeLineStyle = tllsNone
              TabOrder = 0
              OnKeyDown = ConsoleTreeKeyDown
              object ControlConsoleTreeColumn1: TcxTreeListColumn
                Caption.Text = #20659#36865#24207#34399
                DataBinding.ValueType = 'String'
                Position.ColIndex = 2
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object ControlConsoleTreeColumn2: TcxTreeListColumn
                Caption.Text = #22238#25033#24207#34399
                DataBinding.ValueType = 'String'
                Position.ColIndex = 3
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object ControlConsoleTreeColumn3: TcxTreeListColumn
                Caption.Text = #26684#24335
                DataBinding.ValueType = 'String'
                Width = 1500
                Position.ColIndex = 4
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object ControlConsoleTreeColumn4: TcxTreeListColumn
                Caption.Text = #22238#25033
                DataBinding.ValueType = 'String'
                Width = 85
                Position.ColIndex = 0
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object ControlConsoleTreeColumn5: TcxTreeListColumn
                Caption.Text = #26178#38291
                DataBinding.ValueType = 'String'
                Width = 180
                Position.ColIndex = 1
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
            end
          end
          object dxDockPanel4: TdxDockPanel
            Left = 0
            Top = 0
            Width = 405
            Height = 207
            AllowFloating = True
            AutoHide = False
            Caption = 'dxDockPanel4'
            ImageIndex = 4
            OnResize = dxDockPanelResize
            DockType = 1
            OriginalWidth = 260
            OriginalHeight = 309
            object FeedbackTree: TcxTreeList
              Left = 0
              Top = 0
              Width = 405
              Height = 207
              Align = alClient
              Bands = <
                item
                end>
              BufferedPaint = True
              Images = StyleModule.CmdTreeImageList
              LookAndFeel.Kind = lfOffice11
              OptionsBehavior.Sorting = False
              OptionsData.Editing = False
              OptionsData.Deleting = False
              OptionsSelection.CellSelect = False
              OptionsView.GridLines = tlglBoth
              OptionsView.ShowRoot = False
              OptionsView.TreeLineStyle = tllsNone
              TabOrder = 0
              object FeedbackTreecColumn1: TcxTreeListColumn
                Caption.Text = 'SEQNO'
                DataBinding.ValueType = 'String'
                Width = 110
                Position.ColIndex = 0
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object FeedbackTreecColumn2: TcxTreeListColumn
                Caption.Text = 'ICCNO'
                DataBinding.ValueType = 'String'
                Position.ColIndex = 1
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object FeedbackTreecColumn3: TcxTreeListColumn
                Caption.Text = 'CMDID'
                DataBinding.ValueType = 'String'
                Width = 300
                Position.ColIndex = 3
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object FeedbackTreecColumn4: TcxTreeListColumn
                PropertiesClassName = 'TcxImageComboBoxProperties'
                Properties.Images = StyleModule.CmdTreeImageList
                Properties.Items = <
                  item
                    ImageIndex = 4
                    Value = 'W'
                  end
                  item
                    ImageIndex = 5
                    Value = 'C'
                  end
                  item
                    ImageIndex = 6
                    Value = 'E'
                  end>
                Caption.Text = 'WRITESTATUS'
                DataBinding.ValueType = 'String'
                Width = 40
                Position.ColIndex = 4
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object FeedbackTreecColumn5: TcxTreeListColumn
                PropertiesClassName = 'TcxImageComboBoxProperties'
                Properties.Images = StyleModule.CmdTreeImageList
                Properties.Items = <
                  item
                    ImageIndex = 4
                    Value = 'W'
                  end
                  item
                    ImageIndex = 5
                    Value = 'C'
                  end
                  item
                    ImageIndex = 6
                    Value = 'E'
                  end>
                Caption.Text = 'SENDDTATUS'
                DataBinding.ValueType = 'String'
                Width = 40
                Position.ColIndex = 5
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object FeedbackTreecColumn6: TcxTreeListColumn
                Caption.Text = 'UPDTIME'
                DataBinding.ValueType = 'String'
                Width = 150
                Position.ColIndex = 6
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
              object FeedbackTreecColumn7: TcxTreeListColumn
                Caption.Text = 'STBNO'
                DataBinding.ValueType = 'String'
                Position.ColIndex = 2
                Position.RowIndex = 0
                Position.BandIndex = 0
              end
            end
          end
        end
      end
      object dxTabContainerDockSite2: TdxTabContainerDockSite
        Left = 0
        Top = 255
        Width = 409
        Height = 150
        ActiveChildIndex = 0
        AllowFloating = True
        AutoHide = False
        DockType = 5
        OriginalWidth = 185
        OriginalHeight = 150
        object dxDockPanel5: TdxDockPanel
          Left = 0
          Top = 0
          Width = 405
          Height = 102
          AllowFloating = True
          AutoHide = False
          Caption = 'dxDockPanel5'
          ImageIndex = 1
          DockType = 1
          OriginalWidth = 165
          OriginalHeight = 150
          object StartupMsgList: TcxListView
            Left = 0
            Top = 0
            Width = 405
            Height = 102
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
        object dxDockPanel6: TdxDockPanel
          Left = 0
          Top = 0
          Width = 405
          Height = 102
          AllowFloating = True
          AutoHide = False
          Caption = 'dxDockPanel6'
          ImageIndex = 1
          DockType = 1
          OriginalWidth = 165
          OriginalHeight = 150
          object SoDatabaseMsgList: TcxListView
            Left = 0
            Top = 0
            Width = 405
            Height = 102
            Align = alClient
            ColumnClick = False
            Columns = <
              item
                Width = 1500
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
        object dxDockPanel7: TdxDockPanel
          Left = 0
          Top = 0
          Width = 405
          Height = 102
          AllowFloating = True
          AutoHide = False
          Caption = 'dxDockPanel7'
          ImageIndex = 1
          DockType = 1
          OriginalWidth = 165
          OriginalHeight = 150
          object ControlMsgList: TcxListView
            Left = 0
            Top = 0
            Width = 405
            Height = 102
            Align = alClient
            ColumnClick = False
            Columns = <
              item
                Width = 1500
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
        object dxDockPanel8: TdxDockPanel
          Left = 0
          Top = 0
          Width = 405
          Height = 102
          AllowFloating = True
          AutoHide = False
          Caption = 'dxDockPanel8'
          ImageIndex = 1
          DockType = 1
          OriginalWidth = 185
          OriginalHeight = 150
          object FeedbackMsgList: TcxListView
            Left = 0
            Top = 0
            Width = 405
            Height = 102
            Align = alClient
            ColumnClick = False
            Columns = <
              item
                Width = 1500
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
    end
    object dxDockPanel1: TdxDockPanel
      Left = 0
      Top = 0
      Width = 260
      Height = 405
      AllowFloating = True
      AutoHide = False
      Caption = 'dxDockPanel1'
      ImageIndex = 0
      OnResize = dxDockPanelResize
      DockType = 2
      OriginalWidth = 260
      OriginalHeight = 140
      object SoTree: TcxTreeList
        Left = 0
        Top = 0
        Width = 256
        Height = 381
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
        OptionsView.CategorizedColumn = SoTreeCol2
        OptionsView.GridLines = tlglBoth
        OptionsView.PaintStyle = tlpsCategorized
        StateImages = StyleModule.SoTreeImageList
        TabOrder = 0
        object SoTreeCol1: TcxTreeListColumn
          Caption.Text = #20195#30908
          DataBinding.ValueType = 'String'
          Width = 107
          Position.ColIndex = 0
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
        object SoTreeCol2: TcxTreeListColumn
          Caption.Text = #31995#32113#21488
          DataBinding.ValueType = 'String'
          Width = 73
          Position.ColIndex = 1
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
        object SoTreeCol3: TcxTreeListColumn
          Caption.Text = #32047#35336#25351#20196
          DataBinding.ValueType = 'String'
          Width = 70
          Position.ColIndex = 2
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
        object SoTreeCol4: TcxTreeListColumn
          PropertiesClassName = 'TcxProgressBarProperties'
          Properties.AssignedValues.Min = True
          Properties.BeginColor = 54056
          Properties.BorderWidth = 1
          Properties.Max = 100.000000000000000000
          Properties.OverloadBeginColor = clRed
          Properties.ShowOverload = True
          Caption.Text = #24453#34389#29702#20295#21015
          DataBinding.ValueType = 'Integer'
          Width = 130
          Position.ColIndex = 3
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
      end
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 52
    Width = 669
    Height = 34
    Align = alTop
    BevelOuter = bvNone
    BorderWidth = 2
    TabOrder = 5
    object Panel2: TPanel
      Left = 2
      Top = 2
      Width = 665
      Height = 30
      Align = alClient
      BevelOuter = bvNone
      Color = clGray
      ParentBackground = False
      TabOrder = 0
      object lblConfigName: TLabel
        Left = 12
        Top = 7
        Width = 90
        Height = 14
        Caption = 'lblConfigName'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWhite
        Font.Height = -12
        Font.Name = 'Verdana'
        Font.Style = []
        ParentFont = False
      end
      object cmbConfigList: TcxImageComboBox
        Left = 108
        Top = 4
        ParentFont = False
        Properties.Images = StyleModule.BarImageList
        Properties.Items = <>
        Properties.OnChange = cmbConfigListPropertiesChange
        Style.LookAndFeel.Kind = lfOffice11
        Style.StyleController = StyleModule.cxEditStyle
        Style.TransparentBorder = True
        StyleDisabled.Color = clWindow
        StyleDisabled.LookAndFeel.Kind = lfOffice11
        StyleFocused.LookAndFeel.Kind = lfOffice11
        StyleHot.LookAndFeel.Kind = lfOffice11
        TabOrder = 0
        Width = 280
      end
    end
  end
  object dxBarManager: TdxBarManager
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = #26032#32048#26126#39636
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
    ImageOptions.Images = StyleModule.BarImageList
    PopupMenuLinks = <>
    Style = bmsOffice11
    UseSystemFont = True
    Left = 56
    Top = 196
    DockControlHeights = (
      0
      0
      52
      26)
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
      ItemLinks = <
        item
          Visible = True
          ItemName = 'dxNagraServer'
        end
        item
          Visible = True
          ItemName = 'dxEnableControlSend'
        end
        item
          Visible = True
          ItemName = 'dxEnableFeedbackRecv'
        end>
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
          ItemName = 'dxMenuItem1'
        end
        item
          Visible = True
          ItemName = 'dxMenuItem2'
        end
        item
          Visible = True
          ItemName = 'dxMenuItem3'
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
          ItemName = 'dxDbTree'
        end
        item
          Visible = True
          ItemName = 'dxControl'
        end
        item
          Visible = True
          ItemName = 'dxConsole'
        end
        item
          Visible = True
          ItemName = 'dxFeedback'
        end
        item
          Visible = True
          ItemName = 'dxAutoScroll'
        end
        item
          BeginGroup = True
          Visible = True
          ItemName = 'dxCmdExpand'
        end
        item
          Visible = True
          ItemName = 'dxCmdCollapse'
        end>
      NotDocking = [dsLeft, dsRight, dsBottom]
      OldName = 'ToolBar'
      OneOnRow = True
      Row = 1
      UseOwnFont = False
      Visible = True
      WholeRow = True
    end
    object dxMenuItem1: TdxBarSubItem
      Caption = #31995#32113
      Category = 0
      Visible = ivAlways
      ItemLinks = <
        item
          Visible = True
          ItemName = 'dxMenuItem11'
        end>
    end
    object dxMenuItem2: TdxBarSubItem
      Caption = #27298#35222
      Category = 0
      Visible = ivAlways
      ItemLinks = <
        item
          Visible = True
          ItemName = 'dxMenuItem21'
        end
        item
          BeginGroup = True
          Visible = True
          ItemName = 'dxMenuItem22'
        end
        item
          Visible = True
          ItemName = 'dxMenuItem23'
        end
        item
          Visible = True
          ItemName = 'dxMenuItem24'
        end
        item
          Visible = True
          ItemName = 'dxMenuItem25'
        end
        item
          BeginGroup = True
          Visible = True
          ItemName = 'dxMenuItem26'
        end
        item
          Visible = True
          ItemName = 'dxMenuItem27'
        end
        item
          Visible = True
          ItemName = 'dxMenuItem28'
        end
        item
          Visible = True
          ItemName = 'dxMenuItem29'
        end>
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
      OnClick = dxMenuItem3Click
    end
    object dxMenuItem22: TdxBarButton
      Caption = #31995#32113#21488#29376#24907
      Category = 0
      Hint = #31995#32113#21488#29376#24907
      Visible = ivAlways
      OnClick = dxMenuItem22Click
    end
    object dxMenuItem23: TdxBarButton
      Caption = #25351#20196#20659#36865
      Category = 0
      Hint = #25351#20196#20659#36865
      Visible = ivAlways
      OnClick = dxMenuItem22Click
    end
    object dxMenuItem24: TdxBarButton
      Caption = #21629#20196#27169#24335
      Category = 0
      Hint = #21629#20196#27169#24335
      Visible = ivAlways
      OnClick = dxMenuItem22Click
    end
    object dxMenuItem25: TdxBarButton
      Caption = #22238#20659#27231#21046
      Category = 0
      Hint = #22238#20659#27231#21046
      Visible = ivAlways
      OnClick = dxMenuItem22Click
    end
    object dxMenuItem26: TdxBarButton
      Caption = #21855#21205#35338#24687
      Category = 0
      Hint = #21855#21205#35338#24687
      Visible = ivAlways
      OnClick = dxMenuItem22Click
    end
    object dxMenuItem27: TdxBarButton
      Caption = #31995#32113#21488#35338#24687
      Category = 0
      Hint = #31995#32113#21488#35338#24687
      Visible = ivAlways
      OnClick = dxMenuItem22Click
    end
    object dxMenuItem28: TdxBarButton
      Caption = #20659#36865#25351#20196#35338#24687
      Category = 0
      Hint = #20659#36865#25351#20196#35338#24687
      Visible = ivAlways
      OnClick = dxMenuItem22Click
    end
    object dxMenuItem29: TdxBarButton
      Caption = #22238#20659#27231#21046#35338#24687
      Category = 0
      Hint = #22238#20659#27231#21046#35338#24687
      Visible = ivAlways
      OnClick = dxMenuItem22Click
    end
    object dxMenuItem21: TdxBarButton
      Caption = #22238#24489#33267#38928#35373#20301#32622
      Category = 0
      Hint = #22238#24489#33267#38928#35373#20301#32622
      Visible = ivAlways
      OnClick = dxMenuItem21Click
    end
    object dxNagraServer: TdxBarStatic
      Caption = 'NAGRA'#20027#27231
      Category = 1
      Hint = 'NAGRA'#20027#27231
      Visible = ivAlways
      BorderStyle = sbsLowered
      ImageIndex = 5
      LeftIndent = 2
    end
    object dxEnableControlSend: TdxBarStatic
      Caption = #21855#29992#20659#36865#25351#20196
      Category = 1
      Hint = #21855#29992#20659#36865#25351#20196
      Visible = ivAlways
      BorderStyle = sbsLowered
      ImageIndex = 18
      LeftIndent = 3
    end
    object dxEnableFeedbackRecv: TdxBarStatic
      Caption = #21855#29992#22238#20659#27231#21046
      Category = 1
      Hint = #21855#29992#22238#20659#27231#21046
      Visible = ivAlways
      BorderStyle = sbsLowered
      ImageIndex = 18
      LeftIndent = 3
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
    object dxDbTree: TdxBarButton
      Caption = #39023#31034#31995#32113#21488#29376#24907
      Category = 2
      Hint = #39023#31034#31995#32113#21488#29376#24907
      Visible = ivAlways
      ButtonStyle = bsChecked
      ImageIndex = 12
      OnClick = dxDbTreeClick
    end
    object dxControl: TdxBarButton
      Caption = #39023#31034'('#20659#36865#25351#20196')'#27169#24335
      Category = 2
      Hint = #39023#31034'('#20659#36865#25351#20196')'#27169#24335
      Visible = ivAlways
      ButtonStyle = bsChecked
      ImageIndex = 13
      OnClick = dxControlClick
    end
    object dxConsole: TdxBarButton
      Caption = #39023#31034#21629#20196#27169#24335
      Category = 2
      Hint = #21855#29992#35264#23519#25351#20196#27169#24335
      Visible = ivAlways
      ButtonStyle = bsChecked
      ImageIndex = 10
      OnClick = dxConsoleClick
    end
    object dxFeedback: TdxBarButton
      Caption = #39023#31034'('#22238#20659#36039#35338')'#27169#24335
      Category = 2
      Hint = #39023#31034'('#22238#20659#36039#35338')'#27169#24335
      Visible = ivAlways
      ButtonStyle = bsChecked
      ImageIndex = 14
    end
    object dxAutoScroll: TdxBarButton
      Caption = #25351#20196#25458#21205
      Category = 2
      Hint = #33258#21205#25458#21205#25351#20196
      Visible = ivAlways
      ButtonStyle = bsChecked
      ImageIndex = 11
      OnClick = dxAutoScrollClick
    end
    object dxCmdExpand: TdxBarButton
      Caption = #23637#38283#25152#26377#25351#20196
      Category = 2
      Hint = #23637#38283#25152#26377#25351#20196
      Visible = ivAlways
      ImageIndex = 15
      OnClick = dxCmdExpandClick
    end
    object dxCmdCollapse: TdxBarButton
      Caption = #25240#30090#25152#26377#25351#20196
      Category = 2
      Hint = #25240#30090#25152#26377#25351#20196
      Visible = ivAlways
      ImageIndex = 16
      OnClick = dxCmdCollapseClick
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
      Items = (
        'dxCmdExpand'
        'dxCmdCollapse')
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
  object AutoExecTimer: TTimer
    Enabled = False
    Interval = 2200
    OnTimer = AutoExecTimerTimer
    Left = 89
    Top = 196
  end
  object LoadConfigTimer: TTimer
    Enabled = False
    Interval = 1500
    OnTimer = LoadConfigTimerTimer
    Left = 123
    Top = 196
  end
  object ControlSocket: TIdTCPClient
    MaxLineAction = maException
    ReadTimeout = 0
    Port = 0
    Left = 23
    Top = 247
  end
  object FeedbackSocket: TIdTCPClient
    MaxLineAction = maException
    ReadTimeout = 0
    Port = 0
    Left = 23
    Top = 283
  end
  object ControlSendLEDTimer: TTimer
    Enabled = False
    Interval = 150
    OnTimer = ControlSendLEDTimerTimer
    Left = 88
    Top = 242
  end
  object FeedbackLEDTimer: TTimer
    Enabled = False
    Interval = 150
    OnTimer = FeedbackLEDTimerTimer
    Left = 124
    Top = 242
  end
end
