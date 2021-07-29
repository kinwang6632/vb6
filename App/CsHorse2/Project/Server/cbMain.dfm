object fmMain: TfmMain
  Left = 342
  Top = 169
  Width = 679
  Height = 555
  Caption = #38283#21338#31185#25216' - '#23458#26381#23567#24171#25163#25511#21046#21488
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object dxDockHost: TdxDockSite
    Left = 0
    Top = 52
    Width = 671
    Height = 445
    Align = alClient
    DockType = 0
    OriginalWidth = 671
    OriginalHeight = 445
    object dxLayoutDockSite2: TdxLayoutDockSite
      Left = 260
      Top = 0
      Width = 411
      Height = 445
      DockType = 1
      OriginalWidth = 300
      OriginalHeight = 200
      object dxLayoutDockSite3: TdxLayoutDockSite
        Left = 0
        Top = 0
        Width = 411
        Height = 295
        DockType = 1
        OriginalWidth = 300
        OriginalHeight = 200
        object dxLayoutDockSite1: TdxLayoutDockSite
          Left = 0
          Top = 0
          Width = 411
          Height = 295
          DockType = 1
          OriginalWidth = 300
          OriginalHeight = 200
        end
        object dxDockPanel2: TdxDockPanel
          Left = 0
          Top = 0
          Width = 411
          Height = 295
          AllowFloating = True
          AutoHide = False
          Caption = #20351#29992#32773#35222#31383
          CaptionButtons = []
          ImageIndex = 4
          DockType = 1
          OriginalWidth = 185
          OriginalHeight = 124
          object UserGrid: TcxGrid
            Left = 0
            Top = 0
            Width = 407
            Height = 271
            Align = alClient
            TabOrder = 0
            object gvUser: TcxGridDBTableView
              NavigatorButtons.ConfirmDelete = False
              FilterBox.CustomizeDialog = False
              OnCustomDrawCell = gvUserCustomDrawCell
              DataController.DataModeController.SyncMode = False
              DataController.DataSource = MainUIModule.dsUserBuffer
              DataController.Summary.DefaultGroupSummaryItems = <>
              DataController.Summary.FooterSummaryItems = <>
              DataController.Summary.SummaryGroups = <>
              OptionsData.Deleting = False
              OptionsData.Editing = False
              OptionsData.Inserting = False
              OptionsSelection.CellSelect = False
              OptionsView.CellEndEllipsis = True
              OptionsView.GroupByBox = False
              OptionsView.GroupRowStyle = grsOffice11
              OptionsView.HeaderEndEllipsis = True
              object gvUserCol1: TcxGridDBColumn
                DataBinding.FieldName = 'SessionId'
                Width = 150
              end
              object gvUserCol2: TcxGridDBColumn
                Caption = #20844#21496#21029
                DataBinding.FieldName = 'CompName'
                Width = 105
              end
              object gvUserCol3: TcxGridDBColumn
                Caption = #30331#20837#20195#30908
                DataBinding.FieldName = 'UserId'
                Width = 110
              end
              object gvUserCol4: TcxGridDBColumn
                Caption = #20351#29992#32773
                DataBinding.FieldName = 'UserName'
                Width = 120
              end
              object gvUserCol15: TcxGridDBColumn
                Caption = #29376#24907
                DataBinding.FieldName = 'Status'
                PropertiesClassName = 'TcxImageComboBoxProperties'
                Properties.Items = <
                  item
                    ImageIndex = 6
                    Value = '0'
                  end
                  item
                    ImageIndex = 7
                    Value = '1'
                  end
                  item
                    ImageIndex = 8
                    Value = '2'
                  end>
                HeaderAlignmentHorz = taCenter
                Width = 70
              end
              object gvUserCol5: TcxGridDBColumn
                Caption = 'IP'#20301#22336
                DataBinding.FieldName = 'IP'
                Width = 126
              end
              object gvUserCol6: TcxGridDBColumn
                Caption = #38651#33126#21517#31281
                DataBinding.FieldName = 'HostName'
                Width = 128
              end
              object gvUserCol7: TcxGridDBColumn
                Caption = #24037#20316#32676#32068
                DataBinding.FieldName = 'WorkClassName'
                Width = 130
              end
              object gvUserCol8: TcxGridDBColumn
                Caption = #32066#31471#27231#35672#21029#30908
                DataBinding.FieldName = 'TermSId'
                Width = 130
              end
              object gvUserCol9: TcxGridDBColumn
                Caption = #32066#31471#27231#24037#20316#38542#27573
                DataBinding.FieldName = 'TermSName'
                Width = 130
              end
              object gvUserCol10: TcxGridDBColumn
                Caption = #32066#31471#27231#29992#25142#38651#33126#21517#31281
                DataBinding.FieldName = 'TermSPC'
                Width = 155
              end
              object gvUserCol11: TcxGridDBColumn
                Caption = #32066#31471#27231#29992#25142'IP'#20301#22336
                DataBinding.FieldName = 'TermSIP'
                Width = 155
              end
              object gvUserCol12: TcxGridDBColumn
                Caption = #32066#31471#27231#29992#25142#29376#24907
                DataBinding.FieldName = 'TermState'
                Width = 130
              end
              object gvUserCol13: TcxGridDBColumn
                Caption = #19978#32218#26178#38291
                DataBinding.FieldName = 'OnTime'
                Width = 130
              end
              object gvUserCol14: TcxGridDBColumn
                Caption = #38626#32218#26178#38291
                DataBinding.FieldName = 'OffTime'
                Width = 130
              end
              object gvUserCol20: TcxGridDBColumn
                Caption = #20844#21496#21029#20195#30908
                DataBinding.FieldName = 'CompCode'
                Visible = False
              end
              object gvUserCol21: TcxGridDBColumn
                Caption = #24037#20316#32676#32068#20195#30908
                DataBinding.FieldName = 'WorkClassCode'
                Visible = False
              end
            end
            object glUser: TcxGridLevel
              GridView = gvUser
            end
          end
        end
      end
      object dxTabContainerDockSite2: TdxTabContainerDockSite
        Left = 0
        Top = 295
        Width = 411
        Height = 150
        ActiveChildIndex = 2
        AllowFloating = True
        AutoHide = False
        DockType = 5
        OriginalWidth = 185
        OriginalHeight = 150
        object dxDockPanel5: TdxDockPanel
          Left = 0
          Top = 0
          Width = 407
          Height = 102
          AllowFloating = True
          AutoHide = False
          Caption = #21855#21205'/'#22519#34892#35338#24687
          ImageIndex = 8
          DockType = 1
          OriginalWidth = 165
          OriginalHeight = 150
          object MainMsgList: TcxListView
            Left = 0
            Top = 0
            Width = 407
            Height = 102
            Align = alClient
            ColumnClick = False
            Columns = <
              item
                Width = 2000
              end>
            HideSelection = False
            ParentFont = False
            ReadOnly = True
            RowSelect = True
            ShowColumnHeaders = False
            SmallImages = StyleModule.MsgImgList
            TabOrder = 0
            ViewStyle = vsReport
          end
        end
        object dxDockPanel8: TdxDockPanel
          Left = 0
          Top = 0
          Width = 407
          Height = 102
          AllowFloating = True
          AutoHide = False
          Caption = #36039#26009#21516#27493#35338#24687#35222#31383
          ImageIndex = 8
          DockType = 1
          OriginalWidth = 185
          OriginalHeight = 150
          object SoSyncMsgList: TcxListView
            Left = 0
            Top = 0
            Width = 407
            Height = 102
            Align = alClient
            ColumnClick = False
            Columns = <
              item
                Width = 2000
              end>
            HideSelection = False
            ParentFont = False
            ReadOnly = True
            RowSelect = True
            ShowColumnHeaders = False
            SmallImages = StyleModule.MsgImgList
            TabOrder = 0
            ViewStyle = vsReport
          end
        end
        object dxDockPanel7: TdxDockPanel
          Left = 0
          Top = 0
          Width = 407
          Height = 102
          AllowFloating = True
          AutoHide = False
          Caption = #20351#29992#32773#30331#20986'/'#20837
          ImageIndex = 8
          DockType = 1
          OriginalWidth = 165
          OriginalHeight = 150
          object UserMsgList: TcxListView
            Left = 0
            Top = 0
            Width = 407
            Height = 102
            Align = alClient
            ColumnClick = False
            Columns = <
              item
                Width = 2000
              end>
            HideSelection = False
            ParentFont = False
            ReadOnly = True
            RowSelect = True
            ShowColumnHeaders = False
            SmallImages = StyleModule.MsgImgList
            TabOrder = 0
            ViewStyle = vsReport
          end
        end
      end
    end
    object dxTabContainerDockSite1: TdxTabContainerDockSite
      Left = 0
      Top = 0
      Width = 260
      Height = 445
      ActiveChildIndex = 1
      AllowFloating = True
      AutoHide = False
      DockType = 2
      OriginalWidth = 260
      OriginalHeight = 140
      object dxDockPanel1: TdxDockPanel
        Left = 0
        Top = 0
        Width = 256
        Height = 397
        AllowFloating = True
        AutoHide = False
        Caption = 'TCP/IP '#30435#32893#29376#24907
        ImageIndex = 5
        DockType = 1
        OriginalWidth = 260
        OriginalHeight = 140
        object SoTcpTree: TcxTreeList
          Left = 0
          Top = 0
          Width = 256
          Height = 397
          Align = alClient
          Bands = <
            item
            end>
          BufferedPaint = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Images = StyleModule.SoImgList
          OptionsBehavior.Sorting = False
          OptionsData.Editing = False
          OptionsData.Deleting = False
          OptionsSelection.CellSelect = False
          OptionsView.CellEndEllipsis = True
          OptionsView.TreeLineStyle = tllsNone
          ParentFont = False
          TabOrder = 0
          object SoTcpTreeCol1: TcxTreeListColumn
            Caption.Text = #20195#30908
            DataBinding.ValueType = 'String'
            Width = 100
            Position.ColIndex = 0
            Position.RowIndex = 0
            Position.BandIndex = 0
          end
          object SoTcpTreeCol2: TcxTreeListColumn
            Caption.Text = #31995#32113#21488
            DataBinding.ValueType = 'String'
            Width = 110
            Position.ColIndex = 1
            Position.RowIndex = 0
            Position.BandIndex = 0
          end
          object SoTcpTreeCol3: TcxTreeListColumn
            PropertiesClassName = 'TcxTextEditProperties'
            Properties.Alignment.Horz = taRightJustify
            Caption.AlignHorz = taRightJustify
            Caption.Text = #19978#32218#20154#25976
            DataBinding.ValueType = 'String'
            Width = 95
            Position.ColIndex = 2
            Position.RowIndex = 0
            Position.BandIndex = 0
          end
          object SoTcpTreeCol4: TcxTreeListColumn
            PropertiesClassName = 'TcxTextEditProperties'
            Properties.Alignment.Horz = taRightJustify
            Caption.AlignHorz = taRightJustify
            Caption.Text = #38626#32218#20154#25976
            DataBinding.ValueType = 'String'
            Width = 95
            Position.ColIndex = 3
            Position.RowIndex = 0
            Position.BandIndex = 0
          end
        end
      end
      object dxDockPanel4: TdxDockPanel
        Left = 0
        Top = 0
        Width = 256
        Height = 397
        AllowFloating = True
        AutoHide = False
        Caption = #36039#26009#24235#38598#21312#29376#24907
        ImageIndex = 0
        DockType = 1
        OriginalWidth = 260
        OriginalHeight = 140
        object SoSessionTree: TcxTreeList
          Left = 0
          Top = 0
          Width = 256
          Height = 397
          Align = alClient
          Bands = <
            item
            end>
          BufferedPaint = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          Images = StyleModule.SoImgList
          OptionsBehavior.Sorting = False
          OptionsData.Editing = False
          OptionsData.Deleting = False
          OptionsSelection.CellSelect = False
          OptionsView.CellEndEllipsis = True
          OptionsView.TreeLineStyle = tllsNone
          ParentFont = False
          StateImages = StyleModule.SoImgList
          TabOrder = 0
          object SoSessionTreeCol1: TcxTreeListColumn
            Caption.Text = #20195#30908
            DataBinding.ValueType = 'String'
            Width = 100
            Position.ColIndex = 0
            Position.RowIndex = 0
            Position.BandIndex = 0
          end
          object SoSessionTreeCol2: TcxTreeListColumn
            Caption.Text = #31995#32113#21488
            DataBinding.ValueType = 'String'
            Width = 110
            Position.ColIndex = 1
            Position.RowIndex = 0
            Position.BandIndex = 0
          end
          object SoSessionTreeCol3: TcxTreeListColumn
            PropertiesClassName = 'TcxTextEditProperties'
            Properties.Alignment.Horz = taRightJustify
            Caption.AlignHorz = taRightJustify
            Caption.Text = #38598#21312#21487#29992#25976#37327
            DataBinding.ValueType = 'String'
            Width = 95
            Position.ColIndex = 2
            Position.RowIndex = 0
            Position.BandIndex = 0
          end
          object SoSessionTreeCol4: TcxTreeListColumn
            PropertiesClassName = 'TcxTextEditProperties'
            Properties.Alignment.Horz = taRightJustify
            Caption.AlignHorz = taRightJustify
            Caption.Text = #38598#21312#20351#29992#25976#37327
            DataBinding.ValueType = 'String'
            Width = 95
            Position.ColIndex = 3
            Position.RowIndex = 0
            Position.BandIndex = 0
          end
          object SoSessionTreeCol5: TcxTreeListColumn
            PropertiesClassName = 'TcxTextEditProperties'
            Properties.Alignment.Horz = taRightJustify
            Caption.AlignHorz = taRightJustify
            Caption.Text = #38598#21312#21839#38988#25976#37327
            DataBinding.ValueType = 'String'
            Width = 95
            Position.ColIndex = 4
            Position.RowIndex = 0
            Position.BandIndex = 0
          end
        end
      end
    end
  end
  object dxBarManager: TdxBarManager
    AllowReset = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
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
    ImageOptions.Images = StyleModule.ToolBarImgList
    ImageOptions.StretchGlyphs = False
    PopupMenuLinks = <>
    Style = bmsUseLookAndFeel
    UseSystemFont = True
    Left = 56
    Top = 196
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
      ItemLinks = <
        item
          Visible = True
          ItemName = 'dxStatusItem1'
        end
        item
          Visible = True
          ItemName = 'dxStatusItem2'
        end>
      NotDocking = [dsNone]
      OldName = 'dxBarStatus'
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
      OldName = 'dxBarMenu'
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
          UserDefine = [udWidth]
          UserWidth = 210
          Visible = True
          ItemName = 'dxConfigList'
        end
        item
          BeginGroup = True
          Visible = True
          ItemName = 'dxAutoScroll'
        end
        item
          Visible = True
          ItemName = 'dxGroupByBox'
        end
        item
          Visible = True
          ItemName = 'dxExpand'
        end
        item
          Visible = True
          ItemName = 'dxCollapse'
        end>
      NotDocking = [dsLeft, dsRight, dsBottom]
      OldName = 'dxBarTool'
      OneOnRow = True
      Row = 1
      UseOwnFont = False
      Visible = True
      WholeRow = True
    end
    object dxMenuItem1: TdxBarSubItem
      Caption = #21443#25976#31649#29702
      Category = 0
      Visible = ivAlways
      ItemLinks = <
        item
          Visible = True
          ItemName = 'dxMenuItem11'
        end
        item
          Visible = True
          ItemName = 'dxMenuItem12'
        end>
    end
    object dxMenuItem2: TdxBarSubItem
      Caption = #35222#31383#27298#35222
      Category = 0
      Visible = ivAlways
      ItemLinks = <
        item
          Visible = True
          ItemName = 'dxMenuItem21'
        end>
    end
    object dxMenuItem11: TdxBarButton
      Caption = #21443#25976#35373#23450
      Category = 0
      Hint = #21443#25976#35373#23450
      Visible = ivAlways
      ImageIndex = 6
      OnClick = dxMenuItem11Click
    end
    object dxMenuItem12: TdxBarButton
      Caption = #20351#29992#32773#32676#32068#35373#23450
      Category = 0
      Hint = #20351#29992#32773#32676#32068#35373#23450
      Visible = ivAlways
      ImageIndex = 6
      OnClick = dxMenuItem12Click
    end
    object dxMenuItem3: TdxBarButton
      Caption = #38364#26044#26412#31243#24335
      Category = 0
      Hint = #38364#26044#26412#31243#24335
      Visible = ivAlways
      OnClick = dxMenuItem3Click
    end
    object dxMenuItem21: TdxBarButton
      Caption = #22238#24489#33267#38928#35373#20301#32622
      Category = 0
      Hint = #22238#24489#33267#38928#35373#20301#32622
      Visible = ivAlways
      OnClick = dxMenuItem21Click
    end
    object dxStatusItem1: TdxBarStatic
      Caption = #26412#27231#30435#32893'IP'#20301#22336
      Category = 1
      Visible = ivAlways
      ImageIndex = 4
      LeftIndent = 2
    end
    object dxStatusItem2: TdxBarStatic
      Caption = #22519#34892#29376#24907
      Category = 1
      Visible = ivAlways
      ImageIndex = 19
      LeftIndent = 3
    end
    object dxPlay: TdxBarButton
      Caption = #21855#21205
      Category = 2
      Visible = ivAlways
      ImageIndex = 0
      OnClick = dxPlayClick
    end
    object dxStop: TdxBarButton
      Caption = #20572#27490
      Category = 2
      Enabled = False
      Visible = ivAlways
      ImageIndex = 2
      OnClick = dxStopClick
    end
    object dxAutoScroll: TdxBarButton
      Caption = #25351#20196#25458#21205
      Category = 2
      Visible = ivAlways
      ButtonStyle = bsChecked
      ImageIndex = 11
      OnClick = dxAutoScrollClick
    end
    object dxExpand: TdxBarButton
      Caption = #23637#38283
      Category = 2
      Visible = ivAlways
      ImageIndex = 16
      OnClick = dxExpandClick
    end
    object dxCollapse: TdxBarButton
      Caption = #25240#30090
      Category = 2
      Visible = ivAlways
      ImageIndex = 17
      OnClick = dxCollapseClick
    end
    object dxConfigList: TdxBarImageCombo
      Caption = #35373#23450#27284
      Category = 2
      Hint = #35373#23450#27284
      Visible = ivAlways
      OnChange = dxConfigListChange
      ShowCaption = True
      Width = 100
      Images = StyleModule.ImgList16
      Items.Strings = (
        '11'
        '12')
      ItemIndex = -1
      ImageIndexes = (
        7
        7)
    end
    object dxGroupByBox: TdxBarButton
      Caption = #32676#32068#20381#25818#26041#22602
      Category = 2
      Visible = ivAlways
      ButtonStyle = bsChecked
      ImageIndex = 28
      OnClick = dxGroupByBoxClick
    end
    object PlayGroup: TdxBarGroup
      Items = (
        'dxPlay')
    end
    object StopGroup: TdxBarGroup
      Enabled = False
      Items = (
        'dxStop')
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
    Options = [doActivateAfterDocking, doDblClickDocking, doFloatingOnTop, doTabContainerHasCaption, doTabContainerCanAutoHide, doSideContainerCanClose, doSideContainerCanAutoHide, doTabContainerCanInSideContainer]
    ViewStyle = vsUseLookAndFeel
    Left = 24
    Top = 196
  end
  object LoadConfigTimer: TTimer
    Enabled = False
    Interval = 1500
    OnTimer = LoadConfigOnTimer
    Left = 133
    Top = 197
  end
  object SocketWarningTimer: TTimer
    Enabled = False
    Interval = 300
    OnTimer = SocketWarningTimerTimer
    Left = 100
    Top = 197
  end
end
