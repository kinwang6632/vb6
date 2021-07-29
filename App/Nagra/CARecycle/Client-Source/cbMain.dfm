object fmMain: TfmMain
  Left = 202
  Top = 115
  Width = 810
  Height = 600
  Caption = 'CA'#22238#25910#20316#26989
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object dxBarDockControl1: TdxBarDockControl
    Left = 0
    Top = 553
    Width = 802
    Height = 20
    Align = dalBottom
    BarManager = dxBarManager
    SunkenBorder = False
    UseOwnSunkenBorder = True
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 802
    Height = 553
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 2
    TabOrder = 5
    object Bevel4: TBevel
      Left = 112
      Top = 2
      Width = 2
      Height = 549
      Align = alLeft
      Shape = bsSpacer
    end
    object Panel7: TPanel
      Left = 2
      Top = 2
      Width = 110
      Height = 549
      Align = alLeft
      BevelOuter = bvLowered
      TabOrder = 0
      object dxNavBar: TdxNavBar
        Left = 1
        Top = 1
        Width = 108
        Height = 547
        Align = alClient
        Locked = True
        ActiveGroupIndex = 0
        DefaultStyles.GroupBackground.BackColor = 13160661
        DefaultStyles.GroupBackground.BackColor2 = 8551805
        DefaultStyles.GroupBackground.GradientMode = gmForwardDiagonal
        DefaultStyles.GroupBackground.Font.Charset = DEFAULT_CHARSET
        DefaultStyles.GroupBackground.Font.Color = clWindowText
        DefaultStyles.GroupBackground.Font.Height = -11
        DefaultStyles.GroupBackground.Font.Name = 'MS Sans Serif'
        DefaultStyles.GroupBackground.Font.Style = []
        DefaultStyles.GroupBackground.HAlignment = haLeft
        DefaultStyles.Item.BackColor = clWhite
        DefaultStyles.Item.BackColor2 = clWhite
        DefaultStyles.Item.Font.Charset = ANSI_CHARSET
        DefaultStyles.Item.Font.Color = clBtnText
        DefaultStyles.Item.Font.Height = -12
        DefaultStyles.Item.Font.Name = 'Tahoma'
        DefaultStyles.Item.Font.Style = []
        DefaultStyles.ItemDisabled.BackColor = clWhite
        DefaultStyles.ItemDisabled.BackColor2 = clWhite
        DefaultStyles.ItemDisabled.Font.Charset = ANSI_CHARSET
        DefaultStyles.ItemDisabled.Font.Color = clBtnShadow
        DefaultStyles.ItemDisabled.Font.Height = -12
        DefaultStyles.ItemDisabled.Font.Name = 'Tahoma'
        DefaultStyles.ItemDisabled.Font.Style = []
        DefaultStyles.ItemHotTracked.BackColor = clWhite
        DefaultStyles.ItemHotTracked.BackColor2 = clWhite
        DefaultStyles.ItemHotTracked.Font.Charset = ANSI_CHARSET
        DefaultStyles.ItemHotTracked.Font.Color = clBtnText
        DefaultStyles.ItemHotTracked.Font.Height = -12
        DefaultStyles.ItemHotTracked.Font.Name = 'Tahoma'
        DefaultStyles.ItemHotTracked.Font.Style = []
        DefaultStyles.ItemPressed.BackColor = clWhite
        DefaultStyles.ItemPressed.BackColor2 = clWhite
        DefaultStyles.ItemPressed.Font.Charset = ANSI_CHARSET
        DefaultStyles.ItemPressed.Font.Color = clBtnText
        DefaultStyles.ItemPressed.Font.Height = -12
        DefaultStyles.ItemPressed.Font.Name = 'Tahoma'
        DefaultStyles.ItemPressed.Font.Style = []
        DragCopyCursor = -1119
        DragCursor = -1120
        DragDropFlags = [fAllowDragLink, fAllowDropLink, fAllowDragGroup, fAllowDropGroup]
        HotTrackedGroupCursor = crDefault
        HotTrackedLinkCursor = -1118
        LargeImages = StyleModule.BarLargeImageList
        ShowGroupCaptions = False
        View = 3
        OnLinkClick = dxNavBarLinkClick
        object dxNavBarGroup1: TdxNavBarGroup
          Caption = 'dxNavBarGroup1'
          LinksUseSmallImages = False
          SelectedLinkIndex = -1
          ShowAsIconView = False
          ShowControl = False
          TopVisibleLinkIndex = 0
          UseControl = False
          UseSmallImages = True
          Visible = True
          Links = <
            item
              Item = dxInput
            end
            item
              Item = dxList
            end
            item
              Item = dxOption
            end>
        end
        object dxInput: TdxNavBarItem
          Caption = #22238#25910#36039#26009#36664#20837
          Enabled = True
          LargeImageIndex = 0
          SmallImageIndex = -1
          Visible = True
        end
        object dxList: TdxNavBarItem
          Caption = #22238#25910#36039#26009#30906#35469
          Enabled = True
          LargeImageIndex = 1
          SmallImageIndex = -1
          Visible = True
        end
        object dxOption: TdxNavBarItem
          Caption = #36984#38917#35373#23450
          Enabled = True
          LargeImageIndex = 2
          SmallImageIndex = -1
          Visible = True
        end
      end
    end
    object Panel8: TPanel
      Left = 114
      Top = 2
      Width = 686
      Height = 549
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 1
      object Bevel1: TBevel
        Left = 0
        Top = 37
        Width = 686
        Height = 4
        Align = alTop
        Shape = bsSpacer
      end
      object cxMainPage: TcxPageControl
        Left = 0
        Top = 41
        Width = 686
        Height = 508
        ActivePage = cxTabSheet1
        Align = alClient
        LookAndFeel.NativeStyle = False
        TabOrder = 0
        TabPosition = tpBottom
        OnChange = cxMainPageChange
        ClientRectBottom = 481
        ClientRectRight = 686
        ClientRectTop = 0
        object cxTabSheet1: TcxTabSheet
          Caption = '   '#19968'. '#22238#25910#20316#26989#36039#26009#36664#20837'   '
          ImageIndex = 0
          object Panel2: TPanel
            Left = 0
            Top = 0
            Width = 686
            Height = 481
            Align = alClient
            BevelOuter = bvNone
            BorderWidth = 1
            TabOrder = 0
            object Bevel3: TBevel
              Left = 511
              Top = 1
              Width = 4
              Height = 418
              Align = alRight
              Shape = bsSpacer
            end
            object Bevel2: TBevel
              Left = 1
              Top = 419
              Width = 684
              Height = 4
              Align = alBottom
              Shape = bsSpacer
            end
            object cxInputGrid: TcxGrid
              Left = 1
              Top = 1
              Width = 510
              Height = 418
              Align = alClient
              TabOrder = 0
              LookAndFeel.NativeStyle = False
              object gvInput: TcxGridDBTableView
                NavigatorButtons.ConfirmDelete = False
                DataController.DataSource = dsInput
                DataController.Summary.DefaultGroupSummaryItems = <>
                DataController.Summary.FooterSummaryItems = <>
                DataController.Summary.SummaryGroups = <>
                OptionsBehavior.PullFocusing = True
                OptionsCustomize.ColumnFiltering = False
                OptionsSelection.CellSelect = False
                OptionsSelection.MultiSelect = True
                OptionsView.GroupByBox = False
                OptionsView.Indicator = True
                object gvInputICCNO: TcxGridDBColumn
                  Caption = 'ICC'#21345#34399
                  DataBinding.FieldName = 'ICCNO'
                  Width = 134
                end
                object gvInputSTBNO: TcxGridDBColumn
                  Caption = 'STB'#24207#34399
                  DataBinding.FieldName = 'STBNO'
                  Width = 136
                end
                object gvInputSTBAUTOCB: TcxGridDBColumn
                  Caption = #26159#21542#22238#25765
                  DataBinding.FieldName = 'STBAUTOCB'
                  PropertiesClassName = 'TcxCheckBoxProperties'
                  Properties.Alignment = taCenter
                  Properties.ValueChecked = 1
                  Properties.ValueUnchecked = 0
                  Width = 70
                end
                object gvInputRCSTARTDATE: TcxGridDBColumn
                  Caption = #20316#26989#36215#22987#26085
                  DataBinding.FieldName = 'RCSTARTDATE'
                  Width = 157
                end
                object gvInputOPERATOR: TcxGridDBColumn
                  Caption = #20316#26989#20154#21729
                  DataBinding.FieldName = 'OPERATOR'
                end
              end
              object glInput: TcxGridLevel
                GridView = gvInput
              end
            end
            object Panel6: TScrollBox
              Left = 515
              Top = 1
              Width = 170
              Height = 418
              Align = alRight
              BevelKind = bkFlat
              BorderStyle = bsNone
              TabOrder = 1
              object Label1: TLabel
                Left = 16
                Top = 21
                Width = 58
                Height = 16
                Caption = 'STB'#24207#34399':'
                Transparent = True
              end
              object Label2: TLabel
                Left = 16
                Top = 77
                Width = 55
                Height = 16
                Caption = 'ICC'#21345#34399':'
                Transparent = True
              end
              object txtStbNo: TcxMaskEdit
                Left = 15
                Top = 41
                ParentFont = False
                Properties.MaskKind = emkRegExprEx
                Properties.EditMask = '\d\d\d\d\d\d\d\d\d\d\d\d'
                Properties.MaxLength = 0
                Properties.OnValidate = txtStbNoPropertiesValidate
                Style.StyleController = StyleModule.cxEditStyle
                TabOrder = 0
                Width = 118
              end
              object txtIccNo: TcxMaskEdit
                Left = 15
                Top = 101
                ParentFont = False
                Properties.MaskKind = emkRegExprEx
                Properties.EditMask = '\d\d\d\d\d\d\d\d\d\d\d\d'
                Properties.MaxLength = 0
                Properties.OnValidate = txtIccNoPropertiesValidate
                Style.StyleController = StyleModule.cxEditStyle
                TabOrder = 1
                Width = 118
              end
              object btnAdd: TcxButton
                Left = 15
                Top = 190
                Width = 89
                Height = 28
                Caption = #21152#20837
                TabOrder = 2
                OnClick = btnAddClick
                Glyph.Data = {
                  36040000424D3604000000000000360000002800000010000000100000000100
                  2000000000000004000000000000000000000000000000000000000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  00000000000000000000000000000000000000000000004C0004000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  00000000000000000000000000000000000000870012008A00AC000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  00000000000000000000000000000087001E179D2DE3149F27EE000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  000000000000000000000083002D1CA333ED50F48FFF0F8B1CF0000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  000000000000008F003D23AF3DF452EF8BFF49E480FF1B9032FF178F29FF1790
                  29FF159027FF158E27FF17902BFF0E8719EC0000000000000000000000000000
                  0000009900532ABA48F957F390FF4DE781FF4BE57FFF4BE67FFF4BE580FF49E3
                  7EFF47E17CFF45E07AFF45E581FF0F891DF300000000000000000000000000A1
                  056A41CA61FD65FA9BFF52ED86FF4FE882FF4DE680FF4BE47EFF49E27CFF47E0
                  7AFF45DE78FF44DD77FF47E37FFF108E1CF400000000000000000000000002A4
                  076F4ACE69FE98FFC1FF53F189FF53ED87FF51EA84FF4EE882FF4AE67EFF48E4
                  7CFF46E27AFF44E079FF49E781FF129420F60000000000000000000000000000
                  000000A700573CC357FA9FFFC6FF58F38EFF53F088FF60F092FF7FF5A9FF79F3
                  A6FF74F2A0FF71EE9EFF67F29CFF149B25F60000000000000000000000000000
                  00000000000000A8004132BE4CF5A5FFCCFF6CFA9FFF34BD50FF34B548FF32B5
                  49FF31B347FF2FB245FF31B448FF15A626EE0000000000000000000000000000
                  0000000000000000000000A8003029B93EEEB2FFDCFF16A827F0000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  000000000000000000000000000000A6002123B639E516A423F4000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  00000000000000000000000000000000000000B7001400B700B0000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  0000000000000000000000000000000000000000000000C70005000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  0000000000000000000000000000000000000000000000000000}
                LookAndFeel.Kind = lfFlat
              end
              object btnRemove: TcxButton
                Left = 15
                Top = 227
                Width = 89
                Height = 28
                Caption = #31227#38500
                TabOrder = 3
                OnClick = btnRemoveClick
                Glyph.Data = {
                  36040000424D3604000000000000360000002800000010000000100000000100
                  2000000000000004000000000000000000000000000000000000000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  A4140206B0CB030AAEE10000A67B000098070000000000000000000000000000
                  00000000A78500009A8B0000A72A000000000000000000000000000000000000
                  A98D1844F6FF194DF8FF1031D2FF0102ABD00000B62400000000000000010000
                  B1930928D7FF092ED7FF0313B3FF0000AC8F0000000000000000000000000103
                  B3D02451F9FF1F52FFFF1D4FFFFF1744E8FF040BB0EC0000B03A0000AC970D2E
                  DDFF1142F9FF0D3DF5FF0B3BF0FF041ABCFF0000A55D00000000000000000000
                  AE421832DBFF285BFFFF2456FFFF2253FFFF1B4BF1FF050DB1FB0F30DDFF164A
                  FEFF1344F9FF1041F6FF0E3EF6FF0A3CF0FF00009FBF00000000000000000000
                  BE511F37DDFF3A6FFFFF2C5EFFFF295AFFFF2657FFFF2052FCFF1C4FFFFF194A
                  FDFF1646FAFF1445FAFF0F3DF2FF020AB1E70000A84300000000000000000000
                  00010000C866121DC8F03D6AFBFF3567FFFF2C5DFFFF2859FFFF2253FFFF1D4E
                  FFFF1A4DFFFF123DEDFF0002ACCB0000BA1A0000000000000000000000000000
                  0000000000000000CC1C0000B6BA2E4EE7FF3668FFFF2E5EFFFF2859FFFF2254
                  FFFF163DEAFF0000A8BE0000AB0B000000000000000000000000000000000000
                  000000000000000000000000BF4D253FDFFF3B6DFFFF3464FFFF2E5EFFFF2759
                  FFFF1B46EAFF0001ACCF0000A908000000000000000000000000000000000000
                  000000000000000000000203C8C14B7CFFFF4170FFFF3B6BFFFF396CFFFF2D5E
                  FFFF2558FFFF1336D7FF0000B693000000000000000000000000000000000000
                  0000000000000000D92C263CDBFE5080FFFF4575FFFF3662FAFF0D14C3FE3C6D
                  FFFF2A5BFFFF2053FDFF0B1DC4FD0000C0470000000000000000000000000000
                  0000000000000000CB8B527CFAFF5081FFFF4B7DFFFF0B13C9EA0000BB420E15
                  C7E9386AFFFF2456FFFF1A4AF2FF0207B3E30000B50F00000000000000000000
                  000000000000131CDDE06A9CFFFF5788FFFF2B46E7FF0000CD5C000000000000
                  CD320F1BCBF23065FFFF1F51FFFF1439DDFF0000B19C00000000000000000000
                  0000000000000000DE913A52E4FE5782FBFF0101D0C200000000000000000000
                  00000000CC471426D6FA265AFFFF0F2EE3FF0103B8CF00000000000000000000
                  000000000000000000000000CF3C0000C0A40000CE3000000000000000000000
                  0000000000000000C4630001B8BC0000B5490000770200000000000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  0000000000000000000000000000000000000000000000000000}
                LookAndFeel.Kind = lfFlat
              end
              object btnSave: TcxButton
                Left = 16
                Top = 280
                Width = 89
                Height = 28
                Caption = #23384#27284
                TabOrder = 4
                OnClick = btnSaveClick
                Glyph.Data = {
                  36040000424D3604000000000000360000002800000010000000100000000100
                  2000000000000004000000000000000000000000000000000000000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  000000000000000000000000000000000000000000000000000000000000728A
                  0002B2430892A83D00E687583EECA9ADADF2BCB4ADF2CCCBC6F2CDCDD0F2B1B1
                  B1F599A2A8FF802F01EBA33C03E6A34010E3973203700000000000000000C555
                  059FBE6317FFCD7927FF9E8770FFB98759FFCE700DFFDCAF7AFFFEFFFFFFFAFA
                  FAFFD9E2EAFFAD7024FFDD9435FFECBE75FFA5430AF70000000000000000C968
                  0DFFC07028FFC36C1EFFA7876BFFAE7C5AFFBA5809FFC69A6CFFD9DFE5FFF8F8
                  F8FFF8FFFFFFA3631DFFD5872DFFE5AF69FFA5460AFF0000000000000000CB69
                  0CFFBB6823FFBF6619FFB4906EFFB37C60FFAF4900FFB5875EFFBFC6CDFFDEE0
                  E2FFFFFFFFFFA5641FFFD4842DFFE4B06CFFA44409FF0000000000000000C969
                  0AFFB8631FFFBC6219FFB07F55FFC1BFBDFFB0ACA8FF98918BFF9A8E84FFA397
                  8AFFAEA79FFF90581FFFA76723FFDA9542FFA6450DFF0000000000000000C967
                  0AFFB55D1BFFB65D17FFB85B0FFFBA5808FFBE5D0AFFC2620EFFC56710FFC869
                  13FFC96E16FFCE7824FFD07F2AFFD38932FFA5450EFF0000000000000000C868
                  09FFB05514FFB16933FFAF7142FFAE703FFFAF7341FFB17340FFB07642FFB277
                  41FFB47A43FFB47D46FFBD8041FFD1842DFFA6460EFF0000000000000000C765
                  08FFAC4B09FFB59C89FFE6F2FBFFE4EDF3FFE6EEF5FFE6EDF4FFE1E9EFFFDDE5
                  ECFFD5DEE4FFD1DCE6FFAD9A8AFFD07F26FFA4450EFF0000000000000000C865
                  09FFA84303FFB89984FFF4F9FCFFD5D4D2FFBEBDBCFFC0BFBEFFC1C0BFFFC2C0
                  BFFFCBCAC9FFDFE3E6FFB19B85FFCF7C22FFA4440DFF0000000000000000C765
                  08FFA23C00FFBA9B85FFFAFFFFFFF1F1F1FFECECECFFEBEBEBFFE8E8E8FFE4E4
                  E4FFE1E1E1FFE2E5E9FFB29B88FFCD7920FFA5440DFF0000000000000000C662
                  07FF9E3700FFBC9C8AFFFFFFFFFFDDDCDBFFC2C1C0FFC4C3C2FFC4C3C2FFC2C1
                  C0FFCECDCCFFE4E7EBFFB39D89FFCD771FFFA4440EFF0000000000000000C562
                  06FF9A3100FFBF9D8AFFFFFFFFFFFBFBFBFFEEEEEEFFEFEFEFFFE9E9E9FFE4E4
                  E4FFE4E4E4FFE6EAEDFFB5A089FFC2701DFFA9480DFF0000000000000000C460
                  05FA932800FFBE9D8DFFFFFFFFFFE5E4E3FFD0CFCEFFD1D0CFFFCFCECDFFCDCC
                  CBFFD4D3D2FFE9EDF1FFB9A38CFF8A4F14FFB1500EFA0000000000000000D06F
                  0884C25400FBC8A483FFD5D7D7FFD6D3D0FFD6D3D0FFD1CFCCFFCCCAC8FFC7C5
                  C3FFC1BEBBFFBDBDBDFFB1947BFFC25800FBB351098400000000000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  0000000000000000000000000000000000000000000000000000}
                LookAndFeel.Kind = lfFlat
              end
              object chkCallback: TcxCheckBox
                Left = 13
                Top = 143
                Caption = #26159#21542#22238#25765
                TabOrder = 5
                Width = 94
              end
            end
            object Panel5: TPanel
              Left = 1
              Top = 423
              Width = 684
              Height = 57
              Align = alBottom
              BevelOuter = bvNone
              TabOrder = 2
              object Shape1: TShape
                Left = 0
                Top = 0
                Width = 684
                Height = 57
                Align = alClient
                Brush.Style = bsClear
                Pen.Color = clBtnShadow
              end
              object lbInputMsg: TLabel
                Left = 14
                Top = 9
                Width = 609
                Height = 37
                AutoSize = False
                Caption = #23384#27284#20013','#35531#31245#24460'.......'
                Font.Charset = DEFAULT_CHARSET
                Font.Color = clBlue
                Font.Height = -13
                Font.Name = 'Tahoma'
                Font.Style = [fsBold]
                ParentFont = False
              end
            end
          end
        end
        object cxTabSheet2: TcxTabSheet
          Caption = '   '#20108'. '#22238#25910#20316#26989#29376#24907'   '
          ImageIndex = 1
          object Panel3: TPanel
            Left = 0
            Top = 0
            Width = 686
            Height = 481
            Align = alClient
            BevelOuter = bvNone
            BorderWidth = 2
            TabOrder = 0
            object Bevel5: TBevel
              Left = 510
              Top = 2
              Width = 4
              Height = 477
              Align = alRight
              Shape = bsSpacer
            end
            object Panel9: TScrollBox
              Left = 514
              Top = 2
              Width = 170
              Height = 477
              Align = alRight
              BevelKind = bkFlat
              BorderStyle = bsNone
              TabOrder = 0
              object btnSave2: TcxButton
                Left = 25
                Top = 363
                Width = 89
                Height = 28
                Caption = #23384#27284
                TabOrder = 0
                OnClick = btnSave2Click
                Glyph.Data = {
                  36040000424D3604000000000000360000002800000010000000100000000100
                  2000000000000004000000000000000000000000000000000000000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  000000000000000000000000000000000000000000000000000000000000728A
                  0002B2430892A83D00E687583EECA9ADADF2BCB4ADF2CCCBC6F2CDCDD0F2B1B1
                  B1F599A2A8FF802F01EBA33C03E6A34010E3973203700000000000000000C555
                  059FBE6317FFCD7927FF9E8770FFB98759FFCE700DFFDCAF7AFFFEFFFFFFFAFA
                  FAFFD9E2EAFFAD7024FFDD9435FFECBE75FFA5430AF70000000000000000C968
                  0DFFC07028FFC36C1EFFA7876BFFAE7C5AFFBA5809FFC69A6CFFD9DFE5FFF8F8
                  F8FFF8FFFFFFA3631DFFD5872DFFE5AF69FFA5460AFF0000000000000000CB69
                  0CFFBB6823FFBF6619FFB4906EFFB37C60FFAF4900FFB5875EFFBFC6CDFFDEE0
                  E2FFFFFFFFFFA5641FFFD4842DFFE4B06CFFA44409FF0000000000000000C969
                  0AFFB8631FFFBC6219FFB07F55FFC1BFBDFFB0ACA8FF98918BFF9A8E84FFA397
                  8AFFAEA79FFF90581FFFA76723FFDA9542FFA6450DFF0000000000000000C967
                  0AFFB55D1BFFB65D17FFB85B0FFFBA5808FFBE5D0AFFC2620EFFC56710FFC869
                  13FFC96E16FFCE7824FFD07F2AFFD38932FFA5450EFF0000000000000000C868
                  09FFB05514FFB16933FFAF7142FFAE703FFFAF7341FFB17340FFB07642FFB277
                  41FFB47A43FFB47D46FFBD8041FFD1842DFFA6460EFF0000000000000000C765
                  08FFAC4B09FFB59C89FFE6F2FBFFE4EDF3FFE6EEF5FFE6EDF4FFE1E9EFFFDDE5
                  ECFFD5DEE4FFD1DCE6FFAD9A8AFFD07F26FFA4450EFF0000000000000000C865
                  09FFA84303FFB89984FFF4F9FCFFD5D4D2FFBEBDBCFFC0BFBEFFC1C0BFFFC2C0
                  BFFFCBCAC9FFDFE3E6FFB19B85FFCF7C22FFA4440DFF0000000000000000C765
                  08FFA23C00FFBA9B85FFFAFFFFFFF1F1F1FFECECECFFEBEBEBFFE8E8E8FFE4E4
                  E4FFE1E1E1FFE2E5E9FFB29B88FFCD7920FFA5440DFF0000000000000000C662
                  07FF9E3700FFBC9C8AFFFFFFFFFFDDDCDBFFC2C1C0FFC4C3C2FFC4C3C2FFC2C1
                  C0FFCECDCCFFE4E7EBFFB39D89FFCD771FFFA4440EFF0000000000000000C562
                  06FF9A3100FFBF9D8AFFFFFFFFFFFBFBFBFFEEEEEEFFEFEFEFFFE9E9E9FFE4E4
                  E4FFE4E4E4FFE6EAEDFFB5A089FFC2701DFFA9480DFF0000000000000000C460
                  05FA932800FFBE9D8DFFFFFFFFFFE5E4E3FFD0CFCEFFD1D0CFFFCFCECDFFCDCC
                  CBFFD4D3D2FFE9EDF1FFB9A38CFF8A4F14FFB1500EFA0000000000000000D06F
                  0884C25400FBC8A483FFD5D7D7FFD6D3D0FFD6D3D0FFD1CFCCFFCCCAC8FFC7C5
                  C3FFC1BEBBFFBDBDBDFFB1947BFFC25800FBB351098400000000000000000000
                  0000000000000000000000000000000000000000000000000000000000000000
                  0000000000000000000000000000000000000000000000000000}
                LookAndFeel.Kind = lfFlat
              end
              object cxGroupBox1: TcxGroupBox
                Left = 7
                Top = 9
                Caption = ' '#26597#35426#26781#20214' '
                ParentFont = False
                Style.StyleController = StyleModule.cxEditStyle
                TabOrder = 1
                Height = 312
                Width = 153
                object rdoCondition1: TcxRadioButton
                  Left = 13
                  Top = 31
                  Width = 130
                  Height = 17
                  Caption = #23436#25104'('#24453#30906#35469')'
                  TabOrder = 0
                  OnClick = rdoCondition1Click
                  LookAndFeel.Kind = lfFlat
                end
                object rdoCondition2: TcxRadioButton
                  Left = 13
                  Top = 58
                  Width = 130
                  Height = 17
                  Caption = #34389#29702#20013
                  TabOrder = 1
                  OnClick = rdoCondition1Click
                  LookAndFeel.Kind = lfFlat
                end
                object rdoCondition3: TcxRadioButton
                  Left = 13
                  Top = 85
                  Width = 130
                  Height = 17
                  Caption = #22833#25943'('#24453#21462#28040')'
                  TabOrder = 2
                  OnClick = rdoCondition1Click
                  LookAndFeel.Kind = lfFlat
                end
                object rdoCondition4: TcxRadioButton
                  Left = 13
                  Top = 113
                  Width = 130
                  Height = 17
                  Caption = 'ICC'#21345#34399
                  TabOrder = 3
                  OnClick = rdoCondition1Click
                  LookAndFeel.Kind = lfFlat
                end
                object txtIccNo2: TcxMaskEdit
                  Left = 14
                  Top = 137
                  ParentFont = False
                  Properties.MaskKind = emkRegExprEx
                  Properties.EditMask = '\d\d\d\d\d\d\d\d\d\d\d\d'
                  Properties.MaxLength = 0
                  Properties.OnValidate = txtIccNoPropertiesValidate
                  Style.StyleController = StyleModule.cxEditStyle
                  TabOrder = 4
                  Width = 106
                end
                object rdoCondition5: TcxRadioButton
                  Left = 13
                  Top = 170
                  Width = 130
                  Height = 17
                  Caption = #34389#29702#26085#26399'('#36215#36804')'
                  TabOrder = 5
                  OnClick = rdoCondition1Click
                  LookAndFeel.Kind = lfFlat
                end
                object txtRcDateSt: TcxDateEdit
                  Left = 16
                  Top = 195
                  ParentFont = False
                  Style.StyleController = StyleModule.cxEditStyle
                  TabOrder = 6
                  Width = 106
                end
                object txtRcDateEd: TcxDateEdit
                  Left = 16
                  Top = 223
                  ParentFont = False
                  Style.StyleController = StyleModule.cxEditStyle
                  TabOrder = 7
                  Width = 106
                end
                object btnQuery: TcxButton
                  Left = 18
                  Top = 265
                  Width = 89
                  Height = 28
                  Caption = #26597#35426
                  TabOrder = 8
                  OnClick = btnQueryClick
                  Glyph.Data = {
                    36040000424D3604000000000000360000002800000010000000100000000100
                    2000000000000004000000000000000000000000000000000000000000000000
                    0000000000000000000000000000000000000000000000000000000000000000
                    0000000000000000000000000000000000000000000000000000000000000000
                    0000000000000000000000000000000000000000000000000000000000000000
                    0000000000000000FF018B8A7AC385827DA26B6B6B3F00000000000000000000
                    0000000000000000000000000000000000000000000000000000000000000000
                    00000000FF01094FFF974392F6FFEEE9DFFF86827FA700000000000000000000
                    0000000000000000000000000000000000000000000000000000000000000000
                    FF020D51FF9B439AFFFF6ADAFFFF5DADF5FF998F7ED900000000000000000000
                    00000000000000000000000000000000000000000000000000006B7E690F0041
                    FC9F469FFFFF6FDAFFFF50ACFFFF1357FFB9002AFF0A00000000000000000000
                    000000000000000000006164682E515459524C4E523F1617180698928B9E7091
                    B2FF61D3FFFF4EAAFFFF1657FFB4001FFF070000000000000000000000000000
                    000076777A257B7E80C8B3A081FED2B588FDC3AA83FD83817AE566686CEAFFF7
                    F0FF6B93BDFF084AFEAF0028FF06000000000000000000000000000000007B7D
                    7F188D8A82E6F5CB84FEF5CB84FFF1C885FFF5CE8EFFFCD08CFFAE9E85FE7071
                    75F3A9A193BB3956B20A0000000000000000000000000000000000000000767A
                    7F95ECCB8EFFF3D192FFEECE92FFEDCC8EFFECC784FFEDC687FFFDD28FFF8783
                    7CDB000000010000000000000000000000000000000000000000000000009894
                    8BE1FEDFA1FFF2DAA5FFF2DBA7FFF1D79FFFEFD097FFECC886FFF6D298FFC1A9
                    87FF5254592A000000000000000000000000000000000000000000000000A7A1
                    91F4FFEBB9FFF8ECC6FFF7EBC0FFF5E3B2FFF2DAA3FFEECF94FFF4D093FFCFB6
                    8DFF5153583D0000000000000000000000000000000000000000000000009996
                    92D0FFF7CBFFFDFAE6FFFDF9E7FFF8EDC5FFF4E0B0FFF0D49BFFFAD594FFAF9F
                    85FE6264671A000000000000000000000000000000000000000000000000A0A1
                    A56FD5CCB1FFFFFFF2FFFFFFF2FFFBF3D2FFF6E2B4FFF7D99EFFF6D393FF7D7D
                    81B800000000000000000000000000000000000000000000000000000000D9D9
                    DA04A1A1A2B8D6CFB6FFFFFFDCFFFFF6CAFFFFEBB1FFEDD49CFE8C8983E47677
                    7A20000000000000000000000000000000000000000000000000000000000000
                    0000EBECED059C9CA07B979893E0A8A397FE97948CF07679809E7C7D7F1A0000
                    0000000000000000000000000000000000000000000000000000000000000000
                    0000000000000000000000000000000000000000000000000000000000000000
                    0000000000000000000000000000000000000000000000000000}
                  LookAndFeel.Kind = lfFlat
                end
              end
            end
            object cxListPage: TcxPageControl
              Left = 2
              Top = 2
              Width = 508
              Height = 477
              ActivePage = cxTabSheet3
              Align = alClient
              Images = StyleModule.BarSmallImageList
              Style = 9
              TabOrder = 1
              ClientRectBottom = 477
              ClientRectRight = 508
              ClientRectTop = 23
              object cxTabSheet3: TcxTabSheet
                Caption = '   '#19981#38920#22238#25765'   '
                ImageIndex = 2
                object Bevel6: TBevel
                  Left = 0
                  Top = 0
                  Width = 508
                  Height = 3
                  Align = alTop
                  Shape = bsSpacer
                end
                object cxNoCallGrid: TcxGrid
                  Left = 0
                  Top = 3
                  Width = 508
                  Height = 451
                  Align = alClient
                  TabOrder = 0
                  LookAndFeel.NativeStyle = False
                  object gvNoCall: TcxGridDBBandedTableView
                    NavigatorButtons.ConfirmDelete = False
                    OnCellClick = gvNoCallCellClick
                    OnCustomDrawCell = gvNoCallCustomDrawCell
                    DataController.DataSource = dsNoCall
                    DataController.Options = [dcoAssignGroupingValues, dcoAssignMasterDetailKeys, dcoSaveExpanding, dcoImmediatePost]
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
                    OptionsView.RowSeparatorWidth = 5
                    Bands = <
                      item
                        Width = 220
                      end
                      item
                        Width = 425
                      end
                      item
                        FixedKind = fkRight
                        Width = 59
                      end>
                    object gvNoCallICC_NO: TcxGridDBBandedColumn
                      Caption = 'ICC'#21345#34399
                      DataBinding.FieldName = 'ICCNO'
                      Width = 110
                      Position.BandIndex = 0
                      Position.ColIndex = 0
                      Position.RowIndex = 0
                    end
                    object gvNoCallSTB_NO: TcxGridDBBandedColumn
                      Caption = 'STB'#24207#34399
                      DataBinding.FieldName = 'STBNO'
                      Width = 110
                      Position.BandIndex = 0
                      Position.ColIndex = 0
                      Position.RowIndex = 1
                    end
                    object gvNoCallRCSTARTDATE: TcxGridDBBandedColumn
                      Caption = #20316#26989#36215#22987#26085
                      DataBinding.FieldName = 'RCSTARTDATE'
                      Width = 127
                      Position.BandIndex = 0
                      Position.ColIndex = 1
                      Position.RowIndex = 0
                    end
                    object gvNoCallOPERATOR: TcxGridDBBandedColumn
                      Caption = #20316#26989#20154#21729
                      DataBinding.FieldName = 'OPERATOR'
                      Width = 127
                      Position.BandIndex = 0
                      Position.ColIndex = 1
                      Position.RowIndex = 1
                    end
                    object gvNoCallCONFIRM: TcxGridDBBandedColumn
                      Tag = 2
                      Caption = #30906#35469'/'#21462#28040
                      DataBinding.FieldName = 'CONFIRM'
                      HeaderAlignmentHorz = taCenter
                      HeaderAlignmentVert = vaCenter
                      Position.BandIndex = 2
                      Position.ColIndex = 1
                      Position.LineCount = 2
                      Position.RowIndex = 0
                    end
                    object gvNoCallERRMSG: TcxGridDBBandedColumn
                      Caption = #37679#35492#35338#24687
                      DataBinding.FieldName = 'ERRMSG'
                      HeaderAlignmentVert = vaCenter
                      Width = 280
                      Position.BandIndex = 1
                      Position.ColIndex = 0
                      Position.RowIndex = 2
                    end
                    object gvNoCallCMD52_1: TcxGridDBBandedColumn
                      Tag = 1
                      Caption = #37197#23565
                      DataBinding.FieldName = 'CMD52_1'
                      HeaderAlignmentHorz = taCenter
                      Width = 82
                      Position.BandIndex = 1
                      Position.ColIndex = 0
                      Position.RowIndex = 0
                    end
                    object gvNoCallCMD8: TcxGridDBBandedColumn
                      Tag = 1
                      Caption = #40670#25976#28165#38500
                      DataBinding.FieldName = 'CMD8'
                      HeaderAlignmentHorz = taCenter
                      Width = 81
                      Position.BandIndex = 1
                      Position.ColIndex = 1
                      Position.RowIndex = 0
                    end
                    object gvNoCallCMD97_96: TcxGridDBBandedColumn
                      Tag = 1
                      Caption = #25910#35222#35352#37636
                      DataBinding.FieldName = 'CMD97_96'
                      HeaderAlignmentHorz = taCenter
                      Width = 83
                      Position.BandIndex = 1
                      Position.ColIndex = 2
                      Position.RowIndex = 0
                    end
                    object gvNoCallCMD7: TcxGridDBBandedColumn
                      Tag = 1
                      Caption = #38364#38971#36947
                      DataBinding.FieldName = 'CMD7'
                      HeaderAlignmentHorz = taCenter
                      Width = 85
                      Position.BandIndex = 1
                      Position.ColIndex = 3
                      Position.RowIndex = 0
                    end
                    object gvNoCallCMD48: TcxGridDBBandedColumn
                      Tag = 1
                      Caption = #37109#36958#21312#34399
                      DataBinding.FieldName = 'CMD48'
                      HeaderAlignmentHorz = taCenter
                      Width = 94
                      Position.BandIndex = 1
                      Position.ColIndex = 4
                      Position.RowIndex = 0
                    end
                    object gvNoCallCMD52_2: TcxGridDBBandedColumn
                      Caption = #35299#37197#23565
                      DataBinding.FieldName = 'CMD52_2'
                      Visible = False
                      HeaderAlignmentHorz = taCenter
                      Width = 80
                      Position.BandIndex = 1
                      Position.ColIndex = 0
                      Position.RowIndex = 1
                    end
                    object gvNoCallCMDFLAG: TcxGridDBBandedColumn
                      DataBinding.FieldName = 'CMDFLAG'
                      Visible = False
                      Position.BandIndex = 2
                      Position.ColIndex = 0
                      Position.LineCount = 2
                      Position.RowIndex = 0
                    end
                    object gvNoCallRECYCLETEXT: TcxGridDBBandedColumn
                      DataBinding.FieldName = 'RECYCLETEXT'
                      Visible = False
                      Position.BandIndex = 0
                      Position.ColIndex = 2
                      Position.RowIndex = 0
                    end
                  end
                  object glNoCall: TcxGridLevel
                    GridView = gvNoCall
                  end
                end
              end
              object cxTabSheet4: TcxTabSheet
                Caption = '   '#38920#35201#22238#25765'   '
                ImageIndex = 2
                object Bevel7: TBevel
                  Left = 0
                  Top = 0
                  Width = 508
                  Height = 3
                  Align = alTop
                  Shape = bsSpacer
                end
                object cxCallGrid: TcxGrid
                  Left = 0
                  Top = 3
                  Width = 508
                  Height = 451
                  Align = alClient
                  TabOrder = 0
                  LookAndFeel.NativeStyle = False
                  object gvCall: TcxGridDBBandedTableView
                    NavigatorButtons.ConfirmDelete = False
                    OnCellClick = gvNoCallCellClick
                    OnCustomDrawCell = gvNoCallCustomDrawCell
                    DataController.DataSource = dsCall
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
                    OptionsView.RowSeparatorWidth = 5
                    Bands = <
                      item
                        Width = 220
                      end
                      item
                        Width = 602
                      end
                      item
                        FixedKind = fkRight
                        Width = 67
                      end>
                    object gvCallICC_NO: TcxGridDBBandedColumn
                      Caption = 'ICC'#21345#34399
                      DataBinding.FieldName = 'ICCNO'
                      Width = 110
                      Position.BandIndex = 0
                      Position.ColIndex = 0
                      Position.RowIndex = 0
                    end
                    object gvCallSTBNO: TcxGridDBBandedColumn
                      Caption = 'STB'#24207#34399
                      DataBinding.FieldName = 'STBNO'
                      Width = 110
                      Position.BandIndex = 0
                      Position.ColIndex = 0
                      Position.RowIndex = 1
                    end
                    object gvCallRCSTARTDATE: TcxGridDBBandedColumn
                      Caption = #20316#26989#36215#22987#26085
                      DataBinding.FieldName = 'RCSTARTDATE'
                      Width = 127
                      Position.BandIndex = 0
                      Position.ColIndex = 1
                      Position.RowIndex = 0
                    end
                    object gvCallOPERATOR: TcxGridDBBandedColumn
                      Caption = #20316#26989#20154#21729
                      DataBinding.FieldName = 'OPERATOR'
                      Width = 127
                      Position.BandIndex = 0
                      Position.ColIndex = 1
                      Position.RowIndex = 1
                    end
                    object gvCallCONFIRM: TcxGridDBBandedColumn
                      Tag = 2
                      Caption = #30906#35469'/'#21462#28040
                      DataBinding.FieldName = 'CONFIRM'
                      HeaderAlignmentHorz = taCenter
                      HeaderAlignmentVert = vaCenter
                      Width = 48
                      Position.BandIndex = 2
                      Position.ColIndex = 0
                      Position.LineCount = 2
                      Position.RowIndex = 0
                    end
                    object gvCallERRMSG: TcxGridDBBandedColumn
                      Caption = #37679#35492#35338#24687
                      DataBinding.FieldName = 'ERRMSG'
                      HeaderAlignmentVert = vaCenter
                      Width = 188
                      Position.BandIndex = 1
                      Position.ColIndex = 0
                      Position.RowIndex = 2
                    end
                    object gvCallCMD52_1: TcxGridDBBandedColumn
                      Tag = 1
                      Caption = #37197#23565
                      DataBinding.FieldName = 'CMD52_1'
                      HeaderAlignmentHorz = taCenter
                      Width = 89
                      Position.BandIndex = 1
                      Position.ColIndex = 0
                      Position.RowIndex = 0
                    end
                    object gvCallCMD62: TcxGridDBBandedColumn
                      Tag = 1
                      Caption = #33258#21205#22238#25765'('#38364')'
                      DataBinding.FieldName = 'CMD62'
                      HeaderAlignmentHorz = taCenter
                      Width = 104
                      Position.BandIndex = 1
                      Position.ColIndex = 1
                      Position.RowIndex = 0
                    end
                    object gvCallCMD8: TcxGridDBBandedColumn
                      Tag = 1
                      Caption = #40670#25976#28165#38500
                      DataBinding.FieldName = 'CMD8'
                      HeaderAlignmentHorz = taCenter
                      Width = 77
                      Position.BandIndex = 1
                      Position.ColIndex = 2
                      Position.RowIndex = 0
                    end
                    object gvCallCMD60: TcxGridDBBandedColumn
                      Tag = 1
                      Caption = #31435#21363#22238#25765
                      DataBinding.FieldName = 'CMD60'
                      HeaderAlignmentHorz = taCenter
                      Width = 83
                      Position.BandIndex = 1
                      Position.ColIndex = 3
                      Position.RowIndex = 0
                    end
                    object gvCallCMD7: TcxGridDBBandedColumn
                      Tag = 1
                      Caption = #38364#38971#36947
                      DataBinding.FieldName = 'CMD7'
                      HeaderAlignmentHorz = taCenter
                      Width = 83
                      Position.BandIndex = 1
                      Position.ColIndex = 4
                      Position.RowIndex = 0
                    end
                    object gvCallCMD48: TcxGridDBBandedColumn
                      Tag = 1
                      Caption = #37109#36958#21312#34399
                      DataBinding.FieldName = 'CMD48'
                      HeaderAlignmentHorz = taCenter
                      Width = 85
                      Position.BandIndex = 1
                      Position.ColIndex = 5
                      Position.RowIndex = 0
                    end
                    object gvCallCMD99: TcxGridDBBandedColumn
                      Tag = 1
                      Caption = #31561#20505#22238#20659
                      DataBinding.FieldName = 'CMD99'
                      HeaderAlignmentHorz = taCenter
                      Width = 81
                      Position.BandIndex = 1
                      Position.ColIndex = 6
                      Position.RowIndex = 0
                    end
                    object gvCallCMD52_2: TcxGridDBBandedColumn
                      Caption = #35299#37197#23565
                      DataBinding.FieldName = 'CMD52_2'
                      Visible = False
                      HeaderAlignmentHorz = taCenter
                      Width = 91
                      Position.BandIndex = 1
                      Position.ColIndex = 0
                      Position.RowIndex = 1
                    end
                    object gvCallCMDFLAG: TcxGridDBBandedColumn
                      DataBinding.FieldName = 'CMDFLAG'
                      Visible = False
                      Position.BandIndex = 0
                      Position.ColIndex = 2
                      Position.RowIndex = 0
                    end
                    object gvCallRECYCLETEXT: TcxGridDBBandedColumn
                      DataBinding.FieldName = 'RECYCLETEXT'
                      Visible = False
                      Position.BandIndex = 0
                      Position.ColIndex = 3
                      Position.RowIndex = 0
                    end
                  end
                  object glCall: TcxGridLevel
                    GridView = gvCall
                  end
                end
              end
            end
          end
        end
      end
      object Panel4: TPanel
        Left = 0
        Top = 0
        Width = 686
        Height = 37
        Align = alTop
        BevelOuter = bvNone
        Color = clBtnShadow
        TabOrder = 1
        DesignSize = (
          686
          37)
        object lbTitle: TLabel
          Left = 17
          Top = 9
          Width = 37
          Height = 18
          Caption = 'lbTitle'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clHighlightText
          Font.Height = -15
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object cmbSo: TcxComboBox
          Left = 453
          Top = 6
          Anchors = [akTop, akRight]
          ParentFont = False
          Properties.DropDownListStyle = lsFixedList
          Properties.OnChange = cmbSoPropertiesChange
          Style.StyleController = StyleModule.cxEditStyle
          TabOrder = 0
          Width = 224
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
        Caption = 'Menu'
        DockedDockingStyle = dsTop
        DockedLeft = 0
        DockedTop = 0
        DockingStyle = dsTop
        FloatLeft = 464
        FloatTop = 357
        FloatClientWidth = 23
        FloatClientHeight = 22
        ItemLinks = <>
        Name = 'Menu'
        OneOnRow = True
        Row = 0
        UseOwnFont = False
        Visible = False
        WholeRow = False
      end
      item
        AllowClose = False
        AllowCustomizing = False
        AllowQuickCustomizing = False
        AllowReset = False
        Caption = 'Tool'
        DockedDockingStyle = dsTop
        DockedLeft = 0
        DockedTop = 0
        DockingStyle = dsTop
        FloatLeft = 423
        FloatTop = 410
        FloatClientWidth = 23
        FloatClientHeight = 22
        ItemLinks = <
          item
            Item = dxPg1
            Visible = True
          end
          item
            Item = dxPg2
            Visible = True
          end>
        Name = 'Tool'
        NotDocking = [dsNone, dsLeft, dsRight, dsBottom]
        OneOnRow = True
        Row = 0
        UseOwnFont = False
        Visible = False
        WholeRow = True
      end
      item
        AllowClose = False
        AllowCustomizing = False
        AllowQuickCustomizing = False
        BorderStyle = bbsNone
        Caption = 'Status'
        DockControl = dxBarDockControl1
        DockedDockControl = dxBarDockControl1
        DockedLeft = 0
        DockedTop = 0
        FloatLeft = 417
        FloatTop = 317
        FloatClientWidth = 23
        FloatClientHeight = 22
        ItemLinks = <
          item
            Item = dxCompName
            Visible = True
          end
          item
            Item = dxComputer
            Visible = True
          end
          item
            Item = dxLoginName
            Visible = True
          end>
        Name = 'Status'
        NotDocking = [dsNone, dsLeft, dsTop, dsRight]
        OneOnRow = True
        Row = 0
        UseOwnFont = False
        Visible = True
        WholeRow = True
      end>
    Categories.Strings = (
      'Default')
    Categories.ItemsVisibles = (
      2)
    Categories.Visibles = (
      True)
    HotImages = StyleModule.BarLargeImageList
    Images = StyleModule.BarSmallImageList
    LargeImages = StyleModule.BarLargeImageList
    PopupMenuLinks = <>
    Style = bmsOffice11
    UseSystemFont = True
    Left = 141
    Top = 365
    DockControlHeights = (
      0
      0
      0
      0)
    object dxPg1: TdxBarLargeButton
      Caption = #19968'.'#36039#26009#36664#20837
      Category = 0
      Hint = #19968'.'#36039#26009#36664#20837
      Visible = ivAlways
      ButtonStyle = bsChecked
      GroupIndex = 1
      Down = True
      HotImageIndex = 0
      LargeImageIndex = 0
      SyncImageIndex = False
      ImageIndex = -1
    end
    object dxPg2: TdxBarLargeButton
      Caption = #20108'.'#22238#25910#20316#26989
      Category = 0
      Hint = #20108'.'#22238#25910#20316#26989
      Visible = ivAlways
      ButtonStyle = bsChecked
      GroupIndex = 1
      HotImageIndex = 1
      LargeImageIndex = 1
      SyncImageIndex = False
      ImageIndex = -1
    end
    object dxCompName: TdxBarStatic
      Caption = #31995#32113#21488
      Category = 0
      Hint = #31995#32113#21488
      Visible = ivAlways
      BorderStyle = sbsLowered
      RightIndent = 2
    end
    object dxComputer: TdxBarStatic
      Caption = #38651#33126#21517#31281
      Category = 0
      Hint = #38651#33126#21517#31281
      Visible = ivAlways
      BorderStyle = sbsLowered
      RightIndent = 2
    end
    object dxLoginName: TdxBarStatic
      Caption = #20316#26989#20154#21729
      Category = 0
      Hint = #20316#26989#20154#21729
      Visible = ivAlways
      BorderStyle = sbsLowered
      RightIndent = 2
    end
  end
  object StartTimer: TTimer
    Enabled = False
    Interval = 500
    OnTimer = StartTimerTimer
    Left = 176
    Top = 365
  end
  object dsInput: TDataSource
    DataSet = SoDataModule.cdsInput
    Left = 214
    Top = 365
  end
  object dsNoCall: TDataSource
    DataSet = SoDataModule.cdsNoCall
    Left = 214
    Top = 404
  end
  object dsCall: TDataSource
    DataSet = SoDataModule.cdsCall
    Left = 212
    Top = 440
  end
end
