object frmCD078B3_6: TfrmCD078B3_6
  Left = 288
  Top = 160
  ActiveControl = RankGrid
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = 'frmCD078B3_6'
  ClientHeight = 491
  ClientWidth = 691
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Icon.Data = {
    0000010002002020100000000000E80200002600000010101000000000002801
    00000E0300002800000020000000400000000100040000000000800200000000
    0000000000000000000000000000000000000000800000800000008080008000
    0000800080008080000080808000C0C0C0000000FF0000FF000000FFFF00FF00
    0000FF00FF00FFFF0000FFFFFF00000000000000000000000000000000000000
    00000000BBBBBBBB000000000000000000000BBBBBBBBBBBBBB0000000000000
    0000BBBBBBBBBBBBBBBB00000000000000BBBBBBBBBBBBBBBBBBBB0000000000
    0BBBBBBBBBBBBBBBBBBBBBB000000000BBBBBBBBB0000000BBBBBBBB00000000
    BBBBBBB00BBBBBBB00BBBBBB0000000BBBBBBB00BBBBBBBBB00BBBBBB00000BB
    BBBBB00BBBBBBBBBBB00BBBBBB0000BBBBBBB0BBBBBBBBBBBBB0BBBBBB0000BB
    BBBB0BBBBBBBBBBBBBBB0BBBBB000BBBBBBB0BBBBBBBBBBBBBBB0BBBBBB00BBB
    BBBB0BBBBBBBBBBBBBBB0BBBBBB00BBBBBBBBBBBBBBBBBBBBBBBBBBBBBB00BBB
    BBBBBBBBBBBBBBBBBBBBBBBBBBB00BBBBBBBBBBBBBBBBBBBBBBBBBBBBBB00BBB
    BBBBBBB00BBBBBB00BBBBBBBBBB00BBBBBBBBB0000BBBB0000BBBBBBBBB00BBB
    BBBBBB0000BBBB0000BBBBBBBBB000BBBBBBBB0000BBBB0000BBBBBBBB0000BB
    BBBBBB0000BBBB0000BBBBBBBB0000BBBBBBBB0000BBBB0000BBBBBBBB00000B
    BBBBBBB00BBBBBB00BBBBBBBB0000000BBBBBBBBBBBBBBBBBBBBBBBB00000000
    BBBBBBBBBBBBBBBBBBBBBBBB000000000BBBBBBBBBBBBBBBBBBBBBB000000000
    00BBBBBBBBBBBBBBBBBBBB00000000000000BBBBBBBBBBBBBBBB000000000000
    00000BBBBBBBBBBBBBB000000000000000000000BBBBBBBB0000000000000000
    0000000000000000000000000000FFF00FFFFF8001FFFE00007FFC00003FF800
    001FF000000FE0000007C0000003C00000038000000180000001800000010000
    0000000000000000000000000000000000000000000000000000000000008000
    00018000000180000001C0000003C0000003E0000007F000000FF800001FFC00
    003FFE00007FFF8001FFFFF00FFF280000001000000020000000010004000000
    0000C00000000000000000000000000000000000000000000000000080000080
    00000080800080000000800080008080000080808000C0C0C0000000FF0000FF
    000000FFFF00FF000000FF00FF00FFFF0000FFFFFF0000000000000000000000
    0BBBBBB00000000BBBBBBBBBB00000BBB000003BBB0000BB00BBBB03BB000BBB
    0BBBBBB0BBB00BBB0BBBBBB0BBB00BBBBBBBBBBBBBB00BBBB00BB00BBBB00BBB
    B00BB00BBBB00BBBB00BB00BBBB000BBB00BB00BBB0000BBBBBBBBBBBB00000B
    BBBBBBBBB00000000BBBBBB000000000000000000000F81F0000E0070000C003
    0000840100008801000000000000000000000000000000000000000000000000
    00008001000080010000C0030000E0070000F81F0000}
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 459
    Width = 691
    Height = 32
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      691
      32)
    object edtDML: TcxTextEdit
      Left = 730
      Top = -4
      Anchors = [akRight, akBottom]
      Enabled = False
      ParentFont = False
      Properties.Alignment.Horz = taCenter
      Properties.ReadOnly = True
      Style.Color = 16777111
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.StyleController = StyleController.EditorStyle
      Style.IsFontAssigned = True
      StyleDisabled.Color = 16777111
      StyleDisabled.TextColor = clWindowText
      TabOrder = 0
      Width = 67
    end
    object btnSave: TButton
      Left = 494
      Top = 3
      Width = 69
      Height = 26
      Action = actSave
      TabOrder = 1
    end
    object Button1: TButton
      Left = 570
      Top = 3
      Width = 69
      Height = 26
      Action = actCancel
      TabOrder = 2
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 691
    Height = 459
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 1
    object Bevel1: TBevel
      Left = 0
      Top = 456
      Width = 691
      Height = 3
      Align = alBottom
      Shape = bsBottomLine
    end
    object RankPage: TcxPageControl
      Left = 0
      Top = 50
      Width = 641
      Height = 406
      Align = alClient
      LookAndFeel.Kind = lfStandard
      TabOrder = 0
      OnChange = RankPageChange
      OnDrawTabEx = RankPageDrawTabEx
      ClientRectBottom = 404
      ClientRectLeft = 2
      ClientRectRight = 639
      ClientRectTop = 2
    end
    object Panel3: TPanel
      Left = 641
      Top = 50
      Width = 50
      Height = 406
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 1
      object btnDelRnak: TBitBtn
        Left = 5
        Top = 191
        Width = 38
        Height = 60
        Hint = #21034#38500#22810#38542#26041#26696#38913#31844
        Caption = #21034#38500
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = btnDelRnakClick
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0D1402058CCB03099AE10000507B000004070000000000000000000000000000
          0000000057850000548B00001C2A000000000000000000000000000000000000
          5D8D1844F6FF194DF8FF1031D2FF01028BD000001A2400000000000000010000
          66930928D7FF092ED7FF0313B3FF0000608F0000000000000000000000000102
          92D02451F9FF1F52FFFF1D4FFFFF1744E8FF040AA3EC0000283A000066970D2E
          DDFF1142F9FF0D3DF5FF0B3BF0FF041ABCFF00003C5D00000000000000000000
          2D421832DBFF285BFFFF2456FFFF2253FFFF1B4BF1FF050DAEFB0F30DDFF164A
          FEFF1344F9FF1041F6FF0E3EF6FF0A3CF0FF000077BF00000000000000000000
          3C511F37DDFF3A6FFFFF2C5EFFFF295AFFFF2657FFFF2052FCFF1C4FFFFF194A
          FDFF1646FAFF1445FAFF0F3DF2FF0209A0E700002C4300000000000000000000
          000100005066111BBCF03D6AFBFF3567FFFF2C5DFFFF2859FFFF2253FFFF1D4E
          FFFF1A4DFFFF123DEDFF000289CB0000131A0000000000000000000000000000
          0000000000000000161C000085BA2E4EE7FF3668FFFF2E5EFFFF2859FFFF2254
          FFFF163DEAFF00007DBE0000070B000000000000000000000000000000000000
          0000000000000000000000003A4D253FDFFF3B6DFFFF3464FFFF2E5EFFFF2759
          FFFF1B46EAFF00018CCF00000508000000000000000000000000000000000000
          00000000000000000000020297C14B7CFFFF4170FFFF3B6BFFFF396CFFFF2D5E
          FFFF2558FFFF1336D7FF00006993000000000000000000000000000000000000
          0000000000000000252C263CDAFE5080FFFF4575FFFF3662FAFF0D14C2FE3C6D
          FFFF2A5BFFFF2053FDFF0B1DC2FD000035470000000000000000000000000000
          00000000000000006F8B527CFAFF5081FFFF4B7DFFFF0A11B8EA000030420D13
          B6E9386AFFFF2456FFFF1A4AF2FF02069FE300000B0F00000000000000000000
          0000000000001119C2E06A9CFFFF5788FFFF2B46E7FF00004A5C000000000000
          28320E1AC1F23065FFFF1F51FFFF1439DDFF00006C9C00000000000000000000
          00000000000000007E913A52E3FE5782FBFF01019EC200000000000000000000
          0000000039471425D2FA265AFFFF0F2EE3FF010295CF00000000000000000000
          000000000000000000000000313C00007BA40000273000000000000000000000
          00000000000000004C63000188BC000034490000010200000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
        Layout = blGlyphTop
      end
      object btnAddRank: TBitBtn
        Left = 5
        Top = 120
        Width = 38
        Height = 60
        Hint = #26032#22686#22810#38542#26041#26696#38913#31844
        Caption = #26032#22686
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = btnAddRankClick
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000026A06CE0E8E18FF0B8A15FF0A8814FF0985
          12FF005D02C70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000036C07CE51DA7BFF3ACF69FF39CD67FF32C2
          5BFF015F03C70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000046D08CE5CE084FF3ED46EFF3DD36DFF34C5
          5FFF026104C70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000056F0ACE64E48AFF41D771FF3FD56FFF37C8
          61FF026205C7000000000000000000000000000000000000000000000000077A
          0DD407790ED406780DD406760CD408860FF056E282FF44DA74FF41D872FF39CB
          64FF037808EC026606CE026405CE026304CE016103CE000000000000000021A7
          2DFF60E487FF4BDD79FF4ADC78FF49DC77FF4AE07AFF46DD77FF44DB75FF40D6
          70FF3BCD67FF39CB65FF36C862FF33C45EFF0A8413FF000000000000000022A9
          2FFF76F099FF5EEA8AFF5AE888FF56E684FF53E481FF4EE17DFF47DE78FF44DB
          75FF41D872FF3FD56FFF3DD36DFF39CD67FF0B8715FF000000000000000024AB
          30FF7EF39FFF68EE91FF64ED8EFF60EA8BFF59E886FF54E482FF4EE17DFF46DD
          77FF44DA74FF41D771FF3ED46EFF3ACF69FF0C8916FF000000000000000025AD
          32FF91F7ABFF8DF6A8FF8BF5A6FF89F4A5FF7AF09BFF59E886FF53E481FF4ADF
          7AFF5BE486FF67E58CFF5EE186FF53DB7CFF0F8D1AFF0000000000000000077F
          10CE087D10CE077C0FCE077A0FCE0A8D14EF89F4A5FF60EA8BFF56E684FF45D8
          72FF06820EEC056F0BCE056C0ACE046B09CE036907CE00000000000000000000
          0000000000000000000000000000077A0FCE8BF5A6FF64ED8EFF5AE888FF45D7
          71FF056C0AC70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000087C10CE8DF6A8FF68EE91FF5EEA8AFF46D8
          72FF056F0BC70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000097C11CE91F7ABFF7FF39FFF76F099FF58DF
          7FFF05700CC70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000087E12CE26AD33FF25AB32FF23A930FF21A6
          2EFF05710CC70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
        Layout = blGlyphTop
      end
    end
    object Panel4: TPanel
      Left = 0
      Top = 0
      Width = 691
      Height = 50
      Align = alTop
      BevelOuter = bvNone
      TabOrder = 2
      object Label30: TLabel
        Tag = 1
        Left = 60
        Top = 6
        Width = 48
        Height = 13
        Caption = #26041#26696#20195#30908
      end
      object Label18: TLabel
        Left = 60
        Top = 30
        Width = 48
        Height = 14
        Caption = #40670#25976#36774#27861
        Font.Charset = ANSI_CHARSET
        Font.Color = clRed
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
      end
      object edtStepNo: TcxTextEdit
        Tag = 1
        Left = 115
        Top = 1
        Enabled = False
        ParentFont = False
        Properties.MaxLength = 10
        Style.StyleController = StyleController.EditorStyle
        TabOrder = 0
        Width = 77
      end
      object edtDescription: TcxTextEdit
        Tag = 1
        Left = 194
        Top = 1
        ParentFont = False
        Properties.MaxLength = 50
        Properties.OnChange = edtDescriptionPropertiesChange
        Style.StyleController = StyleController.EditorStyle
        TabOrder = 1
        Width = 343
      end
      object chkMasterSale: TcxCheckBox
        Tag = 1
        Left = 546
        Top = 2
        Caption = #20027#25512#22810#38542
        Properties.NullStyle = nssUnchecked
        Properties.OnChange = chkMasterSalePropertiesChange
        Style.StyleController = StyleController.CheckBoxStyle
        TabOrder = 2
        Width = 74
      end
      inline LinkKey: TDataLookup
        Left = 112
        Top = 27
        Width = 424
        Height = 23
        HorzScrollBar.Visible = False
        VertScrollBar.Visible = False
        TabOrder = 3
        inherited CodeNo: TcxTextEdit
          Left = 3
          Properties.OnChange = LinkKeyCodeNoPropertiesChange
          Width = 77
        end
        inherited CodeName: TcxLookupComboBox
          Left = 82
          Properties.OnChange = LinkKeyCodeNamePropertiesChange
          Properties.OnInitPopup = LinkKeyCodeNamePropertiesInitPopup
          Width = 342
        end
      end
    end
    object RankGrid: TcxGrid
      Left = 16
      Top = 88
      Width = 608
      Height = 275
      TabOrder = 3
      LookAndFeel.Kind = lfFlat
      object gvRank: TcxGridDBBandedTableView
        NavigatorButtons.ConfirmDelete = False
        OnCustomDrawCell = gvRankCustomDrawCell
        OnEditing = gvRankEditing
        DataController.DataSource = RankDataSource
        DataController.Summary.DefaultGroupSummaryItems = <>
        DataController.Summary.FooterSummaryItems = <>
        DataController.Summary.SummaryGroups = <>
        OptionsBehavior.GoToNextCellOnEnter = True
        OptionsCustomize.ColumnFiltering = False
        OptionsCustomize.ColumnMoving = False
        OptionsCustomize.ColumnSorting = False
        OptionsCustomize.ColumnVertSizing = False
        OptionsData.Deleting = False
        OptionsData.DeletingConfirmation = False
        OptionsData.Inserting = False
        OptionsView.DataRowHeight = 25
        OptionsView.GroupByBox = False
        OptionsView.HeaderAutoHeight = True
        OptionsView.HeaderEndEllipsis = True
        OptionsView.RowSeparatorColor = clBtnFace
        OptionsView.RowSeparatorWidth = 2
        OptionsView.BandHeaders = False
        Styles.Background = StyleController.cxStyle1
        Bands = <
          item
            Width = 604
          end>
        object gvRankColumn1: TcxGridDBBandedColumn
          DataBinding.FieldName = 'RankLvName'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.ReadOnly = True
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Options.Focusing = False
          Width = 50
          Position.BandIndex = 0
          Position.ColIndex = 0
          Position.RowIndex = 0
          IsCaptionAssigned = True
        end
        object gvRankColumn2: TcxGridDBBandedColumn
          Caption = #20351#29992#26376#25976
          DataBinding.FieldName = 'Mon'
          PropertiesClassName = 'TcxMaskEditProperties'
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '(\d)|(\d\d)|(\d\d\d)'
          Properties.OnChange = gvRankColumn2PropertiesChange
          Properties.OnEditValueChanged = gvRankColumn2PropertiesEditValueChanged
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Width = 37
          Position.BandIndex = 0
          Position.ColIndex = 1
          Position.RowIndex = 0
        end
        object gvRankColumn3: TcxGridDBBandedColumn
          DataBinding.FieldName = 'LblPeriod'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.ReadOnly = True
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Options.Focusing = False
          Width = 64
          Position.BandIndex = 0
          Position.ColIndex = 2
          Position.RowIndex = 0
          IsCaptionAssigned = True
        end
        object gvRankColumn4: TcxGridDBBandedColumn
          Caption = #32371#36027#26399#25976
          DataBinding.FieldName = 'Period'
          PropertiesClassName = 'TcxMaskEditProperties'
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '(\d)|(\d\d)|(\d\d\d)'
          Properties.OnEditValueChanged = gvRankColumn4PropertiesEditValueChanged
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Width = 55
          Position.BandIndex = 0
          Position.ColIndex = 3
          Position.RowIndex = 0
        end
        object gvRankColumn5: TcxGridDBBandedColumn
          Caption = #36027#29575#20381#25818
          DataBinding.FieldName = 'RateType'
          PropertiesClassName = 'TcxImageComboBoxProperties'
          Properties.Items = <
            item
              ImageIndex = 0
            end
            item
              Description = '1.'#20381#36027#29575#34920
              Value = 1
            end
            item
              Description = '2.'#20381#20778#24800#37329#38989
              Value = 2
            end>
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Width = 106
          Position.BandIndex = 0
          Position.ColIndex = 4
          Position.RowIndex = 0
        end
        object gvRankColumn6: TcxGridDBBandedColumn
          Caption = #20778#24800#37329#38989
          DataBinding.FieldName = 'DiscountAmt'
          PropertiesClassName = 'TcxCurrencyEditProperties'
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-,0'
          Properties.MaxValue = 99999999.000000000000000000
          Properties.OnEditValueChanged = gvRankColumn4PropertiesEditValueChanged
          Properties.OnValidate = gvRankColumn6PropertiesValidate
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Width = 72
          Position.BandIndex = 0
          Position.ColIndex = 6
          Position.RowIndex = 0
        end
        object gvRankColumn7: TcxGridDBBandedColumn
          Caption = #21934#26376#37329#38989
          DataBinding.FieldName = 'MonthAmt'
          PropertiesClassName = 'TcxCurrencyEditProperties'
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 99999999.999000000000000000
          Properties.OnValidate = gvRankColumn6PropertiesValidate
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Width = 71
          Position.BandIndex = 0
          Position.ColIndex = 7
          Position.RowIndex = 0
        end
        object gvRankColumn8: TcxGridDBBandedColumn
          Caption = #21934#26085#37329#38989
          DataBinding.FieldName = 'DayAmt'
          PropertiesClassName = 'TcxCurrencyEditProperties'
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000;-,0.000'
          Properties.MaxValue = 9999.999000000000000000
          Properties.OnValidate = gvRankColumn6PropertiesValidate
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Width = 72
          Position.BandIndex = 0
          Position.ColIndex = 8
          Position.RowIndex = 0
        end
        object gvRankColumn10: TcxGridDBBandedColumn
          DataBinding.FieldName = 'RankLv'
          Visible = False
          Position.BandIndex = 0
          Position.ColIndex = 9
          Position.RowIndex = 0
        end
        object gvRankColumn11: TcxGridDBBandedColumn
          Caption = #20778#24800#35498#26126
          DataBinding.FieldName = 'Comment'
          PropertiesClassName = 'TcxButtonEditProperties'
          Properties.Buttons = <
            item
              Default = True
              Kind = bkEllipsis
            end>
          Properties.OnButtonClick = gvRankColumn11PropertiesButtonClick
          OnGetCellHint = gvRankColumn11GetCellHint
          GroupSummaryAlignment = taCenter
          HeaderAlignmentHorz = taCenter
          Options.ShowEditButtons = isebAlways
          Width = 38
          Position.BandIndex = 0
          Position.ColIndex = 10
          Position.RowIndex = 0
        end
        object gvRankColumn12: TcxGridDBBandedColumn
          Caption = #25240#25187#29575
          DataBinding.FieldName = 'DiscountRate'
          PropertiesClassName = 'TcxCurrencyEditProperties'
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = '0%'
          Properties.MaxValue = 100.000000000000000000
          Properties.OnEditValueChanged = gvRankColumn4PropertiesEditValueChanged
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Position.BandIndex = 0
          Position.ColIndex = 5
          Position.RowIndex = 0
        end
      end
      object glRank: TcxGridLevel
        GridView = gvRank
      end
    end
  end
  object ActionList1: TActionList
    Left = 20
    Top = 140
    object actSave: TAction
      Caption = 'F2.'#30906#23450
      ShortCut = 113
      OnExecute = actSaveExecute
    end
    object actCancel: TAction
      Caption = #21462#28040'(&X)'
      ShortCut = 32856
      OnExecute = actCancelExecute
    end
  end
  object RankDataSource: TDataSource
    DataSet = RankDataSet
    Left = 52
    Top = 140
  end
  object RankDataSet: TClientDataSet
    Aggregates = <>
    FieldDefs = <>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    AfterCancel = RankDataSetAfterCancel
    Left = 52
    Top = 174
  end
end
