object frmInvB06_1: TfrmInvB06_1
  Left = 351
  Top = 339
  Width = 683
  Height = 371
  ActiveControl = mskInvUseYear
  Caption = '[B06] '#26597#35426#26781#20214
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object pnlFun1: TPanel
    Left = 0
    Top = 0
    Width = 667
    Height = 111
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    object Bevel1: TBevel
      Left = 0
      Top = 104
      Width = 667
      Height = 7
      Align = alBottom
      Shape = bsTopLine
    end
    object btnQuery: TcxButton
      Left = 14
      Top = 12
      Width = 85
      Height = 29
      Caption = #26597#35426
      TabOrder = 0
      OnClick = btnQueryClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        333333333333333FF3FF3333333333CC30003333333333773777333333333C33
        3000333FF33337F33777339933333C3333333377F33337F3333F339933333C33
        33003377333337F33377333333333C333300333F333337F33377339333333C33
        3333337FF3333733333F33993333C33333003377FF33733333773339933C3333
        330033377FF73F33337733339933C33333333FF377F373F3333F993399333C33
        330077F377F337F33377993399333C33330077FF773337F33377399993333C33
        33333777733337F333FF333333333C33300033333333373FF7773333333333CC
        3000333333333377377733333333333333333333333333333333}
      NumGlyphs = 2
    end
    object btnExit: TcxButton
      Left = 576
      Top = 12
      Width = 84
      Height = 29
      Cancel = True
      Caption = #32080#26463
      TabOrder = 4
      OnClick = btnExitClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        333333333333333333333333333333333333333FFF33FF333FFF339993370733
        999333777FF37FF377733339993000399933333777F777F77733333399970799
        93333333777F7377733333333999399933333333377737773333333333990993
        3333333333737F73333333333331013333333333333777FF3333333333910193
        333333333337773FF3333333399000993333333337377737FF33333399900099
        93333333773777377FF333399930003999333337773777F777FF339993370733
        9993337773337333777333333333333333333333333333333333333333333333
        3333333333333333333333333333333333333333333333333333}
      NumGlyphs = 2
    end
    object btnCreate: TcxButton
      Left = 190
      Top = 12
      Width = 100
      Height = 29
      Caption = #26032#22686#20013#29518#30332#31080
      TabOrder = 2
      OnClick = btnCreateClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        33333333FF33333333FF333993333333300033377F3333333777333993333333
        300033F77FFF3333377739999993333333333777777F3333333F399999933333
        33003777777333333377333993333333330033377F3333333377333993333333
        3333333773333333333F333333333333330033333333F33333773333333C3333
        330033333337FF3333773333333CC333333333FFFFF77FFF3FF33CCCCCCCCCC3
        993337777777777F77F33CCCCCCCCCC3993337777777777377333333333CC333
        333333333337733333FF3333333C333330003333333733333777333333333333
        3000333333333333377733333333333333333333333333333333}
      NumGlyphs = 2
    end
    object BtnCheck: TcxButton
      Left = 102
      Top = 12
      Width = 85
      Height = 29
      Caption = #23565#29518
      TabOrder = 1
      OnClick = BtnCheckClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000120B0000120B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00555555555555
        55555FFFFFFF5F55FFF5777777757559995777777775755777F7555555555550
        305555555555FF57F7F555555550055BB0555555555775F777F55555550FB000
        005555555575577777F5555550FB0BF0F05555555755755757F555550FBFBF0F
        B05555557F55557557F555550BFBF0FB005555557F55575577F555500FBFBFB0
        B05555577F555557F7F5550E0BFBFB00B055557575F55577F7F550EEE0BFB0B0
        B05557FF575F5757F7F5000EEE0BFBF0B055777FF575FFF7F7F50000EEE00000
        B0557777FF577777F7F500000E055550805577777F7555575755500000555555
        05555777775555557F5555000555555505555577755555557555}
      NumGlyphs = 2
    end
    object rgpQueryCond: TcxRadioGroup
      Left = 14
      Top = 46
      Caption = ' '#26597#35426#26781#20214' '
      Properties.Items = <
        item
          Caption = #30332#31080#20351#29992#24180
          Value = 0
        end>
      ItemIndex = 0
      TabOrder = 5
      Height = 45
      Width = 225
    end
    object mskInvUseYear: TcxMaskEdit
      Left = 100
      Top = 64
      ParentFont = False
      Properties.MaskKind = emkRegExprEx
      Properties.EditMask = '([123][0-9])? [0-9][0-9]'
      Properties.MaxLength = 0
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -12
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.StyleController = dtmMain.cxEditStyle
      Style.IsFontAssigned = True
      TabOrder = 6
      Width = 57
    end
    object cmbInvUseMonth: TcxComboBox
      Left = 158
      Top = 64
      Properties.Items.Strings = (
        '01'
        '03'
        '05'
        '07'
        '09'
        '11')
      Properties.OnEditValueChanged = cmbInvUseMonthPropertiesEditValueChanged
      TabOrder = 7
      Text = '01'
      Width = 45
    end
    object cxLabel1: TcxLabel
      Left = 208
      Top = 66
      Caption = #26376
    end
    object btnPrint: TcxButton
      Left = 294
      Top = 12
      Width = 84
      Height = 29
      Caption = #21015#21360
      TabOrder = 3
      OnClick = btnPrintClick
      Glyph.Data = {
        36060000424D3606000000000000360000002800000020000000100000000100
        1800000000000006000000000000000000000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        C9C6C28068508060508060507060507058407058407050407050406048306048
        30604830A29A92FF00FFFF00FFFF00FFC9C6C280685080605080605070605070
        5840705840705040705040604830604830604830A29A92FF00FFFF00FFC1C4C3
        A38D85E0D0C0B0A090B0A090B0A090B0A090B0A090B0A090B0A090B0A090B0A0
        90B0A090604830FF00FFFF00FFC1C4C3A38D85FF00FFB0A090B0A090B0A090B0
        A090B0A090B0A090B0A090B0A090B0A090B0A090604830FF00FFFF00FFBCB7B0
        B29C94FFE8E0FFF8F0FFF0E0FFE8E0F0D8D0F0D0B0F0C0A000A00000A00000A0
        00705840604830857767FF00FFBCB7B0B29C94FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FF705840FF00FF857767FF00FFB29485
        E0D8D0FFFFFFFFFFFFFFFFFFFFFFFFFFF0F0F0E0E0F0D8C000FF1000FFB000A0
        00806850705040604830FF00FFB29485E0D8D0FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FF806850FF00FF604830FF00FFB29485
        F0E8E0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF8F0F0E8E000FF1000FF1000A0
        00907060706050604830FF00FFB29485F0E8E0FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FF907060FF00FF604830FF00FFB09880
        D0C0B0D0C0B0C0B0A0B29C94B09880A088809080709070608068608060507058
        50908070806860705840FF00FFB09880D0C0B0D0C0B0C0B0A0B29C94B09880A0
        8880908070907060806860806050705850908070FF00FF705840FF00FFC0A090
        FFF8F0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF8F0FFF0F0F0F0F0F0E8
        E0A38D85907860806050FF00FFC0A090FF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFA38D85FF00FF806050FF00FFCEC9C3
        B6A18CD0B0A0C0A8A0D0B0A0C0A090B29485A080709070608060507060508070
        60B0A090A08870806050FF00FFCEC9C3B6A18CD0B0A0C0A8A0D0B0A0C0A090B2
        9485A08070907060806050706050807060B0A090A08870806050FF00FFFF00FF
        C9C6C2C0B0A0E0C8C0FFFFFFFFF8FFFFF8FFFFF0F0F0F0E0F0E0E0C0B0A08060
        50A08880C0B0A0806050FF00FFFF00FFC9C6C2C0B0A0E0C8C0FF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFC0B0A0806050A08880C0B0A0806050FF00FFFF00FF
        FF00FFDEDFDDC0B0A0FFFFFFF0E8E0D0C8C0D0C8C0D0B8B0D0B8B0E0D0D08068
        60806050B29C94B0A090FF00FFFF00FFFF00FFDEDFDDC0B0A0FF00FFF0E8E0D0
        C8C0D0C8C0D0B8B0D0B8B0E0D0D0806860806050B29C94B0A090FF00FFFF00FF
        FF00FFFF00FFD8CBBCF0E8E0FFFFFFFFFFFFFFFFFFFFFFFFFFF0F0F0E0D0D0B8
        B0806050BAADA8FF00FFFF00FFFF00FFFF00FFFF00FFD8CBBCFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFD0B8B0FF00FFBAADA8FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFD0C0B0FFFFFFF0F0F0D0C8C0D0C8C0D0B8B0C0B0B0E0D8
        D0857767806050FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFD0C0B0FF00FFFF
        00FFD0C8C0D0C8C0D0B8B0C0B0B0E0D8D0857767806050FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFE0D0D0E0D0D0FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF0E8
        E0D0B8B0806050FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFE0D0D0E0D0D0FF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFD0B8B0806050FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFDED5D3D0C0B0D0B8B0D0C0B0D0C0B0D0B8B0D0C0
        B0D0B8B0D0C0B0FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFDED5D3D0
        C0B0D0B8B0D0C0B0D0C0B0D0B8B0D0C0B0D0B8B0D0C0B0FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      NumGlyphs = 2
    end
    object btnVerify: TcxButton
      Left = 383
      Top = 12
      Width = 84
      Height = 29
      Cancel = True
      Caption = #23529#26680
      TabOrder = 9
      OnClick = btnVerifyClick
      Glyph.Data = {
        36040000424D3604000000000000360000002800000010000000100000000100
        2000000000000004000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000004040410000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        00000000000000000000000000003F3F3FFF3F3F3FFF00000000040404100000
        0000000000000000000000000000000000000000000000000000000000000000
        000000000000000000003F3F3FFF70E584FF70E584FF3F3F3FFF000000000404
        0410000000000000000000000000000000000000000000000000000000000000
        0000000000003F3F3FFF70E584FF70E584FF70E584FF70E584FF3F3F3FFF0000
        0000040404100000000000000000000000000000000000000000000000000000
        00003F3F3FFF70E584FF70E584FF70E584FF70E584FF70E584FF70E584FF3F3F
        3FFF000000000000000000000000000000000000000000000000000000003F3F
        3FFF70E584FF70E584FF70E584FF3F3F3FFF424943FF70E584FF70E584FF70E5
        84FF3F3F3FFF00000000000000000000000000000000000000003F3F3FFF70E5
        84FF70E584FF70E584FF3F3F3FFF00000000000000003F3F3FFF6DDB80FF70E5
        84FF70E584FF3F3F3FFF00000000000000000000000000000000000000003F3F
        3FFF70E584FF3F3F3FFF000000000404041000000000000000003F3F3FFF70E5
        84FF70E584FF70E584FF3F3F3FFF000000000000000000000000000000000000
        00003F3F3FFF0000000004040410000000000000000000000000000000003F3F
        3FFF70E584FF70E584FF70E584FF3F3F3FFF0000000000000000000000000000
        0000000000000404041000000000000000000000000000000000000000000000
        00003F3F3FFF70E584FF70E584FF70E584FF3F3F3FFF00000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000003F3F3FFF70E584FF70E584FF70E584FF3F3F3FFF000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        000000000000000000003F3F3FFF70E584FF3F3F3FFF04040410000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        00000000000000000000000000003F3F3FFF0C0C0C3000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000000000000000
        0000000000000000000000000000000000000000000000000000}
    end
    object rgpRunCompWhere: TcxRadioGroup
      Left = 274
      Top = 46
      Caption = ' '#20844#21496#21029#36984#38917
      Properties.Columns = 2
      Properties.ImmediatePost = True
      Properties.Items = <
        item
          Caption = #20491#21029#20844#21496
          Value = True
        end
        item
          Caption = #25152#26377#20844#21496
          Value = False
        end>
      ItemIndex = 0
      TabOrder = 10
      Height = 45
      Width = 199
    end
    object cxButton1: TcxButton
      Left = 471
      Top = 12
      Width = 100
      Height = 29
      Caption = #21295#20837#20013#29518#28165#20874
      TabOrder = 11
      OnClick = cxButton1Click
      Glyph.Data = {
        36040000424D3604000000000000360000002800000010000000100000000100
        2000000000000004000000000000000000000000000000000000B07E55FFB07E
        55FFB07E55FFB07E55FFB07E55FFB07E55FFB07E55FFB07E55FFB07E55FFB07E
        55FFB07E55FFB07E55FFB07E55FFB07E55FF0000000000000000B28056FFFFFF
        FEFFFFF8F3FFFFF7F1FFFFF6EEFFFFF5EDFFFFFEFBFF8EAF82FF429B4EFFFFF6
        EBFFFFEFE1FFFFEDE0FFFFF8ECFFB07E55FF0000000000000000B48257FFFFFA
        F6FFFEF2EBFFFEF1E8FFFFF2EAFFFFF1E9FF569051FF96DBA7FF429B4EFFFFF2
        EAFFFFECDFFFFFE9DBFFFFF4E9FFB07E55FF0000000000000000B78458FFFFFB
        F7FFFEF3EDFFFFF9F6FFDCDCC9FF45924BFF9BE5AAFF9FEAADFF5FB26BFF419B
        50FF419B50FF3F994EFF3D974CFF409C50FF4D9F5AFF429B4EFFB98759FFFFFC
        F9FFFFF9F6FF93B38BFF57A663FFACF3BAFF88E293FF79D987FF7CDA8CFF77D5
        88FF6CCE7FFF62C675FF57BF6DFF4EB765FF8CDBA1FF429B4EFFBC895AFFFFFC
        F8FFFFFAF7FF93B48CFF57A662FFACF3BAFF88E395FF7BDC88FF80DD8EFF7AD9
        8AFF70D181FF65C879FF5AC16FFF51B967FF8DDCA2FF429B4EFFBE8B5BFFFFFC
        F8FFFEF5EFFFFFFCFAFFDCDFCFFF45924BFF9DE5ACFFA3EEB1FF61B26DFF449B
        50FF439B51FF429B4EFF3F994EFF429D50FF4D9E58FF429B4EFFC08D5CFFFFFB
        F8FFFEF5EFFFFEF6F0FFFFF8F4FFFFF8F5FF579354FF97DBA7FF429B4EFFFFFB
        FAFFFFF5EDFFFFF2EAFFFFFAF4FFB07E55FF0000000000000000C3905EFFFFFB
        F8FFFEF5EFFFFEF6F0FFFEF6F0FFFFF6F1FFFFFFFFFF8AAD82FF429B4EFFFFF7
        F1FFFEEFE7FFFEEEE3FFFFF4ECFFB07E55FF0000000000000000C5925FFFFFFB
        F8FFFEF5EFFFFEF6F0FFFEF6F0FFFEF6F0FFFEF6F0FFFFFAF6FFFFF8F3FFFEF3
        ECFFFEF2E8FFFEEFE6FFFFF7EEFFB07E55FF0000000000000000C89460FFFFFB
        F8FFFEF5EFFFFEF6F0FFFEF6F0FFFEF6F0FFFEF6F0FFFEF6F0FFFEF5EFFFFEF4
        EEFFFEF3ECFFFEF1E8FFFFF7F2FFB07E55FF0000000000000000CA9661FFFFFC
        FBFFFEF7F2FFFEF7F3FFFEF8F3FFFEF8F4FFFFF8F4FFFFF8F5FFFFF9F5FFFFF8
        F4FFFFF7F3FFFFF5F0FFFFFBF9FFB07E55FF0000000000000000CC9962FFFFFC
        FBFFF7E5CAFFF6E3C7FFF5DDC2FFF4DABDFFF1D5B7FFF2D2B3FFF0D0B1FFF0D1
        B2FFF0CFB1FFF0D0AFFFFFFCFBFFB07E55FF0000000000000000CF9B63FFFFFC
        FBFFEFD4A3FFEDCE9AFFEBC590FFE9BB85FFE5B37CFFE5AC73FFE3A86EFFE4A8
        6EFFE4A96EFFE4A96FFFFFFCFBFFB07E55FF0000000000000000D19D64FFFFFC
        FBFFFFFCFBFFFFFCFBFFFFFCFBFFFFFCFBFFFFFCFBFFFFFCFBFFFFFCFBFFFFFC
        FBFFFFFCFBFFFFFCFBFFFFFCFBFFB07E55FF0000000000000000D4A168FFD19D
        64FFCE9A63FFCB9862FFC89560FFC6925FFFC3905EFFC08D5CFFBD8B5BFFBB88
        5AFFB88558FFB58357FFB28056FFB3835CFF0000000000000000}
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 111
    Width = 667
    Height = 195
    Align = alTop
    TabOrder = 1
    object gvGrid: TcxGrid
      Left = 1
      Top = 1
      Width = 665
      Height = 193
      Align = alClient
      TabOrder = 0
      object gvView: TcxGridDBBandedTableView
        OnDblClick = gvViewDblClick
        NavigatorButtons.ConfirmDelete = False
        DataController.DataSource = dsInv035
        DataController.Summary.DefaultGroupSummaryItems = <>
        DataController.Summary.FooterSummaryItems = <>
        DataController.Summary.SummaryGroups = <>
        OptionsCustomize.ColumnFiltering = False
        OptionsCustomize.ColumnSorting = False
        OptionsData.Deleting = False
        OptionsData.DeletingConfirmation = False
        OptionsData.Editing = False
        OptionsData.Inserting = False
        OptionsSelection.InvertSelect = False
        OptionsView.CellEndEllipsis = True
        OptionsView.GroupByBox = False
        OptionsView.HeaderEndEllipsis = True
        OptionsView.Indicator = True
        OptionsView.RowSeparatorWidth = 5
        Styles.Background = dtmMain.cxGridBackGroundStyle
        Styles.Content = dtmMain.cxGridBackGroundStyle
        Styles.Inactive = dtmMain.cxGridInActiveStyle
        Styles.Selection = dtmMain.cxGridActiveStyle
        Bands = <
          item
            Width = 2334
          end>
        object gvYearMonth: TcxGridDBBandedColumn
          Caption = #23565#29518#20351#29992#24180#26376
          DataBinding.FieldName = 'YEARMONTH'
          PropertiesClassName = 'TcxTextEditProperties'
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Width = 87
          Position.BandIndex = 0
          Position.ColIndex = 0
          Position.RowIndex = 0
        end
        object gvExecuteDate: TcxGridDBBandedColumn
          Caption = #22519#34892#23565#29518#26085
          DataBinding.FieldName = 'EXECUTEDATE'
          PropertiesClassName = 'TcxDateEditProperties'
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Width = 150
          Position.BandIndex = 0
          Position.ColIndex = 1
          Position.RowIndex = 0
        end
        object gvSPECIALPRIZE1: TcxGridDBBandedColumn
          Caption = #29305#29518#34399#30908'1'
          DataBinding.FieldName = 'SPECIALPRIZE1'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Width = 88
          Position.BandIndex = 0
          Position.ColIndex = 7
          Position.RowIndex = 0
        end
        object gvSPECIALPRIZE2: TcxGridDBBandedColumn
          Caption = #29305#29518#34399#30908'2'
          DataBinding.FieldName = 'SPECIALPRIZE2'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Width = 86
          Position.BandIndex = 0
          Position.ColIndex = 8
          Position.RowIndex = 0
        end
        object gvSPECIALPRIZE3: TcxGridDBBandedColumn
          Caption = #29305#29518#34399#30908'3'
          DataBinding.FieldName = 'SPECIALPRIZE3'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Width = 90
          Position.BandIndex = 0
          Position.ColIndex = 9
          Position.RowIndex = 0
        end
        object gvSPECIALPRIZE4: TcxGridDBBandedColumn
          Caption = #29305#29518#34399#30908'4'
          DataBinding.FieldName = 'SPECIALPRIZE4'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Options.GroupFooters = False
          Width = 77
          Position.BandIndex = 0
          Position.ColIndex = 10
          Position.RowIndex = 0
        end
        object gvFIRSTPRIZE1: TcxGridDBBandedColumn
          Caption = #38957#29518#34399#30908'1'
          DataBinding.FieldName = 'FIRSTPRIZE1'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Width = 91
          Position.BandIndex = 0
          Position.ColIndex = 12
          Position.RowIndex = 0
        end
        object gvFIRSTPRIZE2: TcxGridDBBandedColumn
          Caption = #38957#29518#34399#30908'2'
          DataBinding.FieldName = 'FIRSTPRIZE2'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Focusing = False
          Width = 82
          Position.BandIndex = 0
          Position.ColIndex = 13
          Position.RowIndex = 0
        end
        object gvFIRSTPRIZE3: TcxGridDBBandedColumn
          Caption = #38957#29518#34399#30908'3'
          DataBinding.FieldName = 'FIRSTPRIZE3'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Focusing = False
          Width = 93
          Position.BandIndex = 0
          Position.ColIndex = 14
          Position.RowIndex = 0
        end
        object gvEXTRAPRIZE1: TcxGridDBBandedColumn
          Caption = #21152#38283#34399#30908'1'
          DataBinding.FieldName = 'EXTRAPRIZE1'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Width = 94
          Position.BandIndex = 0
          Position.ColIndex = 17
          Position.RowIndex = 0
        end
        object gvViewUPTTIME: TcxGridDBBandedColumn
          Caption = #30064#21205#26085#26399
          DataBinding.FieldName = 'UPTTIME'
          Options.Focusing = False
          Width = 99
          Position.BandIndex = 0
          Position.ColIndex = 22
          Position.RowIndex = 0
        end
        object gvViewUPTEN: TcxGridDBBandedColumn
          Caption = #30064#21205#20154#21729
          DataBinding.FieldName = 'UPTEN'
          Options.Focusing = False
          Width = 72
          Position.BandIndex = 0
          Position.ColIndex = 23
          Position.RowIndex = 0
        end
        object gvEXTRAPRIZE2: TcxGridDBBandedColumn
          Caption = #21152#38283#34399#30908'2'
          DataBinding.FieldName = 'EXTRAPRIZE2'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Focusing = False
          Width = 92
          Position.BandIndex = 0
          Position.ColIndex = 18
          Position.RowIndex = 0
        end
        object gvVerifyEn: TcxGridDBBandedColumn
          Caption = #23529#26680#20154#21729
          DataBinding.FieldName = 'VerifyEn'
          Options.Focusing = False
          Width = 68
          Position.BandIndex = 0
          Position.ColIndex = 24
          Position.RowIndex = 0
        end
        object gvVerifyTime: TcxGridDBBandedColumn
          Caption = #23529#26680#26085#26399
          DataBinding.FieldName = 'VerifyTime'
          Options.Editing = False
          Options.Focusing = False
          Width = 117
          Position.BandIndex = 0
          Position.ColIndex = 25
          Position.RowIndex = 0
        end
        object gvSPECIALPRIZE5: TcxGridDBBandedColumn
          Caption = #29305#29518#34399#30908'5'
          DataBinding.FieldName = 'SPECIALPRIZE5'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Width = 88
          Position.BandIndex = 0
          Position.ColIndex = 11
          Position.RowIndex = 0
        end
        object gvFIRSTPRIZE4: TcxGridDBBandedColumn
          Caption = #38957#29518#34399#30908'4'
          DataBinding.FieldName = 'FIRSTPRIZE4'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Width = 93
          Position.BandIndex = 0
          Position.ColIndex = 15
          Position.RowIndex = 0
        end
        object gvFIRSTPRIZE5: TcxGridDBBandedColumn
          Caption = #38957#29518#34399#30908'5'
          DataBinding.FieldName = 'FIRSTPRIZE5'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Options.Sorting = False
          Width = 102
          Position.BandIndex = 0
          Position.ColIndex = 16
          Position.RowIndex = 0
        end
        object gvEXTRAPRIZE3: TcxGridDBBandedColumn
          Caption = #21152#38283#34399#30908'3'
          DataBinding.FieldName = 'EXTRAPRIZE3'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Options.Sorting = False
          Width = 79
          Position.BandIndex = 0
          Position.ColIndex = 19
          Position.RowIndex = 0
        end
        object gvEXTRAPRIZE4: TcxGridDBBandedColumn
          Caption = #21152#38283#34399#30908'4'
          DataBinding.FieldName = 'EXTRAPRIZE4'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Options.Sorting = False
          Width = 97
          Position.BandIndex = 0
          Position.ColIndex = 20
          Position.RowIndex = 0
        end
        object gvEXTRAPRIZE5: TcxGridDBBandedColumn
          Caption = #21152#38283#34399#30908'5'
          DataBinding.FieldName = 'EXTRAPRIZE5'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Options.Sorting = False
          Width = 70
          Position.BandIndex = 0
          Position.ColIndex = 21
          Position.RowIndex = 0
        end
        object gvUNUSUALPRIZE1: TcxGridDBBandedColumn
          Caption = #29305#21029#29518#34399#30908'1'
          DataBinding.FieldName = 'UNUSUALPRIZE1'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Width = 88
          Position.BandIndex = 0
          Position.ColIndex = 2
          Position.RowIndex = 0
        end
        object gvUNUSUALPRIZE2: TcxGridDBBandedColumn
          Caption = #29305#21029#29518#34399#30908'2'
          DataBinding.FieldName = 'UNUSUALPRIZE2'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Width = 90
          Position.BandIndex = 0
          Position.ColIndex = 3
          Position.RowIndex = 0
        end
        object gvUNUSUALPRIZE3: TcxGridDBBandedColumn
          Caption = #29305#21029#29518#34399#30908'3'
          DataBinding.FieldName = 'UNUSUALPRIZE3'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Width = 75
          Position.BandIndex = 0
          Position.ColIndex = 4
          Position.RowIndex = 0
        end
        object gvUNUSUALPRIZE4: TcxGridDBBandedColumn
          Caption = #29305#21029#29518#34399#30908'4'
          DataBinding.FieldName = 'UNUSUALPRIZE4'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Width = 83
          Position.BandIndex = 0
          Position.ColIndex = 5
          Position.RowIndex = 0
        end
        object gvViewColumn1: TcxGridDBBandedColumn
          Caption = #29305#21029#29518#34399#30908'5'
          DataBinding.FieldName = 'UNUSUALPRIZE5'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Options.Focusing = False
          Width = 83
          Position.BandIndex = 0
          Position.ColIndex = 6
          Position.RowIndex = 0
        end
      end
      object gvLevel1: TcxGridLevel
        GridView = gvView
      end
    end
  end
  object pnlStatus: TPanel
    Left = 0
    Top = 306
    Width = 667
    Height = 25
    Align = alTop
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlue
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
  end
  object adoINV035: TADOQuery
    CacheSize = 1000
    Connection = dtmMain.EInvConnection
    CursorType = ctStatic
    Parameters = <>
    Prepared = True
    Left = 204
    Top = 168
  end
  object dsInv035: TDataSource
    DataSet = adoINV035
    Left = 246
    Top = 167
  end
  object cdsTmp: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 282
    Top = 168
  end
  object adoCheck: TADOQuery
    CacheSize = 1000
    Connection = dtmMain.EInvConnectComm
    CursorType = ctStatic
    Parameters = <>
    Prepared = True
    Left = 206
    Top = 204
  end
end
