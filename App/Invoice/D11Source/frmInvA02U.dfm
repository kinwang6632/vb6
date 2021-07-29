object frmInvA02: TfrmInvA02
  Left = 285
  Top = 130
  Width = 792
  Height = 624
  Caption = 'frmInvA02'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  ShowHint = True
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object StatusBar1: TStatusBar
    Left = 0
    Top = 578
    Width = 784
    Height = 19
    Panels = <>
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 784
    Height = 578
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 2
    TabOrder = 1
    object pnlFun1: TPanel
      Left = 2
      Top = 2
      Width = 780
      Height = 197
      Align = alTop
      BevelOuter = bvNone
      TabOrder = 0
      DesignSize = (
        780
        197)
      object Bevel1: TBevel
        Left = 0
        Top = 186
        Width = 780
        Height = 11
        Align = alBottom
        Shape = bsTopLine
      end
      object btnColumnChooce: TSpeedButton
        Left = 752
        Top = 157
        Width = 24
        Height = 24
        Hint = #27396#20301#36984#25799
        Anchors = [akRight, akBottom]
        Flat = True
        Glyph.Data = {
          36050000424D3605000000000000360400002800000010000000100000000100
          08000000000000010000C40E0000C40E00000001000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000C0DCC000F0CA
          A6000020400000206000002080000020A0000020C0000020E000004000000040
          20000040400000406000004080000040A0000040C0000040E000006000000060
          20000060400000606000006080000060A0000060C0000060E000008000000080
          20000080400000806000008080000080A0000080C0000080E00000A0000000A0
          200000A0400000A0600000A0800000A0A00000A0C00000A0E00000C0000000C0
          200000C0400000C0600000C0800000C0A00000C0C00000C0E00000E0000000E0
          200000E0400000E0600000E0800000E0A00000E0C00000E0E000400000004000
          20004000400040006000400080004000A0004000C0004000E000402000004020
          20004020400040206000402080004020A0004020C0004020E000404000004040
          20004040400040406000404080004040A0004040C0004040E000406000004060
          20004060400040606000406080004060A0004060C0004060E000408000004080
          20004080400040806000408080004080A0004080C0004080E00040A0000040A0
          200040A0400040A0600040A0800040A0A00040A0C00040A0E00040C0000040C0
          200040C0400040C0600040C0800040C0A00040C0C00040C0E00040E0000040E0
          200040E0400040E0600040E0800040E0A00040E0C00040E0E000800000008000
          20008000400080006000800080008000A0008000C0008000E000802000008020
          20008020400080206000802080008020A0008020C0008020E000804000008040
          20008040400080406000804080008040A0008040C0008040E000806000008060
          20008060400080606000806080008060A0008060C0008060E000808000008080
          20008080400080806000808080008080A0008080C0008080E00080A0000080A0
          200080A0400080A0600080A0800080A0A00080A0C00080A0E00080C0000080C0
          200080C0400080C0600080C0800080C0A00080C0C00080C0E00080E0000080E0
          200080E0400080E0600080E0800080E0A00080E0C00080E0E000C0000000C000
          2000C0004000C0006000C0008000C000A000C000C000C000E000C0200000C020
          2000C0204000C0206000C0208000C020A000C020C000C020E000C0400000C040
          2000C0404000C0406000C0408000C040A000C040C000C040E000C0600000C060
          2000C0604000C0606000C0608000C060A000C060C000C060E000C0800000C080
          2000C0804000C0806000C0808000C080A000C080C000C080E000C0A00000C0A0
          2000C0A04000C0A06000C0A08000C0A0A000C0A0C000C0A0E000C0C00000C0C0
          2000C0C04000C0C06000C0C08000C0C0A000F0FBFF00A4A0A000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00FDFDFDFDFDFD
          FDFDFDFDFDFDFDFDFDFDFDA4A4A4A4A4A4A4A4A4A4A49B9B9BFDFDA4F6F6F6ED
          F6F6F6F6ED0909099BFDFDA4F6F6F6EDF6F6F6F6EDF609099BFDFDA4F6F6F6ED
          F6F6F6F6EDF6F6099BFDFDF7EDEDEDEDEDEDEDEDEDEDEDED9BFDFDF7F6F6F6ED
          F6F6F6F6EDF6F6F69BFDFDF7F6F6F6EDF6F6F6F6EDF6F6F69BFDFDF7FFFFF6ED
          F6F6F6F6EDF6F6F69BFDFDF7FFFFFFEDF6F6F6F6EDF6F6F69BFDFDF7EDEDEDED
          EDEDEDEDEDEDEDED9BFDFDF7FFFFFFEDFFF6F6F6EDF6F6F69BFDFDF7FFFFFFED
          FFFFF6F6EDF6F6F6A4FDFDF7FFFFFFEDFFFFF6F6EDF6F6F6A4FDFDF7F7F7A4A4
          A4A4A4A4A4A4A4A4A4FDFDFDFDFDFDFDFDFDFDFDFDFDFDFDFDFD}
        ParentShowHint = False
        ShowHint = True
        OnClick = btnColumnChooceClick
      end
      object btnQuery: TcxButton
        Left = 340
        Top = 13
        Width = 100
        Height = 28
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
        Left = 340
        Top = 149
        Width = 100
        Height = 28
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
        Left = 340
        Top = 81
        Width = 100
        Height = 28
        Caption = #38283#31435#30332#31080
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
      object btnView: TcxButton
        Left = 340
        Top = 115
        Width = 100
        Height = 28
        Caption = #39023#31034#26126#32048
        Enabled = False
        TabOrder = 3
        OnClick = btnViewClick
        Glyph.Data = {
          76010000424D7601000000000000760000002800000020000000100000000100
          04000000000000010000120B0000120B00001000000000000000000000000000
          800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333000000
          000033333377777777773333330FFFFFFFF03FF3FF7FF33F3FF700300000FF0F
          00F077F777773F737737E00BFBFB0FFFFFF07773333F7F3333F7E0BFBF000FFF
          F0F077F3337773F3F737E0FBFBFBF0F00FF077F3333FF7F77F37E0BFBF00000B
          0FF077F3337777737337E0FBFBFBFBF0FFF077F33FFFFFF73337E0BF0000000F
          FFF077FF777777733FF7000BFB00B0FF00F07773FF77373377373330000B0FFF
          FFF03337777373333FF7333330B0FFFF00003333373733FF777733330B0FF00F
          0FF03333737F37737F373330B00FFFFF0F033337F77F33337F733309030FFFFF
          00333377737FFFFF773333303300000003333337337777777333}
        NumGlyphs = 2
      end
      object BtnClear: TcxButton
        Left = 340
        Top = 47
        Width = 100
        Height = 28
        Caption = #28165#38500
        Enabled = False
        TabOrder = 1
        OnClick = BtnClearClick
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
        Left = 11
        Top = 7
        Caption = ' '#26597#35426#26781#20214' '
        ItemIndex = 0
        Properties.Items = <
          item
            Caption = #23458#32232
            Value = 0
          end
          item
            Caption = #32113#32232
            Value = 1
          end
          item
            Caption = #30332#31080#26085#26399
            Value = 2
          end
          item
            Caption = #25910#36027#26085#26399
            Value = 3
          end
          item
            Caption = #30332#31080'/'#20027#30332#31080#34399#30908
            Value = 4
          end>
        TabOrder = 5
        Height = 170
        Width = 307
      end
      object edtCustID: TcxTextEdit
        Left = 129
        Top = 29
        Properties.MaxLength = 8
        Style.StyleController = dtmMain.cxEditStyle
        TabOrder = 6
        OnEnter = InputEditEnter
        Width = 125
      end
      object edtBusinessID: TcxTextEdit
        Tag = 1
        Left = 129
        Top = 58
        Properties.MaxLength = 8
        Style.StyleController = dtmMain.cxEditStyle
        TabOrder = 7
        OnEnter = InputEditEnter
        Width = 125
      end
      object txtInvDate: TcxDateEdit
        Tag = 2
        Left = 129
        Top = 87
        Properties.DateButtons = [btnToday]
        Properties.OnValidate = txtInvDatePropertiesValidate
        Style.StyleController = dtmMain.cxEditStyle
        TabOrder = 8
        OnEnter = InputEditEnter
        OnExit = txtInvDateExit
        Width = 125
      end
      object txtChargeDate: TcxDateEdit
        Tag = 3
        Left = 129
        Top = 117
        Properties.DateButtons = [btnToday]
        Properties.OnValidate = txtInvDatePropertiesValidate
        Style.StyleController = dtmMain.cxEditStyle
        TabOrder = 9
        OnEnter = InputEditEnter
        OnExit = txtInvDateExit
        Width = 125
      end
      object edtBeginInvID: TcxTextEdit
        Tag = 4
        Left = 129
        Top = 147
        Properties.CharCase = ecUpperCase
        Properties.MaxLength = 10
        Style.StyleController = dtmMain.cxEditStyle
        TabOrder = 10
        OnEnter = InputEditEnter
        Width = 125
      end
    end
    object pnlMaster: TPanel
      Left = 2
      Top = 199
      Width = 780
      Height = 377
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 1
      object cxGrid1: TcxGrid
        Left = 0
        Top = 0
        Width = 780
        Height = 377
        Hint = #28369#40736#24038#37749#40670'2'#19979', '#36914#20837' "'#30332#31080#26126#32048#36039#26009'"'#30059#38754
        Align = alClient
        TabOrder = 0
        object cxGrid1DBBandedTableView1: TcxGridDBBandedTableView
          OnDblClick = cxGrid1DBBandedTableView1DblClick
          NavigatorButtons.ConfirmDelete = False
          DataController.DataSource = dsrInv007
          DataController.Summary.DefaultGroupSummaryItems = <>
          DataController.Summary.FooterSummaryItems = <>
          DataController.Summary.SummaryGroups = <>
          OptionsCustomize.ColumnFiltering = False
          OptionsCustomize.ColumnHiding = True
          OptionsCustomize.BandHiding = True
          OptionsData.Deleting = False
          OptionsData.DeletingConfirmation = False
          OptionsData.Editing = False
          OptionsData.Inserting = False
          OptionsSelection.CellSelect = False
          OptionsView.CellEndEllipsis = True
          OptionsView.GroupByBox = False
          OptionsView.HeaderEndEllipsis = True
          OptionsView.Indicator = True
          OptionsView.RowSeparatorWidth = 5
          OptionsView.BandHeaderEndEllipsis = True
          Styles.Background = dtmMain.cxGridBackGroundStyle
          Styles.Content = dtmMain.cxGridBackGroundStyle
          Styles.Inactive = dtmMain.cxGridInActiveStyle
          Styles.Selection = dtmMain.cxGridActiveStyle
          Styles.BandHeader = dtmMain.cxBandHeaderStyle
          Bands = <
            item
              Caption = #30332#31080#36039#26009
              Width = 546
            end
            item
              Caption = #30332#31080#37329#38989
              Width = 192
            end
            item
              Caption = #20854#23427
              Styles.Header = dtmMain.cxBandHeaderStyle
              Visible = False
              Width = 224
            end>
          object cxGrid1DBBandedTableView1INVID: TcxGridDBBandedColumn
            Caption = #30332#31080#34399#30908
            DataBinding.FieldName = 'INVID'
            Width = 85
            Position.BandIndex = 0
            Position.ColIndex = 0
            Position.RowIndex = 0
          end
          object cxGrid1DBBandedTableView1INVDATE: TcxGridDBBandedColumn
            Caption = #30332#31080#26085
            DataBinding.FieldName = 'INVDATE'
            Width = 90
            Position.BandIndex = 0
            Position.ColIndex = 2
            Position.RowIndex = 0
          end
          object cxGrid1DBBandedTableView1INVTITLE: TcxGridDBBandedColumn
            Caption = #25260#38957
            DataBinding.FieldName = 'INVTITLE'
            Width = 258
            Position.BandIndex = 0
            Position.ColIndex = 2
            Position.RowIndex = 1
          end
          object cxGrid1DBBandedTableView1BUSINESSID: TcxGridDBBandedColumn
            Caption = #32113#32232
            DataBinding.FieldName = 'BUSINESSID'
            Width = 106
            Position.BandIndex = 0
            Position.ColIndex = 4
            Position.RowIndex = 0
          end
          object cxGrid1DBBandedTableView1INVADDR: TcxGridDBBandedColumn
            Caption = #22320#22336
            DataBinding.FieldName = 'INVADDR'
            Width = 258
            Position.BandIndex = 0
            Position.ColIndex = 2
            Position.RowIndex = 2
          end
          object cxGrid1DBBandedTableView1INVFORMAT: TcxGridDBBandedColumn
            Caption = #26684#24335
            DataBinding.FieldName = 'INVFORMAT'
            HeaderAlignmentHorz = taCenter
            Width = 76
            Position.BandIndex = 0
            Position.ColIndex = 5
            Position.RowIndex = 0
          end
          object cxGrid1DBBandedTableView1CUSTID: TcxGridDBBandedColumn
            Caption = #23458#32232
            DataBinding.FieldName = 'CUSTID'
            Width = 85
            Position.BandIndex = 0
            Position.ColIndex = 0
            Position.RowIndex = 1
          end
          object cxGrid1DBBandedTableView1CUSTSNAME: TcxGridDBBandedColumn
            Caption = #21517#31281
            DataBinding.FieldName = 'CUSTSNAME'
            Width = 203
            Position.BandIndex = 0
            Position.ColIndex = 1
            Position.RowIndex = 1
          end
          object cxGrid1DBBandedTableView1CHARGEDATE: TcxGridDBBandedColumn
            Caption = #25910#36027#26085
            DataBinding.FieldName = 'CHARGEDATE'
            Width = 113
            Position.BandIndex = 0
            Position.ColIndex = 3
            Position.RowIndex = 0
          end
          object cxGrid1DBBandedTableView1ZIPCODE: TcxGridDBBandedColumn
            Caption = #37109#36958#21312#34399
            DataBinding.FieldName = 'ZIPCODE'
            Width = 85
            Position.BandIndex = 0
            Position.ColIndex = 0
            Position.RowIndex = 2
          end
          object cxGrid1DBBandedTableView1MAILADDR: TcxGridDBBandedColumn
            Caption = #37109#23492#22320#22336
            DataBinding.FieldName = 'MAILADDR'
            Width = 203
            Position.BandIndex = 0
            Position.ColIndex = 1
            Position.RowIndex = 2
          end
          object cxGrid1DBBandedTableView1TAXTYPE: TcxGridDBBandedColumn
            Caption = #31237#21029
            DataBinding.FieldName = 'TAXTYPE'
            HeaderAlignmentHorz = taCenter
            Width = 57
            Position.BandIndex = 1
            Position.ColIndex = 0
            Position.RowIndex = 0
          end
          object cxGrid1DBBandedTableView1TAXRATE: TcxGridDBBandedColumn
            Caption = #31237#29575
            DataBinding.FieldName = 'TAXRATE'
            PropertiesClassName = 'TcxCurrencyEditProperties'
            Properties.DisplayFormat = ',0.00;-$,0.00'
            HeaderAlignmentHorz = taRightJustify
            Width = 64
            Position.BandIndex = 1
            Position.ColIndex = 1
            Position.RowIndex = 0
          end
          object cxGrid1DBBandedTableView1SALEAMOUNT: TcxGridDBBandedColumn
            Caption = #37559#21806#38989
            DataBinding.FieldName = 'SALEAMOUNT'
            HeaderAlignmentHorz = taRightJustify
            Width = 71
            Position.BandIndex = 1
            Position.ColIndex = 2
            Position.RowIndex = 0
          end
          object cxGrid1DBBandedTableView1TAXAMOUNT: TcxGridDBBandedColumn
            Caption = #31237#38989
            DataBinding.FieldName = 'TAXAMOUNT'
            PropertiesClassName = 'TcxCurrencyEditProperties'
            Properties.DisplayFormat = ',0.00;-$,0.00'
            HeaderAlignmentHorz = taRightJustify
            Width = 249
            Position.BandIndex = 1
            Position.ColIndex = 0
            Position.RowIndex = 1
          end
          object cxGrid1DBBandedTableView1INVAMOUNT: TcxGridDBBandedColumn
            Caption = #30332#31080#37329#38989
            DataBinding.FieldName = 'INVAMOUNT'
            PropertiesClassName = 'TcxCurrencyEditProperties'
            Properties.DisplayFormat = ',0.00;-$,0.00'
            HeaderAlignmentHorz = taRightJustify
            Width = 94
            Position.BandIndex = 1
            Position.ColIndex = 0
            Position.RowIndex = 2
          end
          object cxGrid1DBBandedTableView1ISOBSOLETE: TcxGridDBBandedColumn
            Caption = #20316#24290#21542
            DataBinding.FieldName = 'ISOBSOLETE'
            HeaderAlignmentHorz = taCenter
            Width = 76
            Position.BandIndex = 0
            Position.ColIndex = 6
            Position.RowIndex = 0
          end
          object cxGrid1DBBandedTableView1OBSOLETEREASON: TcxGridDBBandedColumn
            Caption = #20316#24290#21407#22240
            DataBinding.FieldName = 'OBSOLETEREASON'
            Width = 204
            Position.BandIndex = 2
            Position.ColIndex = 0
            Position.RowIndex = 0
          end
          object cxGrid1DBBandedTableView1MEMO1: TcxGridDBBandedColumn
            Caption = #20633#35387#19968
            DataBinding.FieldName = 'MEMO1'
            Width = 20
            Position.BandIndex = 2
            Position.ColIndex = 0
            Position.RowIndex = 1
          end
          object cxGrid1DBBandedTableView1MEMO2: TcxGridDBBandedColumn
            Caption = #20633#35387#20108
            DataBinding.FieldName = 'MEMO2'
            Width = 20
            Position.BandIndex = 2
            Position.ColIndex = 0
            Position.RowIndex = 2
          end
          object cxGrid1DBBandedTableView1UPTTIME: TcxGridDBBandedColumn
            Caption = #30064#21205#26178#38291
            DataBinding.FieldName = 'UPTTIME'
            Visible = False
            Width = 97
            Position.BandIndex = 2
            Position.ColIndex = 1
            Position.RowIndex = 0
          end
          object cxGrid1DBBandedTableView1UPTEN: TcxGridDBBandedColumn
            Caption = #30064#21205#20154#21729
            DataBinding.FieldName = 'UPTEN'
            Visible = False
            Width = 110
            Position.BandIndex = 2
            Position.ColIndex = 2
            Position.RowIndex = 0
          end
          object cxGrid1DBBandedTableView1MAININVID: TcxGridDBBandedColumn
            Caption = #20027#30332#31080#34399#30908
            DataBinding.FieldName = 'MAININVID'
            Visible = False
            Width = 85
            Position.BandIndex = 0
            Position.ColIndex = 1
            Position.RowIndex = 0
          end
        end
        object cxGrid1Level1: TcxGridLevel
          GridView = cxGrid1DBBandedTableView1
        end
      end
    end
  end
  object dsrInv007: TDataSource
    DataSet = dtmMainH.adoInv007
    Left = 236
    Top = 344
  end
  object cxPropertiesStore: TcxPropertiesStore
    Active = False
    Components = <
      item
        Component = cxGrid1DBBandedTableView1
        Properties.Strings = (
          'Bands')
      end
      item
        Component = cxGrid1DBBandedTableView1BUSINESSID
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1CHARGEDATE
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1CUSTID
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1CUSTSNAME
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1INVADDR
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1INVAMOUNT
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1INVDATE
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1INVFORMAT
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1INVID
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1INVTITLE
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1ISOBSOLETE
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1MAILADDR
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1MAININVID
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1MEMO1
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1MEMO2
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1OBSOLETEREASON
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1SALEAMOUNT
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1TAXAMOUNT
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1TAXRATE
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1TAXTYPE
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1UPTEN
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1UPTTIME
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end
      item
        Component = cxGrid1DBBandedTableView1ZIPCODE
        Properties.Strings = (
          'Position'
          'SortOrder'
          'Visible'
          'Width')
      end>
    StorageName = 'cxPropertiesStore'
    Left = 272
    Top = 344
  end
end
