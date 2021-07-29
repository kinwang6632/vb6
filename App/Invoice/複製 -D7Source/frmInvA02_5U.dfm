object frmInvA02_5: TfrmInvA02_5
  Left = 232
  Top = 105
  ActiveControl = cxGrid
  BorderStyle = bsDialog
  Caption = 'frmInvA02_5'
  ClientHeight = 581
  ClientWidth = 659
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 659
    Height = 40
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      659
      40)
    object btnCancel: TBitBtn
      Left = 570
      Top = 7
      Width = 82
      Height = 26
      Anchors = [akTop, akRight]
      Caption = #21462#28040
      TabOrder = 0
      OnClick = btnCancelClick
      Glyph.Data = {
        36030000424D3603000000000000360000002800000010000000100000000100
        18000000000000030000120B0000120B00000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FF0000A40206B0030AAE0000A6000098FF00FFFF
        00FFFF00FFFF00FF0000A700009A0000A7FF00FFFF00FFFF00FFFF00FF0000A9
        1844F6194DF81031D20102AB0000B6FF00FFFF00FF0000B10928D7092ED70313
        B30000ACFF00FFFF00FFFF00FF0103B32451F91F52FF1D4FFF1744E8040BB000
        00B00000AC0D2EDD1142F90D3DF50B3BF0041ABC0000A5FF00FFFF00FF0000AE
        1832DB285BFF2456FF2253FF1B4BF1050DB10F30DD164AFE1344F91041F60E3E
        F60A3CF000009FFF00FFFF00FF0000BE1F37DD3A6FFF2C5EFF295AFF2657FF20
        52FC1C4FFF194AFD1646FA1445FA0F3DF2020AB10000A8FF00FFFF00FFFF00FF
        0000C8121DC83D6AFB3567FF2C5DFF2859FF2253FF1D4EFF1A4DFF123DED0002
        AC0000BAFF00FFFF00FFFF00FFFF00FFFF00FF0000CC0000B62E4EE73668FF2E
        5EFF2859FF2254FF163DEA0000A80000ABFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FF0000BF253FDF3B6DFF3464FF2E5EFF2759FF1B46EA0001AC0000
        A9FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF0203C84B7CFF4170FF3B
        6BFF396CFF2D5EFF2558FF1336D70000B6FF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FF0000D9263CDB5080FF4575FF3662FA0D14C33C6DFF2A5BFF2053FD0B1D
        C40000C0FF00FFFF00FFFF00FFFF00FFFF00FF0000CB527CFA5081FF4B7DFF0B
        13C90000BB0E15C7386AFF2456FF1A4AF20207B30000B5FF00FFFF00FFFF00FF
        FF00FF131CDD6A9CFF5788FF2B46E70000CDFF00FF0000CD0F1BCB3065FF1F51
        FF1439DD0000B1FF00FFFF00FFFF00FFFF00FF0000DE3A52E45782FB0101D0FF
        00FFFF00FFFF00FF0000CC1426D6265AFF0F2EE30103B8FF00FFFF00FFFF00FF
        FF00FFFF00FF0000CF0000C00000CEFF00FFFF00FFFF00FFFF00FF0000C40001
        B80000B5000077FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
    end
    object StaticText1: TStaticText
      Left = 17
      Top = 11
      Width = 39
      Height = 20
      Caption = #23458#32232':'
      TabOrder = 1
    end
    object StaticText2: TStaticText
      Left = 146
      Top = 12
      Width = 69
      Height = 20
      Caption = #23458#25142#21517#31281':'
      TabOrder = 2
    end
    object btnOk: TBitBtn
      Left = 483
      Top = 7
      Width = 82
      Height = 26
      Anchors = [akTop, akRight]
      Caption = #30906#23450
      Default = True
      TabOrder = 3
      OnClick = btnOkClick
      Glyph.Data = {
        36030000424D3603000000000000360000002800000010000000100000000100
        18000000000000030000F00A0000F00A00000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF008B0000820003880500
        8A00FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FF00A300008E0013AC2715B32B06940E009300FF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FF00B10000940016AE2E17B23114AD2A13
        B12906920B009500FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF00A900
        009A0019B2331CB63A18B33434BC4D17B03013B02A06910B009500FF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FF0CA3161FB73E20BA421FB84011A824018E0162
        C77119B23213B12B06910B009500FF00FFFF00FFFF00FFFF00FFFF00FF28B338
        42C8651FBB4615B32C00A40000920000930061C87119B13113B22C06900B0095
        00FF00FFFF00FFFF00FFFF00FF0EB92443C45636BC4E00AF08FF00FFFF00FF00
        AA0000920063CA7217B13114B12C069109009B00FF00FFFF00FFFF00FFFF00FF
        10D2240DBE20009700FF00FFFF00FFFF00FF00A80000930163CA7217B13114B3
        2D069009008F00FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FF00AD0000940064CB7515B13116B330028904FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF00AB0001970366CD
        782ABC46008A02FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FF009A00009401049306008400FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
    end
    object lblCustID: TcxLabel
      Left = 55
      Top = 8
      AutoSize = False
      Caption = 'lblCustID'
      ParentColor = False
      ParentFont = False
      Properties.Alignment.Vert = taVCenter
      Style.StyleController = dtmMain.cxLabelStyle
      Height = 24
      Width = 85
    end
    object lblCustSName: TcxLabel
      Left = 213
      Top = 8
      Anchors = [akLeft, akTop, akRight]
      AutoSize = False
      Caption = 'lblCustSName'
      ParentColor = False
      ParentFont = False
      Properties.Alignment.Vert = taVCenter
      Style.StyleController = dtmMain.cxLabelStyle
      Height = 24
      Width = 260
    end
  end
  object cxGrid: TcxGrid
    Left = 0
    Top = 40
    Width = 659
    Height = 541
    Align = alClient
    TabOrder = 1
    object cxGridView: TcxGridDBBandedTableView
      OnDblClick = cxGridViewDblClick
      NavigatorButtons.ConfirmDelete = False
      DataController.DataSource = DataSource1
      DataController.Summary.DefaultGroupSummaryItems = <>
      DataController.Summary.FooterSummaryItems = <>
      DataController.Summary.SummaryGroups = <>
      OptionsCustomize.ColumnFiltering = False
      OptionsCustomize.ColumnSorting = False
      OptionsCustomize.BandMoving = False
      OptionsData.CancelOnExit = False
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
          Caption = #35531#36984#25799#25260#38957
          FixedKind = fkLeft
          HeaderAlignmentHorz = taLeftJustify
          Width = 332
        end
        item
          Width = 372
        end
        item
          Width = 192
        end>
      object cxGridViewTITLESNAME: TcxGridDBBandedColumn
        Caption = #25260#38957#31777#31281
        DataBinding.FieldName = 'TITLESNAME'
        Width = 89
        Position.BandIndex = 0
        Position.ColIndex = 0
        Position.LineCount = 2
        Position.RowIndex = 0
      end
      object cxGridViewTITLENAME: TcxGridDBBandedColumn
        Caption = #30332#31080#25260#38957
        DataBinding.FieldName = 'TITLENAME'
        Width = 243
        Position.BandIndex = 0
        Position.ColIndex = 1
        Position.LineCount = 2
        Position.RowIndex = 0
      end
      object cxGridViewBUSINESSID: TcxGridDBBandedColumn
        Caption = #32113#19968#32232#34399
        DataBinding.FieldName = 'BUSINESSID'
        Width = 81
        Position.BandIndex = 1
        Position.ColIndex = 0
        Position.RowIndex = 0
      end
      object cxGridViewMZIPCODE: TcxGridDBBandedColumn
        Caption = #37109#36958#21312#34399
        DataBinding.FieldName = 'MZIPCODE'
        Width = 81
        Position.BandIndex = 1
        Position.ColIndex = 0
        Position.RowIndex = 1
      end
      object cxGridViewMAILADDR: TcxGridDBBandedColumn
        Caption = #23492#20214#22320#22336
        DataBinding.FieldName = 'MAILADDR'
        Width = 291
        Position.BandIndex = 1
        Position.ColIndex = 1
        Position.RowIndex = 1
      end
      object cxGridViewINVADDR: TcxGridDBBandedColumn
        Caption = #30332#31080#22320#22336
        DataBinding.FieldName = 'INVADDR'
        Width = 291
        Position.BandIndex = 1
        Position.ColIndex = 1
        Position.RowIndex = 0
      end
      object cxGridViewMEMO: TcxGridDBBandedColumn
        Caption = #20633#35387
        DataBinding.FieldName = 'MEMO'
        Width = 402
        Position.BandIndex = 2
        Position.ColIndex = 0
        Position.LineCount = 2
        Position.RowIndex = 0
      end
    end
    object cxGridLevel1: TcxGridLevel
      GridView = cxGridView
    end
  end
  object DataSource1: TDataSource
    DataSet = dtmMainH.adoComm
    Left = 112
    Top = 188
  end
end
