object frmInvA01_1: TfrmInvA01_1
  Left = 266
  Top = 102
  BorderStyle = bsDialog
  Caption = #30332#31080#36039#26009
  ClientHeight = 573
  ClientWidth = 792
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 792
    Height = 573
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 2
    TabOrder = 0
    object Bevel1: TBevel
      Left = 2
      Top = 28
      Width = 788
      Height = 3
      Align = alTop
      Shape = bsSpacer
    end
    object Panel2: TPanel
      Left = 2
      Top = 2
      Width = 788
      Height = 26
      Align = alTop
      BevelOuter = bvNone
      Color = clGray
      TabOrder = 0
      object lblMasterTitle: TLabel
        Left = 14
        Top = 5
        Width = 80
        Height = 16
        Caption = #26410#38283#31435#30332#31080
        Font.Charset = ANSI_CHARSET
        Font.Color = clWhite
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
    end
    object Panel3: TPanel
      Left = 2
      Top = 31
      Width = 788
      Height = 277
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 1
      object GridMaster: TcxGrid
        Left = 0
        Top = 0
        Width = 788
        Height = 277
        Align = alClient
        TabOrder = 0
        object GridMasterDBTableView1: TcxGridDBTableView
          NavigatorButtons.ConfirmDelete = False
          DataController.DataSource = drInv016
          DataController.Summary.DefaultGroupSummaryItems = <>
          DataController.Summary.FooterSummaryItems = <>
          DataController.Summary.SummaryGroups = <>
          OptionsCustomize.ColumnFiltering = False
          OptionsData.Deleting = False
          OptionsData.DeletingConfirmation = False
          OptionsData.Editing = False
          OptionsData.Inserting = False
          OptionsSelection.CellSelect = False
          OptionsView.CellEndEllipsis = True
          OptionsView.GroupByBox = False
          OptionsView.HeaderEndEllipsis = True
          OptionsView.Indicator = True
          Styles.Background = dtmMain.cxGridBackGroundStyle
          Styles.Content = dtmMain.cxGridBackGroundStyle
          Styles.Inactive = dtmMain.cxGridInActiveStyle
          Styles.Selection = dtmMain.cxGridActiveStyle
          object GridMasterDBTableView1SEQ: TcxGridDBColumn
            Caption = #27969#27700#34399
            DataBinding.FieldName = 'SEQ'
          end
          object GridMasterDBTableView1CUSTID: TcxGridDBColumn
            Caption = #23458#32232
            DataBinding.FieldName = 'CUSTID'
          end
          object GridMasterDBTableView1TITLE: TcxGridDBColumn
            Caption = #30332#31080#25260#38957
            DataBinding.FieldName = 'TITLE'
          end
          object GridMasterDBTableView1TEL: TcxGridDBColumn
            Caption = #38651#35441
            DataBinding.FieldName = 'TEL'
          end
          object GridMasterDBTableView1BUSINESSID: TcxGridDBColumn
            Caption = #32113#32232
            DataBinding.FieldName = 'BUSINESSID'
            Width = 78
          end
          object GridMasterDBTableView1ZIPCODE: TcxGridDBColumn
            Caption = #37109#36958#21312#34399
            DataBinding.FieldName = 'ZIPCODE'
            Width = 74
          end
          object GridMasterDBTableView1INVADDR: TcxGridDBColumn
            Caption = #30332#31080#22320#22336
            DataBinding.FieldName = 'INVADDR'
            Width = 85
          end
          object GridMasterDBTableView1MAILADDR: TcxGridDBColumn
            Caption = #23492#20214#22320#22336
            DataBinding.FieldName = 'MAILADDR'
            Width = 134
          end
          object GridMasterDBTableView1CHARGEDATE: TcxGridDBColumn
            Caption = #25910#36027#26085
            DataBinding.FieldName = 'CHARGEDATE'
            Width = 70
          end
          object GridMasterDBTableView1DESCRIPTION: TcxGridDBColumn
            Caption = #31237#21029
            DataBinding.FieldName = 'DESCRIPTION'
            HeaderAlignmentHorz = taCenter
            Width = 73
          end
          object GridMasterDBTableView1TAXRATE: TcxGridDBColumn
            Caption = #31237#29575
            DataBinding.FieldName = 'TAXRATE'
            HeaderAlignmentHorz = taCenter
            Width = 71
          end
          object GridMasterDBTableView1SALEAMOUNT: TcxGridDBColumn
            Caption = #37559#21806#38989
            DataBinding.FieldName = 'SALEAMOUNT'
            PropertiesClassName = 'TcxCurrencyEditProperties'
            Properties.DisplayFormat = ',0.00;-,0.00'
            HeaderAlignmentHorz = taRightJustify
            Width = 106
          end
          object GridMasterDBTableView1TAXAMOUNT: TcxGridDBColumn
            Caption = #31237#38989
            DataBinding.FieldName = 'TAXAMOUNT'
            PropertiesClassName = 'TcxCurrencyEditProperties'
            Properties.DisplayFormat = ',0.00;-,0.00'
            HeaderAlignmentHorz = taRightJustify
            Width = 97
          end
          object GridMasterDBTableView1INVAMOUNT: TcxGridDBColumn
            Caption = #30332#31080#37329#38989
            DataBinding.FieldName = 'INVAMOUNT'
            PropertiesClassName = 'TcxCurrencyEditProperties'
            Properties.DisplayFormat = ',0.00;-,0.00'
            HeaderAlignmentHorz = taRightJustify
            Width = 95
          end
          object GridMasterDBTableView1HOWTOCREATE: TcxGridDBColumn
            Caption = #38283#31435#20358#28304
            DataBinding.FieldName = 'HOWTOCREATE'
            HeaderAlignmentHorz = taCenter
            Width = 108
          end
          object GridMasterDBTableView1CHARGETITLE: TcxGridDBColumn
            Caption = #25910#20214#20154
            DataBinding.FieldName = 'CHARGETITLE'
            Width = 93
          end
          object GridMasterDBTableView1UPTTIME: TcxGridDBColumn
            Caption = #30064#21205#26178#38291
            DataBinding.FieldName = 'UPTTIME'
            Width = 86
          end
          object GridMasterDBTableView1UPTEN: TcxGridDBColumn
            Caption = #30064#21205#20154#21729
            DataBinding.FieldName = 'UPTEN'
          end
        end
        object GridMasterLevel1: TcxGridLevel
          GridView = GridMasterDBTableView1
        end
      end
    end
    object cxSplitter1: TcxSplitter
      Left = 2
      Top = 308
      Width = 788
      Height = 8
      HotZoneClassName = 'TcxSimpleStyle'
      AlignSplitter = salBottom
      Control = Panel7
    end
    object Panel5: TPanel
      Left = 2
      Top = 533
      Width = 788
      Height = 38
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 3
      object BitBtn1: TBitBtn
        Left = 349
        Top = 8
        Width = 85
        Height = 26
        Caption = #30906#23450
        Default = True
        ModalResult = 1
        TabOrder = 0
        OnClick = BitBtn1Click
        Glyph.Data = {
          DE010000424DDE01000000000000760000002800000024000000120000000100
          0400000000006801000000000000000000001000000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
          3333333333333333333333330000333333333333333333333333F33333333333
          00003333344333333333333333388F3333333333000033334224333333333333
          338338F3333333330000333422224333333333333833338F3333333300003342
          222224333333333383333338F3333333000034222A22224333333338F338F333
          8F33333300003222A3A2224333333338F3838F338F33333300003A2A333A2224
          33333338F83338F338F33333000033A33333A222433333338333338F338F3333
          0000333333333A222433333333333338F338F33300003333333333A222433333
          333333338F338F33000033333333333A222433333333333338F338F300003333
          33333333A222433333333333338F338F00003333333333333A22433333333333
          3338F38F000033333333333333A223333333333333338F830000333333333333
          333A333333333333333338330000333333333333333333333333333333333333
          0000}
        NumGlyphs = 2
      end
    end
    object Panel7: TPanel
      Left = 2
      Top = 316
      Width = 788
      Height = 217
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 4
      object Bevel2: TBevel
        Left = 0
        Top = 26
        Width = 788
        Height = 3
        Align = alTop
        Shape = bsSpacer
      end
      object Panel6: TPanel
        Left = 0
        Top = 0
        Width = 788
        Height = 26
        Align = alTop
        BevelOuter = bvNone
        Color = clGray
        TabOrder = 0
        object lblDetailTitle: TLabel
          Left = 16
          Top = 5
          Width = 64
          Height = 16
          Caption = #30332#31080#26126#32048
          Font.Charset = ANSI_CHARSET
          Font.Color = clWhite
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentFont = False
        end
      end
      object Panel4: TPanel
        Left = 0
        Top = 29
        Width = 788
        Height = 188
        Align = alClient
        BevelOuter = bvNone
        TabOrder = 1
        object GridDetail: TcxGrid
          Left = 0
          Top = 0
          Width = 788
          Height = 188
          Align = alClient
          TabOrder = 0
          object GridDetailDBTableView1: TcxGridDBTableView
            NavigatorButtons.ConfirmDelete = False
            DataController.DataSource = drInv017
            DataController.Summary.DefaultGroupSummaryItems = <>
            DataController.Summary.FooterSummaryItems = <>
            DataController.Summary.SummaryGroups = <>
            OptionsCustomize.ColumnFiltering = False
            OptionsData.Deleting = False
            OptionsData.DeletingConfirmation = False
            OptionsData.Editing = False
            OptionsData.Inserting = False
            OptionsSelection.CellSelect = False
            OptionsView.CellEndEllipsis = True
            OptionsView.GroupByBox = False
            OptionsView.HeaderEndEllipsis = True
            OptionsView.Indicator = True
            Styles.Background = dtmMain.cxGridBackGroundStyle
            Styles.Content = dtmMain.cxGridBackGroundStyle
            Styles.Inactive = dtmMain.cxGridInActiveStyle
            Styles.Selection = dtmMain.cxGridActiveStyle
            object GridDetailDBTableView1BILLID: TcxGridDBColumn
              Caption = #25910#36027#21934#34399
              DataBinding.FieldName = 'BILLID'
              Width = 87
            end
            object GridDetailDBTableView1BILLIDITEMNO: TcxGridDBColumn
              Caption = #38917#27425
              DataBinding.FieldName = 'BILLIDITEMNO'
              HeaderAlignmentHorz = taRightJustify
              Width = 40
            end
            object GridDetailDBTableView1ITEMDESCRIPTION: TcxGridDBColumn
              Caption = #21697#21517
              DataBinding.FieldName = 'ITEMDESCRIPTION'
              Width = 112
            end
            object GridDetailDBTableView1CHARGEDATE: TcxGridDBColumn
              Caption = #25910#36027#26085
              DataBinding.FieldName = 'CHARGEDATE'
              Width = 64
            end
            object GridDetailDBTableView1TAXDESCRIPTION: TcxGridDBColumn
              Caption = #31237#21029
              DataBinding.FieldName = 'TAXDESCRIPTION'
              HeaderAlignmentHorz = taCenter
              Width = 62
            end
            object GridDetailDBTableView1TAXRATE: TcxGridDBColumn
              Caption = #31237#29575
              DataBinding.FieldName = 'TAXRATE'
              HeaderAlignmentHorz = taRightJustify
              Width = 77
            end
            object GridDetailDBTableView1QUANTITY: TcxGridDBColumn
              Caption = #25976#37327
              DataBinding.FieldName = 'QUANTITY'
              HeaderAlignmentHorz = taRightJustify
              Width = 50
            end
            object GridDetailDBTableView1UNITPRICE: TcxGridDBColumn
              Caption = #21934#20729
              DataBinding.FieldName = 'UNITPRICE'
              PropertiesClassName = 'TcxCurrencyEditProperties'
              Properties.DisplayFormat = ',0.00;-,0.00'
              HeaderAlignmentHorz = taRightJustify
              Width = 81
            end
            object GridDetailDBTableView1TAXAMOUNT: TcxGridDBColumn
              Caption = #31237#38989
              DataBinding.FieldName = 'TAXAMOUNT'
              PropertiesClassName = 'TcxCurrencyEditProperties'
              Properties.DisplayFormat = ',0.00;-,0.00'
              HeaderAlignmentHorz = taRightJustify
              Width = 82
            end
            object GridDetailDBTableView1TOTALAMOUNT: TcxGridDBColumn
              Caption = #32317#37329#38989
              DataBinding.FieldName = 'TOTALAMOUNT'
              PropertiesClassName = 'TcxCurrencyEditProperties'
              Properties.DisplayFormat = ',0.00;-,0.00'
              HeaderAlignmentHorz = taRightJustify
              Width = 85
            end
            object GridDetailDBTableView1STARTDATE: TcxGridDBColumn
              Caption = #26377#25928#26085#36215
              DataBinding.FieldName = 'STARTDATE'
              Width = 82
            end
            object GridDetailDBTableView1ENDDATE: TcxGridDBColumn
              Caption = #26377#25928#26085#36804
              DataBinding.FieldName = 'ENDDATE'
              Width = 86
            end
            object GridDetailDBTableView1CHARGEEN: TcxGridDBColumn
              Caption = #25910#36027#20154#21729
              DataBinding.FieldName = 'CHARGEEN'
            end
            object GridDetailDBTableView1SERVICETYPE: TcxGridDBColumn
              Caption = #26381#21209#21029
              DataBinding.FieldName = 'SERVICETYPE'
              HeaderAlignmentHorz = taCenter
            end
          end
          object GridDetailLevel1: TcxGridLevel
            GridView = GridDetailDBTableView1
          end
        end
      end
    end
  end
  object drInv016: TDataSource
    DataSet = dtmMain.adoComm
    Left = 22
    Top = 216
  end
  object drInv017: TDataSource
    DataSet = dtmMain.adoComm2
    Left = 22
    Top = 441
  end
  object Timer1: TTimer
    Enabled = False
    Interval = 300
    OnTimer = Timer1Timer
    Left = 66
    Top = 215
  end
  object cxPropertiesStore: TcxPropertiesStore
    Active = False
    Components = <
      item
        Component = GridDetailDBTableView1BILLID
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridDetailDBTableView1BILLIDITEMNO
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridDetailDBTableView1CHARGEDATE
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridDetailDBTableView1CHARGEEN
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridDetailDBTableView1ENDDATE
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridDetailDBTableView1ITEMDESCRIPTION
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridDetailDBTableView1QUANTITY
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridDetailDBTableView1SERVICETYPE
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridDetailDBTableView1STARTDATE
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridDetailDBTableView1TAXAMOUNT
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridDetailDBTableView1TAXDESCRIPTION
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridDetailDBTableView1TAXRATE
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridDetailDBTableView1TOTALAMOUNT
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridDetailDBTableView1UNITPRICE
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1BUSINESSID
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1CHARGEDATE
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1CHARGETITLE
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1CUSTID
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1DESCRIPTION
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1HOWTOCREATE
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1INVADDR
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1INVAMOUNT
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1MAILADDR
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1SALEAMOUNT
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1SEQ
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1TAXAMOUNT
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1TAXRATE
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1TEL
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1TITLE
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1UPTEN
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1UPTTIME
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end
      item
        Component = GridMasterDBTableView1ZIPCODE
        Properties.Strings = (
          'SortOrder'
          'Tag'
          'Width')
      end>
    StorageName = 'cxPropertiesStore'
    Left = 114
    Top = 215
  end
end
