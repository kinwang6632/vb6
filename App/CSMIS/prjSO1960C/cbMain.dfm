object fmSO1960C: TfmSO1960C
  Left = 279
  Top = 169
  Width = 752
  Height = 515
  ActiveControl = txtMac
  Caption = 'EMTA/CM'#36039#26009#31649#29702' [SO1960C]'
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
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 447
    Width = 744
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    object btnSearch: TButton
      Left = 7
      Top = 5
      Width = 60
      Height = 26
      Action = actSearch
      TabOrder = 0
    end
    object btnClear: TButton
      Left = 70
      Top = 5
      Width = 60
      Height = 26
      Action = actClear
      TabOrder = 1
    end
    object Button7: TButton
      Left = 134
      Top = 5
      Width = 60
      Height = 26
      Cancel = True
      Caption = #32080#26463'(&X)'
      TabOrder = 2
      OnClick = Button7Click
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 744
    Height = 69
    Align = alTop
    BevelOuter = bvNone
    BorderWidth = 3
    TabOrder = 1
    object GroupBox1: TGroupBox
      Left = 3
      Top = 3
      Width = 738
      Height = 63
      Align = alClient
      TabOrder = 0
      object Label1: TLabel
        Left = 31
        Top = 28
        Width = 48
        Height = 14
        Caption = 'MAC'#24207#34399
      end
      object cmbSearchType: TComboBox
        Left = 86
        Top = 24
        Width = 92
        Height = 22
        Style = csDropDownList
        ItemHeight = 14
        ItemIndex = 1
        TabOrder = 0
        Text = #23436#20840#26597#35426
        Items.Strings = (
          #27169#31946#26597#35426
          #23436#20840#26597#35426)
      end
      object txtMac: TEdit
        Left = 182
        Top = 24
        Width = 150
        Height = 22
        MaxLength = 20
        TabOrder = 1
      end
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 69
    Width = 744
    Height = 378
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 3
    TabOrder = 2
    object cxGridMac: TcxGrid
      Left = 3
      Top = 3
      Width = 738
      Height = 372
      Align = alClient
      TabOrder = 0
      LookAndFeel.Kind = lfFlat
      object gvMac: TcxGridDBTableView
        NavigatorButtons.ConfirmDelete = False
        DataController.DataSource = dsDataTable
        DataController.KeyFieldNames = 'CMCOMPCODE;HFCMAC'
        DataController.Summary.DefaultGroupSummaryItems = <>
        DataController.Summary.FooterSummaryItems = <>
        DataController.Summary.SummaryGroups = <>
        OptionsCustomize.ColumnFiltering = False
        OptionsData.Deleting = False
        OptionsData.Editing = False
        OptionsSelection.CellSelect = False
        OptionsView.CellEndEllipsis = True
        OptionsView.ExpandButtonsForEmptyDetails = False
        OptionsView.GroupByBox = False
        OptionsView.HeaderEndEllipsis = True
        OptionsView.Indicator = True
        object gvMacCOMPCODE_1: TcxGridDBColumn
          Caption = #29289#26009#20489#21029
          DataBinding.FieldName = 'CMCOMPNAME'
          Width = 82
        end
        object gvMacHFC: TcxGridDBColumn
          Caption = #29289#21697#24207#34399
          DataBinding.FieldName = 'HFCMAC'
          Width = 118
        end
        object gvMacDESCRIPTION: TcxGridDBColumn
          Caption = #21697#21517
          DataBinding.FieldName = 'DESCRIPTION'
          Width = 128
        end
        object gvMacMODELNO: TcxGridDBColumn
          Caption = #27231#31278
          DataBinding.FieldName = 'MODELNO'
          Width = 106
        end
        object gvMacMTAMAC: TcxGridDBColumn
          DataBinding.FieldName = 'MTAMAC'
          Width = 80
        end
        object gvMacCMISSUE: TcxGridDBColumn
          Caption = 'CM'#21295#20837#26085#26399
          DataBinding.FieldName = 'CMISSUE'
          Width = 180
        end
        object gvMacMODEID: TcxGridDBColumn
          Caption = #35373#20633#22411#24335
          DataBinding.FieldName = 'MODEID'
          Width = 78
        end
        object gvMacMODELCODE: TcxGridDBColumn
          Caption = #35373#20633#22411#34399#20195#30908
          DataBinding.FieldName = 'MODELCODE'
          Width = 91
        end
        object gvMacUSEFLAG: TcxGridDBColumn
          Caption = #29289#26009#21487#29992#29376#24907
          DataBinding.FieldName = 'USEFLAG'
          Width = 100
        end
        object gvMacVENDORCODE: TcxGridDBColumn
          Caption = #24288#21830#20195#30908
          DataBinding.FieldName = 'VENDORCODE'
          Width = 77
        end
        object gvMacVENDORNAME: TcxGridDBColumn
          Caption = #24288#21830#21517#31281
          DataBinding.FieldName = 'VENDORNAME'
          Width = 127
        end
        object gvMacSPEC: TcxGridDBColumn
          Caption = #35215#26684
          DataBinding.FieldName = 'SPEC'
          Width = 136
        end
        object gvMacUPDEN: TcxGridDBColumn
          Caption = #30064#21205#20154#21729
          DataBinding.FieldName = 'UPDEN'
          Width = 85
        end
        object gvMacUPDTIME: TcxGridDBColumn
          Caption = #30064#21205#26178#38291
          DataBinding.FieldName = 'UPDTIME'
          Width = 106
        end
        object gvMacBATCHNO: TcxGridDBColumn
          Caption = #25209#34399
          DataBinding.FieldName = 'BATCHNO'
          Width = 119
        end
        object gvMacMATERIALNO: TcxGridDBColumn
          Caption = #26009#34399
          DataBinding.FieldName = 'MATERIALNO'
          Width = 130
        end
      end
      object gvSo: TcxGridDBTableView
        NavigatorButtons.ConfirmDelete = False
        OnCustomDrawCell = gvSoCustomDrawCell
        DataController.DataSource = dsSoData
        DataController.DetailKeyFieldNames = 'HFCMAC'
        DataController.MasterKeyFieldNames = 'HFCMAC'
        DataController.Summary.DefaultGroupSummaryItems = <>
        DataController.Summary.FooterSummaryItems = <>
        DataController.Summary.SummaryGroups = <>
        OptionsCustomize.ColumnFiltering = False
        OptionsData.Deleting = False
        OptionsData.Editing = False
        OptionsData.Inserting = False
        OptionsSelection.CellSelect = False
        OptionsView.CellEndEllipsis = True
        OptionsView.GroupByBox = False
        OptionsView.HeaderEndEllipsis = True
        object gvSoKeyId: TcxGridDBColumn
          DataBinding.FieldName = 'HFCMAC'
          Visible = False
        end
        object gvSoCompCode: TcxGridDBColumn
          Caption = #31995#32113#21488#20195#30908
          DataBinding.FieldName = 'COMPCODE'
          Visible = False
        end
        object gvSoCompName: TcxGridDBColumn
          Caption = #31995#32113#21488
          DataBinding.FieldName = 'COMPNAME'
          Width = 71
        end
        object gvSoCustId: TcxGridDBColumn
          Caption = #23458#32232
          DataBinding.FieldName = 'CUSTID'
          Width = 75
        end
        object gvSoInstDate: TcxGridDBColumn
          Caption = #35037#27231#26085#26399
          DataBinding.FieldName = 'INSTDATE'
          SortIndex = 0
          SortOrder = soDescending
          Width = 180
        end
        object gvSoFaciStatus: TcxGridDBColumn
          Caption = #29376#24907
          DataBinding.FieldName = 'FACISTATUS'
          Width = 95
        end
      end
      object glMac: TcxGridLevel
        GridView = gvMac
        object glSo: TcxGridLevel
          GridView = gvSo
        end
      end
    end
  end
  object ActionList1: TActionList
    Left = 24
    Top = 309
    object actSearch: TAction
      Caption = 'F3.'#23563#25214
      ShortCut = 114
      OnExecute = actSearchExecute
    end
    object actClear: TAction
      Caption = 'F4.'#28165#38500
      ShortCut = 115
      OnExecute = actClearExecute
    end
  end
  object dsDataTable: TDataSource
    DataSet = DataTable
    Left = 124
    Top = 185
  end
  object ADOConnection1: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 256
    Top = 149
  end
  object DataReader: TADOQuery
    CacheSize = 1000
    Connection = ADOConnection1
    Parameters = <>
    Left = 224
    Top = 149
  end
  object DataTable: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 156
    Top = 185
  end
  object DataGeter: TADOQuery
    CacheSize = 1000
    Connection = ADOConnection1
    Parameters = <>
    Left = 224
    Top = 185
  end
  object DataSO507: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 156
    Top = 149
  end
  object SoData: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 156
    Top = 225
  end
  object dsSoData: TDataSource
    DataSet = SoData
    Left = 124
    Top = 225
  end
end
