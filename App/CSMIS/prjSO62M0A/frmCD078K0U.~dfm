object frmCD078K: TfrmCD078K
  Left = 392
  Top = 267
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #35373#23450#25286#24115#27604#20363' [frmCD078K]'
  ClientHeight = 244
  ClientWidth = 558
  Color = clBtnFace
  DefaultMonitor = dmMainForm
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object pnl1: TPanel
    Left = 0
    Top = 0
    Width = 558
    Height = 203
    Align = alClient
    TabOrder = 0
    object RankGrid: TcxGrid
      Left = 1
      Top = 1
      Width = 556
      Height = 201
      Align = alClient
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      LookAndFeel.Kind = lfFlat
      object gvRank: TcxGridDBBandedTableView
        NavigatorButtons.ConfirmDelete = False
        OnCustomDrawCell = gvRankCustomDrawCell
        DataController.DataSource = dsRankDataSource
        DataController.Summary.DefaultGroupSummaryItems = <>
        DataController.Summary.FooterSummaryItems = <
          item
            Kind = skSum
            OnGetText = gvRankTcxGridDBDataControllerTcxDataSummaryFooterSummaryItems0GetText
            Column = gvRankColumn3
          end
          item
            Format = ',0.000%;-,0.000%'
            Kind = skSum
            Column = gvRankColumn4
          end
          item
            Kind = skSum
            OnGetText = gvRankTcxGridDBDataControllerTcxDataSummaryFooterSummaryItems0GetText
            Column = gvRankColumn5
          end
          item
            OnGetText = gvRankTcxGridDBDataControllerTcxDataSummaryFooterSummaryItems3GetText
            Column = gvRankColumn2
          end>
        DataController.Summary.SummaryGroups = <>
        OptionsBehavior.GoToNextCellOnEnter = True
        OptionsCustomize.ColumnFiltering = False
        OptionsCustomize.ColumnGrouping = False
        OptionsCustomize.ColumnMoving = False
        OptionsCustomize.ColumnSorting = False
        OptionsCustomize.ColumnVertSizing = False
        OptionsData.Deleting = False
        OptionsData.DeletingConfirmation = False
        OptionsData.Inserting = False
        OptionsView.DataRowHeight = 25
        OptionsView.Footer = True
        OptionsView.FooterMultiSummaries = True
        OptionsView.GroupByBox = False
        OptionsView.HeaderAutoHeight = True
        OptionsView.HeaderEndEllipsis = True
        OptionsView.RowSeparatorColor = clBtnFace
        OptionsView.RowSeparatorWidth = 2
        OptionsView.BandHeaders = False
        Styles.Background = StyleController.cxStyle1
        OnCustomDrawFooterCell = gvRankCustomDrawFooterCell
        Bands = <
          item
            Width = 550
          end>
        object gvRankColumn1: TcxGridDBBandedColumn
          Caption = #25910#36027#38917#30446
          DataBinding.FieldName = 'CitemName'
          PropertiesClassName = 'TcxMemoProperties'
          Properties.ReadOnly = True
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Options.Focusing = False
          Width = 234
          Position.BandIndex = 0
          Position.ColIndex = 0
          Position.RowIndex = 0
        end
        object gvRankColumn2: TcxGridDBBandedColumn
          Caption = #20027#25286#24115#38917#30446
          DataBinding.FieldName = 'MasterItem'
          PropertiesClassName = 'TcxCheckBoxProperties'
          Properties.NullStyle = nssUnchecked
          Properties.ValueChecked = 1
          Properties.ValueUnchecked = 0
          Properties.OnEditValueChanged = gvRankColumn2PropertiesEditValueChanged
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Options.Grouping = False
          Width = 68
          Position.BandIndex = 0
          Position.ColIndex = 1
          Position.RowIndex = 0
        end
        object gvRankColumn3: TcxGridDBBandedColumn
          Caption = #37329#38989
          DataBinding.FieldName = 'Amount'
          PropertiesClassName = 'TcxCurrencyEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          Properties.AssignedValues.MinValue = True
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-,0'
          Properties.MaxValue = 99999999.000000000000000000
          Properties.NullString = '0'
          Properties.OnEditValueChanged = gvRankColumn2PropertiesEditValueChanged
          Properties.OnValidate = gvRankColumn3PropertiesValidate
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Options.Filtering = False
          Options.Grouping = False
          Width = 83
          Position.BandIndex = 0
          Position.ColIndex = 2
          Position.RowIndex = 0
        end
        object gvRankColumn4: TcxGridDBBandedColumn
          Caption = #25286#24115#27604#20363
          DataBinding.FieldName = 'IFRSPercent'
          PropertiesClassName = 'TcxCurrencyEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          Properties.DecimalPlaces = 3
          Properties.DisplayFormat = ',0.000%;-,0.000%'
          Properties.MaxValue = 99.999000000000000000
          Properties.Nullable = False
          Properties.NullString = '0'
          Properties.OnEditValueChanged = gvRankColumn2PropertiesEditValueChanged
          Properties.OnValidate = gvRankColumn4PropertiesValidate
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Options.Filtering = False
          Options.Grouping = False
          Width = 69
          Position.BandIndex = 0
          Position.ColIndex = 3
          Position.RowIndex = 0
        end
        object gvRankColumn5: TcxGridDBBandedColumn
          Caption = 'IFRS'#37329#38989
          DataBinding.FieldName = 'IFRSAmt'
          PropertiesClassName = 'TcxCurrencyEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          Properties.AssignedValues.MaxValue = True
          Properties.DisplayFormat = ',0;-,0'
          Properties.NullString = '0'
          Properties.ReadOnly = True
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Options.Editing = False
          Options.Filtering = False
          Options.Focusing = False
          Options.Moving = False
          Width = 96
          Position.BandIndex = 0
          Position.ColIndex = 4
          Position.RowIndex = 0
        end
      end
      object glRank: TcxGridLevel
        GridView = gvRank
      end
    end
  end
  object pnl2: TPanel
    Left = 0
    Top = 203
    Width = 558
    Height = 41
    Align = alBottom
    TabOrder = 1
    object btnSave: TButton
      Left = 402
      Top = 6
      Width = 69
      Height = 26
      Action = actSave
      TabOrder = 0
    end
    object btnCancel: TButton
      Left = 484
      Top = 6
      Width = 69
      Height = 26
      Action = actCancel
      Cancel = True
      TabOrder = 1
    end
  end
  object cdsRankDataSet: TClientDataSet
    Aggregates = <>
    FieldDefs = <>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 52
    Top = 174
  end
  object dsRankDataSource: TDataSource
    DataSet = cdsRankDataSet
    Left = 52
    Top = 140
  end
  object actlstact: TActionList
    Left = 20
    Top = 142
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
end
