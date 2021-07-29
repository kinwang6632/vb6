object frmInvA02_A: TfrmInvA02_A
  Left = 616
  Top = 281
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = #27331#21488#21015#21360#35352#37636
  ClientHeight = 193
  ClientWidth = 693
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  PixelsPerInch = 96
  TextHeight = 13
  object pnl1: TPanel
    Left = 0
    Top = 0
    Width = 693
    Height = 193
    Align = alClient
    TabOrder = 0
    object cxGrid1: TcxGrid
      Left = 1
      Top = 1
      Width = 691
      Height = 191
      Align = alClient
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      LookAndFeel.Kind = lfFlat
      LookAndFeel.SkinName = ''
      object cxGridDBTableView1: TcxGridDBTableView
        NavigatorButtons.ConfirmDelete = False
        FilterBox.CustomizeDialog = False
        DataController.DataSource = dsMaster
        DataController.Summary.DefaultGroupSummaryItems = <>
        DataController.Summary.FooterSummaryItems = <>
        DataController.Summary.SummaryGroups = <>
        OptionsCustomize.ColumnFiltering = False
        OptionsCustomize.ColumnGrouping = False
        OptionsCustomize.ColumnHidingOnGrouping = False
        OptionsCustomize.ColumnMoving = False
        OptionsCustomize.ColumnSorting = False
        OptionsData.Deleting = False
        OptionsData.DeletingConfirmation = False
        OptionsData.Editing = False
        OptionsData.Inserting = False
        OptionsSelection.CellSelect = False
        OptionsView.GroupByBox = False
        object cxPrintTime: TcxGridDBColumn
          Caption = #21015#21360#26085#26399
          DataBinding.FieldName = 'PRINTTIME'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Width = 173
        end
        object csPrintEn: TcxGridDBColumn
          Caption = #21015#21360#20154#21729
          DataBinding.FieldName = 'PRINTEN'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          HeaderAlignmentHorz = taCenter
          Width = 147
        end
        object csReasonCode: TcxGridDBColumn
          Caption = #21015#21360#21407#22240#20195#30908
          DataBinding.FieldName = 'REASONCODE'
          PropertiesClassName = 'TcxTextEditProperties'
          Width = 109
        end
        object cxReasonName: TcxGridDBColumn
          Caption = #21015#21360#21407#22240#21517#31281
          DataBinding.FieldName = 'REASONNAME'
          Width = 245
        end
      end
      object cxGridLevel1: TcxGridLevel
        GridView = cxGridDBTableView1
      end
    end
  end
  object dspINV046: TDataSetProvider
    DataSet = dtmMainH.adoInv046
    Options = [poAllowCommandText]
    Left = 102
    Top = 107
  end
  object cdsMaster: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspINV046'
    Left = 62
    Top = 103
  end
  object dsMaster: TDataSource
    DataSet = cdsMaster
    Left = 24
    Top = 105
  end
end
