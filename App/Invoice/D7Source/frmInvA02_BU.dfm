object frmInvA02_B: TfrmInvA02_B
  Left = 412
  Top = 307
  BorderStyle = bsDialog
  Caption = 'frmInvA02_B [ '#19978#20659#35352#37636#28687#28768' ]'
  ClientHeight = 211
  ClientWidth = 481
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object cxGrid1: TcxGrid
    Left = 0
    Top = 0
    Width = 481
    Height = 211
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
      object cxUploadType: TcxGridDBColumn
        Caption = #19978#20659#31278#39006
        DataBinding.FieldName = 'UPLOADTYPE'
        PropertiesClassName = 'TcxTextEditProperties'
        Properties.Alignment.Horz = taCenter
        HeaderAlignmentHorz = taCenter
        Options.Editing = False
        Width = 163
      end
      object cxUploadTime: TcxGridDBColumn
        Caption = #19978#20659#26178#38291
        DataBinding.FieldName = 'UPLOADTIME'
        PropertiesClassName = 'TcxTextEditProperties'
        Properties.Alignment.Horz = taCenter
        HeaderAlignmentHorz = taCenter
        Width = 182
      end
      object cxUploadSource: TcxGridDBColumn
        Caption = #35352#37636#20358#28304
        DataBinding.FieldName = 'UPLOADSOURCE'
        PropertiesClassName = 'TcxTextEditProperties'
        Width = 125
      end
    end
    object cxGridLevel1: TcxGridLevel
      GridView = cxGridDBTableView1
    end
  end
  object dsMaster: TDataSource
    Left = 144
    Top = 52
  end
end
