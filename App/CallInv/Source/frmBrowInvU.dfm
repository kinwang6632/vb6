object frmBrowINV: TfrmBrowINV
  Left = 362
  Top = 173
  BorderStyle = bsDialog
  Caption = #30332#31080#36039#26009#28687#35261' [ CallInv ]'
  ClientHeight = 467
  ClientWidth = 494
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poScreenCenter
  ShowHint = True
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 494
    Height = 467
    Align = alClient
    TabOrder = 0
    object Panel2: TPanel
      Left = 1
      Top = 1
      Width = 492
      Height = 36
      Align = alTop
      TabOrder = 0
      object btnLog: TcxButton
        Left = 2
        Top = 4
        Width = 113
        Height = 27
        Caption = #20659#36865#36890#30693#27511#21490#35338#24687
        TabOrder = 0
        OnClick = btnLogClick
        LookAndFeel.Kind = lfOffice11
        LookAndFeel.NativeStyle = False
        LookAndFeel.SkinName = 'Black'
        LookAndFeel.SkinName = 'Black'
      end
    end
    object cxGrid: TcxGrid
      Left = 1
      Top = 37
      Width = 492
      Height = 429
      Align = alClient
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      LookAndFeel.Kind = lfFlat
      LookAndFeel.SkinName = ''
      object cxGridBrow: TcxGridDBTableView
        NavigatorButtons.ConfirmDelete = False
        FilterBox.CustomizeDialog = False
        OnCustomDrawCell = cxGridBrowCustomDrawCell
        DataController.DataSource = dsINV008
        DataController.Summary.DefaultGroupSummaryItems = <>
        DataController.Summary.FooterSummaryItems = <>
        DataController.Summary.SummaryGroups = <>
        OptionsCustomize.ColumnFiltering = False
        OptionsCustomize.ColumnGrouping = False
        OptionsCustomize.ColumnHidingOnGrouping = False
        OptionsCustomize.ColumnMoving = False
        OptionsCustomize.ColumnSorting = False
        OptionsData.Appending = True
        OptionsData.Deleting = False
        OptionsData.DeletingConfirmation = False
        OptionsView.GroupByBox = False
        Styles.Background = cxStyle1
        object cxColName: TcxGridDBColumn
          Caption = #27396#20301#21517#31281
          DataBinding.FieldName = 'FldName'
          HeaderGlyphAlignmentHorz = taCenter
          Options.Editing = False
          Options.Filtering = False
          Options.Focusing = False
          Options.Grouping = False
          Options.Sorting = False
          Width = 126
        end
        object cxColDesc: TcxGridDBColumn
          Caption = #27396#20301#20839#23481
          DataBinding.FieldName = 'FldDesc'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.ReadOnly = True
          Options.Filtering = False
          Options.Sorting = False
          Width = 352
        end
      end
      object cxGridLevel: TcxGridLevel
        GridView = cxGridBrow
      end
    end
  end
  object dsINV008: TDataSource
    DataSet = cdsBrowInv
    Left = 48
    Top = 72
  end
  object cdsBrowInv: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 18
    Top = 70
  end
  object cxStyleRepository1: TcxStyleRepository
    Left = 362
    Top = 232
    PixelsPerInch = 96
    object cxStyle1: TcxStyle
      AssignedValues = [svColor]
      Color = clBtnFace
    end
  end
end
