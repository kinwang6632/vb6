object frmCD078B3_5: TfrmCD078B3_5
  Left = 398
  Top = 193
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #36949#32004#37329#26781#20214#35215#21063#35373#23450'[frmCD078B33]'
  ClientHeight = 425
  ClientWidth = 446
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
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 384
    Width = 446
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      446
      41)
    object edtDML: TcxTextEdit
      Left = 485
      Top = 5
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
      Left = 280
      Top = 6
      Width = 69
      Height = 26
      Action = actSave
      TabOrder = 1
    end
    object Button1: TButton
      Left = 359
      Top = 6
      Width = 69
      Height = 26
      Action = actCancel
      TabOrder = 2
    end
  end
  object Panel4: TPanel
    Left = 0
    Top = 0
    Width = 446
    Height = 27
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    object Label30: TLabel
      Tag = 1
      Left = 18
      Top = 10
      Width = 120
      Height = 13
      Caption = #36949#32004#37329#38989#26781#20214#35215#21063#35373#23450
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 27
    Width = 446
    Height = 357
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 5
    TabOrder = 2
    object Bevel1: TBevel
      Left = 5
      Top = 349
      Width = 436
      Height = 3
      Align = alBottom
      Shape = bsBottomLine
    end
    object PenalRuleGrid: TcxGrid
      Left = 5
      Top = 5
      Width = 436
      Height = 344
      Align = alClient
      TabOrder = 0
      LookAndFeel.Kind = lfFlat
      object gvPenalRule: TcxGridDBBandedTableView
        NavigatorButtons.ConfirmDelete = False
        OnCustomDrawCell = gvPenalRuleCustomDrawCell
        OnEditing = gvPenalRuleEditing
        DataController.DataSource = PenalRuleDataSource
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
        Styles.Background = StyleController.cxStyle1
        Bands = <
          item
            Width = 412
          end>
        object gvPenalRuleCol1: TcxGridDBBandedColumn
          Caption = #21512#32004#26376#25976'('#36215')'
          DataBinding.FieldName = 'MonthStart'
          PropertiesClassName = 'TcxMaskEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '(\d)|(\d\d)|(\d\d\d)'
          Properties.OnChange = gvPenalRuleCol1PropertiesChange
          Properties.OnEditValueChanged = gvPenalRuleCol1PropertiesEditValueChanged
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Width = 70
          Position.BandIndex = 0
          Position.ColIndex = 1
          Position.RowIndex = 0
        end
        object gvPenalRuleCol2: TcxGridDBBandedColumn
          Caption = #21512#32004#26376#25976'('#36804')'
          DataBinding.FieldName = 'MonthStop'
          PropertiesClassName = 'TcxMaskEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.MaskKind = emkRegExpr
          Properties.EditMask = '(\d)|(\d\d)|(\d\d\d)'
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Width = 68
          Position.BandIndex = 0
          Position.ColIndex = 2
          Position.RowIndex = 0
        end
        object gvPenalRuleCol3: TcxGridDBBandedColumn
          Caption = #32317#32129#32004#37329#38989
          DataBinding.FieldName = 'MonthAmt'
          PropertiesClassName = 'TcxCurrencyEditProperties'
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-,0'
          Properties.MaxValue = 99999999.000000000000000000
          Properties.OnEditValueChanged = gvPenalRuleCol3PropertiesEditValueChanged
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Width = 104
          Position.BandIndex = 0
          Position.ColIndex = 3
          Position.RowIndex = 0
        end
        object gvPenalRuleCol4: TcxGridDBBandedColumn
          Caption = #25353#26376#36958#28187
          DataBinding.FieldName = 'DecreaseAmt'
          PropertiesClassName = 'TcxCurrencyEditProperties'
          Properties.DecimalPlaces = 0
          Properties.DisplayFormat = ',0;-,0'
          Properties.MaxValue = 9999.000000000000000000
          Properties.OnEditValueChanged = gvPenalRuleCol3PropertiesEditValueChanged
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Width = 87
          Position.BandIndex = 0
          Position.ColIndex = 4
          Position.RowIndex = 0
        end
        object gvPenalRuleCol5: TcxGridDBBandedColumn
          Caption = #38917#27425
          DataBinding.FieldName = 'Item'
          Visible = False
          Width = 51
          Position.BandIndex = 0
          Position.ColIndex = 0
          Position.RowIndex = 0
        end
      end
      object glPenalRule: TcxGridLevel
        GridView = gvPenalRule
      end
    end
  end
  object ActionList1: TActionList
    Left = 52
    Top = 212
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
  object PenalRuleDataSource: TDataSource
    DataSet = PenalRuleDataSet
    Left = 52
    Top = 140
  end
  object PenalRuleDataSet: TClientDataSet
    Aggregates = <>
    FieldDefs = <>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 52
    Top = 174
  end
end
