object frmCD078B1: TfrmCD078B1
  Left = 285
  Top = 181
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #20778#24800#32068#21512' [frmCD078B1]'
  ClientHeight = 563
  ClientWidth = 764
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Bevel1: TBevel
    Left = 0
    Top = 249
    Width = 764
    Height = 2
    Align = alTop
    Shape = bsSpacer
  end
  object Panel1: TPanel
    Left = 0
    Top = 525
    Width = 764
    Height = 38
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      764
      38)
    object Bevel4: TBevel
      Left = 0
      Top = 0
      Width = 764
      Height = 3
      Align = alTop
      Shape = bsTopLine
    end
    object btnSave: TButton
      Left = 7
      Top = 6
      Width = 70
      Height = 26
      Action = actSave
      TabOrder = 0
    end
    object edtDML: TcxTextEdit
      Left = 689
      Top = 7
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
      TabOrder = 1
      Width = 67
    end
    object btnInsert: TButton
      Left = 79
      Top = 6
      Width = 70
      Height = 26
      Action = actInsert
      TabOrder = 2
    end
    object btnUpdate: TButton
      Left = 223
      Top = 6
      Width = 70
      Height = 26
      Action = actUpdate
      TabOrder = 3
    end
    object btnDelete: TButton
      Left = 155
      Top = 6
      Width = 70
      Height = 26
      Action = actDelete
      TabOrder = 4
    end
    object btnCancel: TButton
      Left = 293
      Top = 6
      Width = 70
      Height = 26
      Action = actCancel
      TabOrder = 5
    end
    object btnIFRS: TButton
      Left = 370
      Top = 5
      Width = 80
      Height = 26
      Caption = #35373#23450#25286#24115#27604#20363
      TabOrder = 6
      OnClick = btnIFRSClick
    end
    object btnShare: TButton
      Left = 455
      Top = 5
      Width = 94
      Height = 26
      Caption = #35373#23450#20998#25892#35519#25972#25976
      TabOrder = 7
      OnClick = btnShareClick
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 251
    Width = 764
    Height = 274
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 1
    object MainPageControl: TcxPageControl
      Left = 0
      Top = 0
      Width = 764
      Height = 274
      ActivePage = Sheet3
      Align = alClient
      OwnerDraw = True
      Style = 1
      TabOrder = 0
      TabPosition = tpBottom
      OnChange = MainPageControlChange
      OnDrawTab = MainPageControlDrawTab
      ClientRectBottom = 251
      ClientRectLeft = 2
      ClientRectRight = 762
      ClientRectTop = 2
      object Sheet3: TcxTabSheet
        Caption = #26234#24935#22411#27966#24037#21443#25976
        object Panel5: TPanel
          Left = 0
          Top = 0
          Width = 760
          Height = 249
          Align = alClient
          BevelOuter = bvNone
          BorderWidth = 3
          TabOrder = 0
          object CD078BGrid: TcxGrid
            Left = 3
            Top = 3
            Width = 754
            Height = 243
            Align = alClient
            TabOrder = 0
            LookAndFeel.Kind = lfFlat
            object CD078BBandedTableView: TcxGridDBBandedTableView
              OnDblClick = BandedTableViewDblClick
              NavigatorButtons.ConfirmDelete = False
              DataController.DataSource = CD078BDataSource
              DataController.Summary.DefaultGroupSummaryItems = <>
              DataController.Summary.FooterSummaryItems = <>
              DataController.Summary.SummaryGroups = <>
              OptionsCustomize.ColumnFiltering = False
              OptionsData.Deleting = False
              OptionsData.DeletingConfirmation = False
              OptionsData.Editing = False
              OptionsData.Inserting = False
              OptionsSelection.CellSelect = False
              OptionsSelection.HideFocusRectOnExit = False
              OptionsView.CellEndEllipsis = True
              OptionsView.GroupByBox = False
              OptionsView.HeaderAutoHeight = True
              OptionsView.HeaderEndEllipsis = True
              OptionsView.Indicator = True
              OptionsView.BandHeaders = False
              Styles.Content = StyleController.GridBackgroundStyle
              Styles.Header = StyleController.HeaderStyle
              Bands = <
                item
                  Width = 754
                end>
              object CD078BBandedTableViewINSTNAME: TcxGridDBBandedColumn
                Caption = #35037#27231#39006#21029
                DataBinding.FieldName = 'INSTNAME'
                Width = 253
                Position.BandIndex = 0
                Position.ColIndex = 1
                Position.RowIndex = 0
              end
              object CD078BBandedTableViewSERVICETYPE: TcxGridDBBandedColumn
                Caption = #26381#21209#39006#21029
                DataBinding.FieldName = 'SERVICETYPE'
                Width = 61
                Position.BandIndex = 0
                Position.ColIndex = 0
                Position.RowIndex = 0
              end
              object CD078BBandedTableViewGROUPNAME: TcxGridDBBandedColumn
                Caption = #31649#27966#39006#21029
                DataBinding.FieldName = 'GROUPNAME'
                Width = 152
                Position.BandIndex = 0
                Position.ColIndex = 2
                Position.RowIndex = 0
              end
              object CD078BBandedTableViewWORKUNIT: TcxGridDBBandedColumn
                Caption = #27966#24037#40670#25976
                DataBinding.FieldName = 'WORKUNIT'
                HeaderAlignmentHorz = taRightJustify
                Width = 87
                Position.BandIndex = 0
                Position.ColIndex = 3
                Position.RowIndex = 0
              end
              object CD078BBandedTableViewREFNO: TcxGridDBBandedColumn
                Caption = #21443#32771#34399
                DataBinding.FieldName = 'REFNO'
                HeaderAlignmentHorz = taRightJustify
                HeaderAlignmentVert = vaCenter
                Width = 76
                Position.BandIndex = 0
                Position.ColIndex = 4
                Position.RowIndex = 0
              end
              object CD078BBandedTableViewINTERDEPEND: TcxGridDBBandedColumn
                Caption = #26381#21209#20381#23384
                DataBinding.FieldName = 'INTERDEPEND'
                PropertiesClassName = 'TcxCheckBoxProperties'
                Properties.DisplayChecked = '1'
                Properties.DisplayUnchecked = '0'
                Properties.NullStyle = nssUnchecked
                Properties.ValueChecked = '1'
                Properties.ValueUnchecked = '0'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 72
                Position.BandIndex = 0
                Position.ColIndex = 5
                Position.RowIndex = 0
              end
            end
            object CD078BGridLevel: TcxGridLevel
              GridView = CD078BBandedTableView
            end
          end
        end
      end
      object Sheet1: TcxTabSheet
        Caption = #35373#20633#21443#25976
        object Panel2: TPanel
          Left = 0
          Top = 0
          Width = 760
          Height = 249
          Align = alClient
          BevelOuter = bvNone
          BorderWidth = 3
          TabOrder = 0
          object CD078DGrid: TcxGrid
            Left = 3
            Top = 3
            Width = 754
            Height = 243
            Align = alClient
            TabOrder = 0
            LookAndFeel.Kind = lfFlat
            object CD078DBandedTableView: TcxGridDBBandedTableView
              OnDblClick = BandedTableViewDblClick
              NavigatorButtons.ConfirmDelete = False
              DataController.DataSource = CD078DDataSource
              DataController.Summary.DefaultGroupSummaryItems = <>
              DataController.Summary.FooterSummaryItems = <>
              DataController.Summary.SummaryGroups = <>
              OptionsCustomize.ColumnFiltering = False
              OptionsCustomize.ColumnSorting = False
              OptionsData.Deleting = False
              OptionsData.DeletingConfirmation = False
              OptionsData.Editing = False
              OptionsData.Inserting = False
              OptionsSelection.CellSelect = False
              OptionsSelection.HideFocusRectOnExit = False
              OptionsView.CellEndEllipsis = True
              OptionsView.GroupByBox = False
              OptionsView.HeaderAutoHeight = True
              OptionsView.HeaderEndEllipsis = True
              OptionsView.Indicator = True
              OptionsView.BandHeaders = False
              Styles.Content = StyleController.GridBackgroundStyle
              Styles.Header = StyleController.HeaderStyle
              Bands = <
                item
                  Width = 758
                end>
              object CD078DBandedTableViewFACIITEM: TcxGridDBBandedColumn
                Caption = #35373#20633#38918#24207
                DataBinding.FieldName = 'FACIITEM'
                HeaderAlignmentHorz = taRightJustify
                Width = 37
                Position.BandIndex = 0
                Position.ColIndex = 0
                Position.RowIndex = 0
              end
              object CD078DBandedTableViewFACILTYPE: TcxGridDBBandedColumn
                Caption = #35373#20633#23660#24615
                DataBinding.FieldName = 'FACITYPE'
                PropertiesClassName = 'TcxImageComboBoxProperties'
                Properties.Alignment.Horz = taLeftJustify
                Properties.Items = <
                  item
                    Description = #35373#20633#26032#22686
                    ImageIndex = 0
                    Value = '1'
                  end
                  item
                    Description = #35373#20633#25563#25286
                    Value = '2'
                  end
                  item
                    Description = #35373#20633#32173#20462
                    Value = '3'
                  end
                  item
                    Description = #35373#20633#25286#38500
                    Value = '4'
                  end
                  item
                    Description = #26381#21209#35722#26356
                    Value = '5'
                  end>
                HeaderAlignmentVert = vaCenter
                Width = 108
                Position.BandIndex = 0
                Position.ColIndex = 2
                Position.RowIndex = 0
              end
              object CD078DBandedTableViewFACILNAME: TcxGridDBBandedColumn
                Caption = #35373#20633#38917#30446
                DataBinding.FieldName = 'FACINAME'
                HeaderAlignmentVert = vaCenter
                Width = 144
                Position.BandIndex = 0
                Position.ColIndex = 3
                Position.RowIndex = 0
              end
              object CD078DBandedTableViewSERVICETYPE: TcxGridDBBandedColumn
                Caption = #26381#21209#39006#21029
                DataBinding.FieldName = 'SERVICETYPE'
                HeaderAlignmentHorz = taCenter
                Width = 38
                Position.BandIndex = 0
                Position.ColIndex = 1
                Position.RowIndex = 0
              end
              object CD078DBandedTableViewMODELNAME: TcxGridDBBandedColumn
                Caption = #35373#20633#22411#34399
                DataBinding.FieldName = 'MODELNAME'
                HeaderAlignmentVert = vaCenter
                Width = 98
                Position.BandIndex = 0
                Position.ColIndex = 4
                Position.RowIndex = 0
              end
              object CD078DBandedTableViewBUYNAME: TcxGridDBBandedColumn
                Caption = #36023#36067#26041#24335
                DataBinding.FieldName = 'BUYNAME'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 56
                Position.BandIndex = 0
                Position.ColIndex = 5
                Position.RowIndex = 0
              end
              object CD078DBandedTableViewCMBAUDRATE: TcxGridDBBandedColumn
                Caption = 'CM'#36895#29575
                DataBinding.FieldName = 'CMBAUDRATE'
                HeaderAlignmentVert = vaCenter
                Width = 106
                Position.BandIndex = 0
                Position.ColIndex = 6
                Position.RowIndex = 0
              end
              object CD078DBandedTableViewFIXIPCOUNT: TcxGridDBBandedColumn
                Caption = #22266#23450'IP'#25976
                DataBinding.FieldName = 'FIXIPCOUNT'
                HeaderAlignmentHorz = taRightJustify
                HeaderAlignmentVert = vaCenter
                Width = 63
                Position.BandIndex = 0
                Position.ColIndex = 7
                Position.RowIndex = 0
              end
              object CD078DBandedTableViewDYNIPCOUNT: TcxGridDBBandedColumn
                Caption = #28014#21205'IP'#25976#30446
                DataBinding.FieldName = 'DYNIPCOUNT'
                HeaderAlignmentHorz = taRightJustify
                HeaderAlignmentVert = vaCenter
                Width = 66
                Position.BandIndex = 0
                Position.ColIndex = 8
                Position.RowIndex = 0
              end
            end
            object CD078DGridLevel: TcxGridLevel
              GridView = CD078DBandedTableView
            end
          end
        end
      end
      object Sheet2: TcxTabSheet
        Caption = #20778#24800#32068#21512#21443#25976
        object Panel4: TPanel
          Left = 0
          Top = 0
          Width = 760
          Height = 249
          Align = alClient
          BevelOuter = bvNone
          BorderWidth = 3
          TabOrder = 0
          object CD078AGrid: TcxGrid
            Left = 3
            Top = 3
            Width = 754
            Height = 243
            Align = alClient
            TabOrder = 0
            LookAndFeel.Kind = lfFlat
            object CD078ABandedTableView: TcxGridDBBandedTableView
              OnDblClick = BandedTableViewDblClick
              NavigatorButtons.ConfirmDelete = False
              OnCustomDrawCell = CD078ABandedTableViewCustomDrawCell
              DataController.DataSource = CD078ADataSource
              DataController.Summary.DefaultGroupSummaryItems = <>
              DataController.Summary.FooterSummaryItems = <>
              DataController.Summary.SummaryGroups = <>
              OptionsCustomize.ColumnFiltering = False
              OptionsData.Deleting = False
              OptionsData.DeletingConfirmation = False
              OptionsData.Editing = False
              OptionsData.Inserting = False
              OptionsSelection.CellSelect = False
              OptionsSelection.HideFocusRectOnExit = False
              OptionsView.CellEndEllipsis = True
              OptionsView.GroupByBox = False
              OptionsView.HeaderAutoHeight = True
              OptionsView.HeaderEndEllipsis = True
              OptionsView.Indicator = True
              OptionsView.BandHeaders = False
              Styles.Content = StyleController.GridBackgroundStyle
              Styles.Header = StyleController.HeaderStyle
              Bands = <
                item
                  Width = 945
                end>
              object CD078ABandedTableViewCITEMNAME: TcxGridDBBandedColumn
                Caption = #25910#36027#38917#30446
                DataBinding.FieldName = 'CITEMNAME'
                HeaderAlignmentVert = vaCenter
                Width = 175
                Position.BandIndex = 0
                Position.ColIndex = 1
                Position.RowIndex = 0
              end
              object CD078ABandedTableViewSERVICETYPE: TcxGridDBBandedColumn
                Caption = #26381#21209#39006#21029
                DataBinding.FieldName = 'SERVICETYPE'
                PropertiesClassName = 'TcxTextEditProperties'
                Properties.Alignment.Horz = taCenter
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 36
                Position.BandIndex = 0
                Position.ColIndex = 0
                Position.RowIndex = 0
              end
              object CD078ABandedTableViewBUNDLE: TcxGridDBBandedColumn
                Caption = #26159#21542#32129#32004
                DataBinding.FieldName = 'BUNDLE'
                PropertiesClassName = 'TcxCheckBoxProperties'
                Properties.Alignment = taCenter
                Properties.DisplayChecked = '1'
                Properties.DisplayUnchecked = '0'
                Properties.NullStyle = nssUnchecked
                Properties.ValueChecked = '1'
                Properties.ValueUnchecked = '0'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 35
                Position.BandIndex = 0
                Position.ColIndex = 2
                Position.RowIndex = 0
              end
              object CD078ABandedTableViewSTOPFLAG: TcxGridDBBandedColumn
                Caption = #26159#21542#20572#29992
                DataBinding.FieldName = 'STOPFLAG'
                PropertiesClassName = 'TcxCheckBoxProperties'
                Properties.DisplayChecked = '1'
                Properties.DisplayUnchecked = '0'
                Properties.NullStyle = nssUnchecked
                Properties.ValueChecked = '1'
                Properties.ValueUnchecked = '0'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 32
                Position.BandIndex = 0
                Position.ColIndex = 3
                Position.RowIndex = 0
              end
              object CD078ABandedTableViewBUNDLEMON: TcxGridDBBandedColumn
                Caption = #32129#32004#26376#25976
                DataBinding.FieldName = 'BUNDLEMON'
                PropertiesClassName = 'TcxTextEditProperties'
                Properties.Alignment.Horz = taCenter
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 40
                Position.BandIndex = 0
                Position.ColIndex = 4
                Position.RowIndex = 0
              end
              object CD078ABandedTableViewPENALTYPE: TcxGridDBBandedColumn
                Caption = #36949#32004#26178#35336#20729#20381#25818
                DataBinding.FieldName = 'PENALTYPE'
                PropertiesClassName = 'TcxTextEditProperties'
                Properties.Alignment.Horz = taLeftJustify
                OnGetDisplayText = CD078ABandedTableViewPENALTYPEGetDisplayText
                HeaderAlignmentVert = vaCenter
                Width = 90
                Position.BandIndex = 0
                Position.ColIndex = 7
                Position.RowIndex = 0
              end
              object CD078ABandedTableViewEXPIRETYPE: TcxGridDBBandedColumn
                Caption = #20778#24800#21040#26399#35336#20729#20381#25818
                DataBinding.FieldName = 'EXPIRETYPE'
                PropertiesClassName = 'TcxTextEditProperties'
                Properties.Alignment.Horz = taLeftJustify
                OnGetDisplayText = CD078ABandedTableViewEXPIRETYPEGetDisplayText
                HeaderAlignmentVert = vaCenter
                Width = 108
                Position.BandIndex = 0
                Position.ColIndex = 10
                Position.RowIndex = 0
              end
              object CD078ABandedTableViewDEPOSITATTR: TcxGridDBBandedColumn
                Caption = #20445#35657#37329#23660#24615
                DataBinding.FieldName = 'DEPOSITATTR'
                PropertiesClassName = 'TcxTextEditProperties'
                Properties.Alignment.Horz = taCenter
                OnGetDisplayText = CD078ABandedTableViewDEPOSITATTRGetDisplayText
                HeaderAlignmentVert = vaCenter
                Width = 72
                Position.BandIndex = 0
                Position.ColIndex = 8
                Position.RowIndex = 0
              end
              object CD078ABandedTableViewDEPOSITNAME: TcxGridDBBandedColumn
                Caption = #20445#35657#37329#25910#36027#38917#30446
                DataBinding.FieldName = 'DEPOSITNAME'
                HeaderAlignmentVert = vaCenter
                Width = 175
                Position.BandIndex = 0
                Position.ColIndex = 9
                Position.RowIndex = 0
              end
              object CD078ABandedTableViewDEPOSITAMT: TcxGridDBBandedColumn
                Caption = #20445#35657#37329#37329#38989
                DataBinding.FieldName = 'DEPOSITAMT'
                PropertiesClassName = 'TcxCurrencyEditProperties'
                Properties.DecimalPlaces = 0
                Properties.DisplayFormat = ',0;-,0'
                HeaderAlignmentHorz = taRightJustify
                HeaderAlignmentVert = vaCenter
                Width = 99
                Position.BandIndex = 0
                Position.ColIndex = 11
                Position.RowIndex = 0
              end
              object CD078ABandedTableViewPENAL: TcxGridDBBandedColumn
                Caption = #36949#32004#37329#35373#23450
                DataBinding.FieldName = 'PENAL'
                PropertiesClassName = 'TcxCheckBoxProperties'
                Properties.Alignment = taCenter
                Properties.DisplayChecked = '1'
                Properties.DisplayUnchecked = '0'
                Properties.NullStyle = nssUnchecked
                Properties.ValueChecked = '1'
                Properties.ValueUnchecked = '0'
                HeaderAlignmentHorz = taCenter
                Width = 43
                Position.BandIndex = 0
                Position.ColIndex = 5
                Position.RowIndex = 0
              end
              object CD078ABandedTableViewSIGN: TcxGridDBBandedColumn
                Caption = #27491#36000#38917
                DataBinding.FieldName = 'SIGN'
                Visible = False
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Position.BandIndex = 0
                Position.ColIndex = 12
                Position.RowIndex = 0
              end
              object CD078ABandedTableViewLongPayFlag: TcxGridDBBandedColumn
                Caption = #38263#32371#21029
                DataBinding.FieldName = 'LongPayFlag'
                PropertiesClassName = 'TcxCheckBoxProperties'
                Properties.NullStyle = nssUnchecked
                Properties.ValueChecked = 1
                Properties.ValueUnchecked = 0
                GroupSummaryAlignment = taCenter
                HeaderAlignmentHorz = taCenter
                Width = 40
                Position.BandIndex = 0
                Position.ColIndex = 6
                Position.RowIndex = 0
              end
            end
            object CD078AGridLevel: TcxGridLevel
              GridView = CD078ABandedTableView
            end
          end
        end
      end
      object Sheet4: TcxTabSheet
        Caption = #26045#24037'/'#35373#23450'/'#35373#20633#20043#36027#29992#21443#25976
        object Panel6: TPanel
          Left = 0
          Top = 0
          Width = 760
          Height = 249
          Align = alClient
          BevelOuter = bvNone
          BorderWidth = 3
          TabOrder = 0
          object CD078CGrid: TcxGrid
            Left = 3
            Top = 3
            Width = 754
            Height = 243
            Align = alClient
            TabOrder = 0
            LookAndFeel.Kind = lfFlat
            object CD078CBandedTableView: TcxGridDBBandedTableView
              OnDblClick = BandedTableViewDblClick
              NavigatorButtons.ConfirmDelete = False
              OnCustomDrawCell = CD078CBandedTableViewCustomDrawCell
              DataController.DataSource = CD078CDataSource
              DataController.Summary.DefaultGroupSummaryItems = <>
              DataController.Summary.FooterSummaryItems = <>
              DataController.Summary.SummaryGroups = <>
              OptionsCustomize.ColumnFiltering = False
              OptionsData.Deleting = False
              OptionsData.DeletingConfirmation = False
              OptionsData.Editing = False
              OptionsData.Inserting = False
              OptionsSelection.CellSelect = False
              OptionsSelection.HideFocusRectOnExit = False
              OptionsView.CellEndEllipsis = True
              OptionsView.GroupByBox = False
              OptionsView.HeaderAutoHeight = True
              OptionsView.HeaderEndEllipsis = True
              OptionsView.Indicator = True
              OptionsView.BandHeaders = False
              Styles.Content = StyleController.GridBackgroundStyle
              Styles.Header = StyleController.HeaderStyle
              Bands = <
                item
                  Width = 757
                end>
              object CD078CBandedTableViewCITEMNAME: TcxGridDBBandedColumn
                Caption = #25910#36027#38917#30446
                DataBinding.FieldName = 'CITEMNAME'
                HeaderAlignmentVert = vaCenter
                Width = 175
                Position.BandIndex = 0
                Position.ColIndex = 1
                Position.RowIndex = 0
              end
              object CD078CBandedTableViewSERVICETYPE: TcxGridDBBandedColumn
                Caption = #26381#21209#39006#21029
                DataBinding.FieldName = 'SERVICETYPE'
                Width = 39
                Position.BandIndex = 0
                Position.ColIndex = 0
                Position.RowIndex = 0
              end
              object CD078CBandedTableViewFACIITEM: TcxGridDBBandedColumn
                Caption = #23565#25033#35373#20633
                DataBinding.FieldName = 'FACIITEM'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 40
                Position.BandIndex = 0
                Position.ColIndex = 6
                Position.RowIndex = 0
              end
              object CD078CBandedTableViewAMOUNT: TcxGridDBBandedColumn
                Caption = #37559#21806#29260#20729
                DataBinding.FieldName = 'AMOUNT'
                PropertiesClassName = 'TcxCurrencyEditProperties'
                Properties.DecimalPlaces = 0
                Properties.DisplayFormat = ',0;-,0'
                HeaderAlignmentHorz = taRightJustify
                HeaderAlignmentVert = vaCenter
                Width = 98
                Position.BandIndex = 0
                Position.ColIndex = 2
                Position.RowIndex = 0
              end
              object CD078CBandedTableViewDISCOUNTAMT: TcxGridDBBandedColumn
                Caption = #20778#24800#37329#38989
                DataBinding.FieldName = 'DISCOUNTAMT'
                PropertiesClassName = 'TcxCurrencyEditProperties'
                Properties.DecimalPlaces = 0
                Properties.DisplayFormat = ',0;-,0'
                HeaderAlignmentHorz = taRightJustify
                HeaderAlignmentVert = vaCenter
                Width = 101
                Position.BandIndex = 0
                Position.ColIndex = 3
                Position.RowIndex = 0
              end
              object CD078CBandedTableViewPUNISH: TcxGridDBBandedColumn
                Caption = #36949#32004#34389#20221
                DataBinding.FieldName = 'PUNISH'
                PropertiesClassName = 'TcxCheckBoxProperties'
                Properties.DisplayChecked = '1'
                Properties.DisplayUnchecked = '0'
                Properties.NullStyle = nssUnchecked
                Properties.ValueChecked = '1'
                Properties.ValueUnchecked = '0'
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 39
                Position.BandIndex = 0
                Position.ColIndex = 4
                Position.RowIndex = 0
              end
              object CD078CBandedTableViewPENALTYPE: TcxGridDBBandedColumn
                Caption = #36949#32004#35336#20729#20381#25818
                DataBinding.FieldName = 'PENALTYPE'
                PropertiesClassName = 'TcxImageComboBoxProperties'
                Properties.Items = <
                  item
                    Description = #22238#28335#29260#20729'('#20381#20840#38989')'
                    ImageIndex = 0
                    Value = '1'
                  end
                  item
                    Description = #22238#28335#29260#20729'('#20381#26410#20351#29992#26085#25976#27604#20363')'
                    Value = '2'
                  end>
                HeaderAlignmentVert = vaCenter
                Width = 155
                Position.BandIndex = 0
                Position.ColIndex = 5
                Position.RowIndex = 0
              end
              object CD078CBandedTableViewINSTCODESTR: TcxGridDBBandedColumn
                Caption = #25351#23450#27966#24037#39006#21029
                DataBinding.FieldName = 'INSTCODESTR'
                HeaderAlignmentVert = vaCenter
                Width = 110
                Position.BandIndex = 0
                Position.ColIndex = 7
                Position.RowIndex = 0
              end
            end
            object CD078CGridLevel: TcxGridLevel
              GridView = CD078CBandedTableView
            end
          end
        end
      end
      object Sheet5: TcxTabSheet
        Caption = #40670#25976#35373#23450
        ImageIndex = 4
        object Panel7: TPanel
          Left = 0
          Top = 0
          Width = 760
          Height = 249
          Align = alClient
          BevelOuter = bvNone
          BorderWidth = 3
          TabOrder = 0
          object cxGrid1: TcxGrid
            Left = 3
            Top = 3
            Width = 754
            Height = 243
            Align = alClient
            TabOrder = 0
            LookAndFeel.Kind = lfFlat
            object CD078GBandedTableView: TcxGridDBBandedTableView
              OnDblClick = CD078GBandedTableViewDblClick
              NavigatorButtons.ConfirmDelete = False
              DataController.DataSource = CD078GDataSource
              DataController.Summary.DefaultGroupSummaryItems = <>
              DataController.Summary.FooterSummaryItems = <>
              DataController.Summary.SummaryGroups = <>
              OptionsCustomize.ColumnFiltering = False
              OptionsData.Deleting = False
              OptionsData.DeletingConfirmation = False
              OptionsData.Editing = False
              OptionsData.Inserting = False
              OptionsSelection.CellSelect = False
              OptionsSelection.HideFocusRectOnExit = False
              OptionsView.CellEndEllipsis = True
              OptionsView.GroupByBox = False
              OptionsView.HeaderAutoHeight = True
              OptionsView.HeaderEndEllipsis = True
              OptionsView.Indicator = True
              OptionsView.BandHeaders = False
              Styles.Content = StyleController.GridBackgroundStyle
              Styles.Header = StyleController.HeaderStyle
              Bands = <
                item
                  Width = 945
                end>
              object CD078GBandedTableViewCITEMNAME: TcxGridDBBandedColumn
                Caption = #25910#36027#38917#30446
                DataBinding.FieldName = 'CITEMNAME'
                HeaderAlignmentVert = vaCenter
                Width = 124
                Position.BandIndex = 0
                Position.ColIndex = 1
                Position.RowIndex = 0
              end
              object CD078GBandedTableViewSERVICETYPE: TcxGridDBBandedColumn
                Caption = #26381#21209#39006#21029
                DataBinding.FieldName = 'SERVICETYPE'
                PropertiesClassName = 'TcxTextEditProperties'
                Properties.Alignment.Horz = taCenter
                HeaderAlignmentHorz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 134
                Position.BandIndex = 0
                Position.ColIndex = 0
                Position.RowIndex = 0
              end
              object CD078GBandedTableViewFACIITEM: TcxGridDBBandedColumn
                Caption = #35373#20633#38918#24207
                DataBinding.FieldName = 'FACIITEM'
                PropertiesClassName = 'TcxTextEditProperties'
                Properties.Alignment.Horz = taCenter
                HeaderAlignmentVert = vaCenter
                Width = 58
                Position.BandIndex = 0
                Position.ColIndex = 2
                Position.RowIndex = 0
              end
              object CD078GBandedTableViewSALEPOINTNAME: TcxGridDBBandedColumn
                Caption = #40670#25976#35373#23450#21517#31281
                DataBinding.FieldName = 'SALEPOINTNAME'
                HeaderAlignmentVert = vaCenter
                Width = 629
                Position.BandIndex = 0
                Position.ColIndex = 3
                Position.RowIndex = 0
              end
            end
            object cxGridLevel1: TcxGridLevel
              GridView = CD078GBandedTableView
            end
          end
        end
      end
      object Sheet6: TcxTabSheet
        Caption = 'PVOD'#35373#23450
        ImageIndex = 5
        object CD078IGrid: TcxGrid
          Left = 0
          Top = 0
          Width = 760
          Height = 249
          Align = alClient
          TabOrder = 0
          LookAndFeel.Kind = lfFlat
          object CD078IBandedTableView: TcxGridDBBandedTableView
            OnDblClick = BandedTableViewDblClick
            NavigatorButtons.ConfirmDelete = False
            OnCustomDrawCell = CD078IBandedTableViewCustomDrawCell
            DataController.DataSource = SO562DataSource
            DataController.Summary.DefaultGroupSummaryItems = <>
            DataController.Summary.FooterSummaryItems = <>
            DataController.Summary.SummaryGroups = <>
            OptionsBehavior.GoToNextCellOnEnter = True
            OptionsCustomize.ColumnFiltering = False
            OptionsData.CancelOnExit = False
            OptionsData.Deleting = False
            OptionsData.DeletingConfirmation = False
            OptionsData.Editing = False
            OptionsData.Inserting = False
            OptionsSelection.CellSelect = False
            OptionsSelection.HideFocusRectOnExit = False
            OptionsView.CellEndEllipsis = True
            OptionsView.ShowEditButtons = gsebAlways
            OptionsView.GroupByBox = False
            OptionsView.HeaderAutoHeight = True
            OptionsView.HeaderEndEllipsis = True
            OptionsView.Indicator = True
            OptionsView.BandHeaders = False
            Styles.Content = StyleController.GridBackgroundStyle
            Styles.Header = StyleController.HeaderStyle
            Bands = <
              item
                Width = 1800
              end>
            object cxGridDBBandedColumn1: TcxGridDBBandedColumn
              Caption = #36984#21462
              DataBinding.FieldName = 'CHOICE'
              PropertiesClassName = 'TcxCheckBoxProperties'
              Properties.NullStyle = nssUnchecked
              Properties.ValueChecked = 1
              Properties.ValueUnchecked = 0
              Visible = False
              HeaderAlignmentHorz = taCenter
              HeaderAlignmentVert = vaCenter
              Width = 56
              Position.BandIndex = 0
              Position.ColIndex = 0
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn2: TcxGridDBBandedColumn
              Caption = #29986#21697#31278#39006
              DataBinding.FieldName = 'TYPE'
              PropertiesClassName = 'TcxTextEditProperties'
              Properties.Alignment.Horz = taCenter
              Properties.Alignment.Vert = taVCenter
              OnGetDisplayText = cxGridDBBandedColumn2GetDisplayText
              HeaderAlignmentHorz = taCenter
              HeaderAlignmentVert = vaCenter
              Width = 74
              Position.BandIndex = 0
              Position.ColIndex = 1
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn3: TcxGridDBBandedColumn
              Caption = #29986#21697#29376#24907
              DataBinding.FieldName = 'ENDPURCHASE'
              PropertiesClassName = 'TcxTextEditProperties'
              Properties.Alignment.Horz = taCenter
              OnGetDisplayText = cxGridDBBandedColumn3GetDisplayText
              HeaderAlignmentHorz = taCenter
              HeaderAlignmentVert = vaCenter
              Width = 73
              Position.BandIndex = 0
              Position.ColIndex = 2
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn4: TcxGridDBBandedColumn
              Caption = #29986#21697#21517#31281
              DataBinding.FieldName = 'SHOWTITLE'
              HeaderAlignmentVert = vaCenter
              Width = 271
              Position.BandIndex = 0
              Position.ColIndex = 3
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn5: TcxGridDBBandedColumn
              Caption = #35330#36092#30908
              DataBinding.FieldName = 'CASID'
              PropertiesClassName = 'TcxTextEditProperties'
              Properties.Alignment.Horz = taLeftJustify
              Properties.Alignment.Vert = taVCenter
              HeaderAlignmentVert = vaCenter
              Width = 69
              Position.BandIndex = 0
              Position.ColIndex = 4
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn6: TcxGridDBBandedColumn
              Caption = #29260#20729
              DataBinding.FieldName = 'AMOUNT'
              PropertiesClassName = 'TcxTextEditProperties'
              Properties.Alignment.Horz = taRightJustify
              Properties.Alignment.Vert = taVCenter
              HeaderAlignmentHorz = taRightJustify
              HeaderAlignmentVert = vaCenter
              Width = 73
              Position.BandIndex = 0
              Position.ColIndex = 5
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn7: TcxGridDBBandedColumn
              Caption = #36000#38917#20778#24800#38917#30446
              DataBinding.FieldName = 'CITEMNAME2'
              OnGetDisplayText = cxGridDBBandedColumn7GetDisplayText
              HeaderAlignmentVert = vaCenter
              Width = 174
              Position.BandIndex = 0
              Position.ColIndex = 7
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn8: TcxGridDBBandedColumn
              Caption = #20778#24800#37329#38989
              DataBinding.FieldName = 'DISCOUNTAMT'
              PropertiesClassName = 'TcxTextEditProperties'
              Properties.Alignment.Horz = taRightJustify
              Properties.Alignment.Vert = taVCenter
              HeaderAlignmentHorz = taRightJustify
              HeaderAlignmentVert = vaCenter
              Width = 89
              Position.BandIndex = 0
              Position.ColIndex = 6
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn9: TcxGridDBBandedColumn
              Caption = #29986#21697#26377#25928#36215#22987#26085#26399
              DataBinding.FieldName = 'STARTVALIDITY'
              PropertiesClassName = 'TcxTextEditProperties'
              Properties.Alignment.Vert = taVCenter
              HeaderAlignmentVert = vaCenter
              Width = 158
              Position.BandIndex = 0
              Position.ColIndex = 8
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn10: TcxGridDBBandedColumn
              Caption = #29986#21697#26377#25928#32066#27490#26085#26399
              DataBinding.FieldName = 'ENDVALIDITY'
              PropertiesClassName = 'TcxTextEditProperties'
              Properties.Alignment.Vert = taVCenter
              HeaderAlignmentVert = vaCenter
              Width = 167
              Position.BandIndex = 0
              Position.ColIndex = 9
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn11: TcxGridDBBandedColumn
              Caption = #35330#36092#26377#25928#36215#22987#26085#26399
              DataBinding.FieldName = 'STARTPURCHASE'
              PropertiesClassName = 'TcxTextEditProperties'
              Properties.Alignment.Vert = taVCenter
              HeaderAlignmentVert = vaCenter
              Width = 176
              Position.BandIndex = 0
              Position.ColIndex = 10
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn12: TcxGridDBBandedColumn
              Caption = #35330#36092#26377#25928#32066#27490#26085#26399
              DataBinding.FieldName = 'ENDPURCHASE'
              PropertiesClassName = 'TcxTextEditProperties'
              Properties.Alignment.Vert = taVCenter
              HeaderAlignmentVert = vaCenter
              Width = 192
              Position.BandIndex = 0
              Position.ColIndex = 11
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn13: TcxGridDBBandedColumn
              Caption = #19979#26550#26085#26399
              DataBinding.FieldName = 'ONSALESTOPDATE'
              PropertiesClassName = 'TcxLabelProperties'
              Properties.Alignment.Horz = taLeftJustify
              Properties.Alignment.Vert = taVCenter
              HeaderAlignmentVert = vaCenter
              Width = 228
              Position.BandIndex = 0
              Position.ColIndex = 12
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn14: TcxGridDBBandedColumn
              DataBinding.FieldName = 'FLAG'
              Visible = False
              SortIndex = 0
              SortOrder = soAscending
              Position.BandIndex = 0
              Position.ColIndex = 13
              Position.RowIndex = 0
            end
          end
          object cxGridLevel2: TcxGridLevel
            GridView = CD078IBandedTableView
          end
        end
      end
      object Sheet7: TcxTabSheet
        Caption = #29986#21697#35373#23450
        ImageIndex = 6
        object CD078JGrid: TcxGrid
          Left = 0
          Top = 0
          Width = 760
          Height = 249
          Align = alClient
          TabOrder = 0
          LookAndFeel.Kind = lfFlat
          object CD078JBandedTableView: TcxGridDBBandedTableView
            OnDblClick = BandedTableViewDblClick
            NavigatorButtons.ConfirmDelete = False
            DataController.DataSource = CD078JDataSource
            DataController.Summary.DefaultGroupSummaryItems = <>
            DataController.Summary.FooterSummaryItems = <>
            DataController.Summary.SummaryGroups = <>
            OptionsBehavior.GoToNextCellOnEnter = True
            OptionsCustomize.ColumnFiltering = False
            OptionsData.CancelOnExit = False
            OptionsData.Deleting = False
            OptionsData.DeletingConfirmation = False
            OptionsData.Editing = False
            OptionsData.Inserting = False
            OptionsSelection.CellSelect = False
            OptionsSelection.HideFocusRectOnExit = False
            OptionsView.CellEndEllipsis = True
            OptionsView.ShowEditButtons = gsebAlways
            OptionsView.GroupByBox = False
            OptionsView.HeaderAutoHeight = True
            OptionsView.HeaderEndEllipsis = True
            OptionsView.Indicator = True
            OptionsView.BandHeaders = False
            Styles.Content = StyleController.GridBackgroundStyle
            Styles.Header = StyleController.HeaderStyle
            Bands = <
              item
                Width = 900
              end>
            object cxGridDBBandedColumn16: TcxGridDBBandedColumn
              Caption = #29986#21697#20195#30908
              DataBinding.FieldName = 'PRODUCTCODE'
              PropertiesClassName = 'TcxTextEditProperties'
              Properties.Alignment.Horz = taCenter
              Properties.Alignment.Vert = taVCenter
              HeaderAlignmentHorz = taCenter
              HeaderAlignmentVert = vaCenter
              Width = 85
              Position.BandIndex = 0
              Position.ColIndex = 0
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn18: TcxGridDBBandedColumn
              Caption = #29986#21697#21517#31281
              DataBinding.FieldName = 'PRODUCTNAME'
              HeaderAlignmentVert = vaCenter
              Width = 402
              Position.BandIndex = 0
              Position.ColIndex = 1
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn20: TcxGridDBBandedColumn
              Caption = #26381#21209#39006#21029
              DataBinding.FieldName = 'SERVICETYPE'
              PropertiesClassName = 'TcxTextEditProperties'
              Properties.Alignment.Horz = taRightJustify
              Properties.Alignment.Vert = taVCenter
              HeaderAlignmentVert = vaCenter
              Width = 59
              Position.BandIndex = 0
              Position.ColIndex = 2
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn21: TcxGridDBBandedColumn
              Caption = #35373#20633#38918#24207
              DataBinding.FieldName = 'FACIITEM'
              OnGetDisplayText = cxGridDBBandedColumn7GetDisplayText
              HeaderAlignmentVert = vaCenter
              Width = 63
              Position.BandIndex = 0
              Position.ColIndex = 3
              Position.RowIndex = 0
            end
            object cxGridDBBandedColumn23: TcxGridDBBandedColumn
              Caption = #25351#23450#27966#24037#39006#21029
              DataBinding.FieldName = 'INSTCODESTR'
              PropertiesClassName = 'TcxTextEditProperties'
              Properties.Alignment.Vert = taVCenter
              HeaderAlignmentVert = vaCenter
              Width = 291
              Position.BandIndex = 0
              Position.ColIndex = 4
              Position.RowIndex = 0
            end
          end
          object cxGridLevel3: TcxGridLevel
            GridView = CD078JBandedTableView
          end
        end
      end
    end
  end
  object ScrollBox3: TScrollBox
    Left = 0
    Top = 0
    Width = 764
    Height = 249
    VertScrollBar.Visible = False
    Align = alTop
    AutoScroll = False
    BevelInner = bvNone
    BevelKind = bkFlat
    BorderStyle = bsNone
    TabOrder = 2
    DesignSize = (
      762
      247)
    object Label1: TLabel
      Left = 13
      Top = 12
      Width = 72
      Height = 14
      Caption = #20778#24800#32068#21512#20195#30908
    end
    object Label2: TLabel
      Left = 185
      Top = 11
      Width = 72
      Height = 14
      Caption = #20778#24800#32068#21512#21517#31281
    end
    object Label3: TLabel
      Left = 13
      Top = 42
      Width = 72
      Height = 14
      Caption = #29986#21697#19978#26550#26399#38291
    end
    object Label4: TLabel
      Left = 174
      Top = 43
      Width = 12
      Height = 14
      Caption = #33267
    end
    object Label5: TLabel
      Left = 688
      Top = 43
      Width = 36
      Height = 14
      Caption = #21443#32771#34399
    end
    object Shape1: TShape
      Left = 13
      Top = 123
      Width = 735
      Height = 2
      Anchors = [akLeft, akTop, akRight]
      Brush.Style = bsClear
      Pen.Color = clGray
    end
    object Label6: TLabel
      Left = 285
      Top = 43
      Width = 48
      Height = 14
      Caption = #35336#36027#27231#21046
    end
    object Label7: TLabel
      Left = 115
      Top = 72
      Width = 96
      Height = 14
      Caption = #36899#21205#20778#24800#32068#21512#20195#30908
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object lbl1: TLabel
      Left = 63
      Top = 100
      Width = 21
      Height = 14
      Caption = 'URL'
    end
    object lbl2: TLabel
      Left = 602
      Top = 99
      Width = 46
      Height = 14
      Caption = 'EPG'#38918#24207
    end
    object lbl3: TLabel
      Left = 601
      Top = 71
      Width = 48
      Height = 14
      Caption = #36969#29992#23565#35937
    end
    object edtCodeNo: TcxTextEdit
      Left = 88
      Top = 8
      ParentFont = False
      Properties.MaxLength = 10
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 0
      Width = 75
    end
    object edtDescription: TcxTextEdit
      Left = 265
      Top = 8
      ParentFont = False
      Properties.MaxLength = 50
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 1
      Width = 487
    end
    object edtDateSt: TcxMaskEdit
      Left = 88
      Top = 39
      ParentFont = False
      Properties.MaskKind = emkRegExprEx
      Properties.EditMask = 
        '[1-9][0-9][0-9][0-9] / (0?[1-9] | 1[012]) / ([012]?[1-9] | [123]' +
        '0 |31)'
      Properties.MaxLength = 0
      Properties.OnValidate = edtDateStPropertiesValidate
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 2
      Width = 77
    end
    object edtDateEd: TcxMaskEdit
      Left = 193
      Top = 39
      ParentFont = False
      Properties.MaskKind = emkRegExprEx
      Properties.EditMask = 
        '[1-9][0-9][0-9][0-9] / (0?[1-9] | 1[012]) / ([012]?[1-9] | [123]' +
        '0 |31)'
      Properties.MaxLength = 0
      Properties.OnValidate = edtDateStPropertiesValidate
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 3
      Width = 82
    end
    object chkMasterProd: TcxCheckBox
      Left = 548
      Top = 40
      Caption = #20027#25512#21830#21697
      Style.StyleController = StyleController.CheckBoxStyle
      TabOrder = 6
      Width = 74
    end
    object chkStopFlag: TcxCheckBox
      Left = 617
      Top = 40
      Caption = #26159#21542#20572#29992
      Style.StyleController = StyleController.CheckBoxStyle
      TabOrder = 7
      Width = 74
    end
    object edtRefNo: TcxTextEdit
      Left = 731
      Top = 39
      ParentFont = False
      Properties.MaxLength = 3
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 8
      Width = 21
    end
    object MemoPageControl: TcxPageControl
      Left = 11
      Top = 128
      Width = 750
      Height = 109
      ActivePage = cxTabSheet1
      Style = 1
      TabOrder = 14
      ClientRectBottom = 107
      ClientRectLeft = 2
      ClientRectRight = 748
      ClientRectTop = 23
      object cxTabSheet1: TcxTabSheet
        Caption = #20778#24800#32068#21512#35498#26126
        ImageIndex = 0
        object edtBundleProdNote: TcxMemo
          Left = 0
          Top = 0
          Align = alClient
          ParentFont = False
          Properties.MaxLength = 4000
          Properties.ScrollBars = ssVertical
          Properties.WordWrap = False
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 0
          Height = 84
          Width = 746
        end
      end
      object cxTabSheet2: TcxTabSheet
        Caption = #24037#21934#21015#21360#35498#26126
        ImageIndex = 1
        object edtWorkNote: TcxMemo
          Left = 0
          Top = 0
          Align = alClient
          ParentFont = False
          Properties.MaxLength = 4000
          Properties.ScrollBars = ssVertical
          Properties.WordWrap = False
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 0
          Height = 84
          Width = 746
        end
      end
      object cxTabSheet3: TcxTabSheet
        Caption = #27298#38468#26781#20214
        ImageIndex = 2
        object edtConditional: TcxMemo
          Left = 0
          Top = 0
          Align = alClient
          ParentFont = False
          Properties.MaxLength = 4000
          Properties.ScrollBars = ssVertical
          Properties.WordWrap = False
          Style.StyleController = StyleController.EditorStyle
          TabOrder = 0
          Height = 84
          Width = 746
        end
      end
    end
    object edtKindFunction: TcxComboBox
      Left = 339
      Top = 39
      ParentFont = False
      Properties.DropDownListStyle = lsFixedList
      Properties.Items.Strings = (
        '0.'#29986#21697'Bundle'
        '1.'#29986#21697#38568#36984)
      Properties.OnChange = edtKindFunctionPropertiesChange
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 4
      Width = 108
    end
    object chkChooseKind: TcxCheckBox
      Left = 451
      Top = 40
      Caption = #38971#36947#25361#36984#27169#24335
      Style.StyleController = StyleController.CheckBoxStyle
      TabOrder = 5
      Width = 97
    end
    object chkPayKind: TcxCheckBox
      Left = 9
      Top = 69
      Caption = #21487#20351#29992#29694#20184#21046
      Style.StyleController = StyleController.CheckBoxStyle
      TabOrder = 9
      Width = 97
    end
    inline SyncBPCode: TDataLookup
      Left = 218
      Top = 66
      Width = 381
      Height = 23
      HorzScrollBar.Visible = False
      VertScrollBar.Visible = False
      TabOrder = 10
      inherited CodeNo: TcxTextEdit
        Properties.CharCase = ecUpperCase
        Properties.OnChange = SyncBPCodeCodeNoPropertiesChange
        Width = 79
      end
      inherited CodeName: TcxLookupComboBox
        Left = 82
        Properties.OnChange = SyncBPCodeCodeNamePropertiesChange
        Properties.OnInitPopup = SyncBPCodeCodeNamePropertiesInitPopup
        Width = 295
      end
    end
    object edtImgURL: TcxTextEdit
      Left = 89
      Top = 95
      ParentFont = False
      Properties.MaxLength = 50
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 12
      Width = 504
    end
    object edtEPGOrder: TcxTextEdit
      Left = 654
      Top = 94
      ParentFont = False
      Properties.MaxLength = 4
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 13
      OnKeyPress = edtEPGOrderKeyPress
      Width = 46
    end
    object cobCanUseType: TcxComboBox
      Left = 653
      Top = 65
      ParentFont = False
      Properties.DropDownListStyle = lsFixedList
      Properties.Items.Strings = (
        '0.'#19981#20998
        '1.CC&B'
        '2.EPG')
      Properties.OnChange = cobCanUseTypePropertiesChange
      Properties.OnEditValueChanged = cobCanUseTypePropertiesEditValueChanged
      Style.StyleController = StyleController.EditorStyle
      TabOrder = 11
      Text = '0.'#19981#20998
      Width = 100
    end
    object chkByHouse: TcxCheckBox
      Left = 702
      Top = 94
      Caption = 'By '#25142
      Style.StyleController = StyleController.CheckBoxStyle
      TabOrder = 15
      Width = 60
    end
  end
  object ActionList1: TActionList
    Left = 32
    Top = 450
    object actSave: TAction
      Caption = 'F2.'#23384#27284
      ShortCut = 113
      OnExecute = actSaveExecute
    end
    object actInsert: TAction
      Caption = 'F6.'#26032#22686
      ShortCut = 117
      OnExecute = actInsertExecute
    end
    object actUpdate: TAction
      Caption = 'F11.'#20462#25913
      ShortCut = 122
      OnExecute = actUpdateExecute
    end
    object actDelete: TAction
      Caption = 'F10.'#21034#38500
      ShortCut = 121
      OnExecute = actDeleteExecute
    end
    object actCancel: TAction
      Caption = #21462#28040'(&C)'
      ShortCut = 32835
      OnExecute = actCancelExecute
    end
    object actClose: TAction
      Caption = #32080#26463'(&X)'
      ShortCut = 32856
      OnExecute = actCloseExecute
    end
  end
  object CD078ADataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 76
    Top = 376
  end
  object CD078ADataSource: TDataSource
    DataSet = CD078ADataSet
    Left = 113
    Top = 376
  end
  object CD078BDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 76
    Top = 412
  end
  object CD078BDataSource: TDataSource
    DataSet = CD078BDataSet
    Left = 113
    Top = 412
  end
  object CD078DDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    AfterPost = DataSetAfterDeleteOrPost
    AfterDelete = DataSetAfterDeleteOrPost
    Left = 76
    Top = 340
  end
  object CD078DDataSource: TDataSource
    DataSet = CD078DDataSet
    Left = 113
    Top = 340
  end
  object CD078CDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 76
    Top = 446
  end
  object CD078CDataSource: TDataSource
    DataSet = CD078CDataSet
    Left = 113
    Top = 446
  end
  object CD078DataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    AfterScroll = CD078DataSetAfterScroll
    Left = 76
    Top = 267
  end
  object CD078Provider: TDataSetProvider
    DataSet = DBController.DataReader
    Options = [poAllowCommandText]
    Left = 108
    Top = 267
  end
  object OtherDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 529
    Top = 369
  end
  object InstCodeDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 564
    Top = 369
  end
  object CD078A1DataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 76
    Top = 304
  end
  object CD078A1CloneDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 494
    Top = 369
  end
  object CD078A3DataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 152
    Top = 304
  end
  object CD078A3CloneDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 598
    Top = 369
  end
  object CD078A2DataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 112
    Top = 304
  end
  object CD078A2CloneDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 668
    Top = 369
  end
  object CD078B1CloneDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 703
    Top = 369
  end
  object CD078B1DataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 148
    Top = 412
  end
  object CD078A4DataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 192
    Top = 305
  end
  object CD078A4CloneDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 633
    Top = 369
  end
  object MCItemDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 460
    Top = 369
  end
  object LinkKeyDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 460
    Top = 405
  end
  object CD078GDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 154
    Top = 342
  end
  object CD078G1DataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 222
    Top = 342
  end
  object CD078GDataSource: TDataSource
    DataSet = CD078GDataSet
    Left = 187
    Top = 344
  end
  object CD078G1CloneDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 732
    Top = 370
  end
  object CD078IDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 158
    Top = 444
  end
  object CD078ICloneDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 730
    Top = 338
  end
  object SO562DataSet: TClientDataSet
    Aggregates = <>
    FieldDefs = <>
    IndexDefs = <>
    Params = <>
    ProviderName = 'SO562Provider'
    StoreDefs = True
    Left = 298
    Top = 450
  end
  object SO562Provider: TDataSetProvider
    DataSet = DBController.ComReader
    Options = [poAllowCommandText]
    Left = 164
    Top = 269
  end
  object SO562DataSource: TDataSource
    DataSet = SO562DataSet
    Left = 267
    Top = 450
  end
  object SO562_1_DataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 328
    Top = 450
  end
  object CD078JDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 224
    Top = 394
  end
  object CD078JDataSource: TDataSource
    DataSet = CD078JDataSet
    Left = 255
    Top = 394
  end
  object CD078KDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 224
    Top = 426
  end
  object CD078LDataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078Provider'
    Left = 260
    Top = 427
  end
end
