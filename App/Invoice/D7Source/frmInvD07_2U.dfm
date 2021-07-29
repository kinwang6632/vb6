object frmInvD07_2: TfrmInvD07_2
  Left = 348
  Top = 315
  Width = 876
  Height = 306
  Caption = 'frmInvD07_2'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object pnl1: TPanel
    Left = 0
    Top = 0
    Width = 860
    Height = 231
    Align = alTop
    TabOrder = 0
    object grdCSV: TcxGrid
      Left = 1
      Top = 1
      Width = 858
      Height = 229
      Align = alClient
      TabOrder = 0
      object gvCSV: TcxGridDBTableView
        NavigatorButtons.ConfirmDelete = False
        FilterBox.CustomizeDialog = False
        DataController.DataSource = dsView
        DataController.Summary.DefaultGroupSummaryItems = <>
        DataController.Summary.FooterSummaryItems = <>
        DataController.Summary.SummaryGroups = <>
        OptionsCustomize.ColumnFiltering = False
        OptionsCustomize.ColumnSorting = False
        OptionsData.Deleting = False
        OptionsData.DeletingConfirmation = False
        OptionsData.Editing = False
        OptionsData.Inserting = False
        OptionsSelection.InvertSelect = False
        OptionsView.CellEndEllipsis = True
        OptionsView.GroupByBox = False
        OptionsView.HeaderEndEllipsis = True
        OptionsView.Indicator = True
        object gvCSVInvIsErr: TcxGridDBColumn
          Caption = ' '#26159#21542#37679#35492
          DataBinding.FieldName = 'INVISERR'
          PropertiesClassName = 'TcxCheckBoxProperties'
          Properties.NullStyle = nssUnchecked
          Properties.ValueChecked = 1
          Properties.ValueUnchecked = 0
          GroupSummaryAlignment = taCenter
          HeaderAlignmentHorz = taCenter
        end
        object gvCSVBussiness: TcxGridDBColumn
          Caption = #29151#26989#20154#32113#32232
          DataBinding.FieldName = 'BUSSINESSID'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Width = 98
        end
        object gvCSVInvSortID: TcxGridDBColumn
          Caption = #30332#31080#39006#21029#20195#34399
          DataBinding.FieldName = 'INVSORTID'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taRightJustify
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taRightJustify
          HeaderAlignmentVert = vaCenter
          Width = 93
        end
        object gvCSVInvType: TcxGridDBColumn
          Caption = #30332#31080#39006#21029
          DataBinding.FieldName = 'INVTYPE'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taLeftJustify
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentVert = vaCenter
          Width = 93
        end
        object gvCSVInvPeriod: TcxGridDBColumn
          Caption = #30332#31080#26399#21029
          DataBinding.FieldName = 'INVPERIOD'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
          Width = 186
        end
        object gvCSVInvHead: TcxGridDBColumn
          Caption = #30332#31080#23383#36556#21517#31281
          DataBinding.FieldName = 'INVHEAD'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentVert = vaCenter
          Width = 82
        end
        object gvCSVInvStart: TcxGridDBColumn
          Caption = #30332#31080#36215#34399
          DataBinding.FieldName = 'INVSTART'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentVert = vaCenter
          Width = 94
        end
        object gvCSVInvEnd: TcxGridDBColumn
          Caption = #30332#31080#36804#34399
          DataBinding.FieldName = 'INVEND'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Vert = taVCenter
          HeaderAlignmentVert = vaCenter
          Width = 90
        end
        object gvCSVInvUploadFlag: TcxGridDBColumn
          Caption = #19978#20659#35387#35352
          DataBinding.FieldName = 'INVUPLOADFLAG'
          PropertiesClassName = 'TcxCheckBoxProperties'
          Properties.Alignment = taCenter
          Properties.NullStyle = nssUnchecked
          Properties.ValueChecked = 1
          Properties.ValueUnchecked = 0
          HeaderAlignmentHorz = taCenter
          HeaderAlignmentVert = vaCenter
        end
        object gvCSVNotes: TcxGridDBColumn
          Caption = #20633#35387
          DataBinding.FieldName = 'NOTES'
          PropertiesClassName = 'TcxMemoProperties'
          Properties.Alignment = taLeftJustify
          HeaderAlignmentVert = vaCenter
          Width = 225
        end
        object gvCSVErrMsg: TcxGridDBColumn
          Caption = #37679#35492#35338#24687
          DataBinding.FieldName = 'ERRMSG'
          PropertiesClassName = 'TcxMemoProperties'
          HeaderAlignmentVert = vaCenter
          Width = 443
        end
      end
      object glCSV: TcxGridLevel
        GridView = gvCSV
      end
    end
  end
  object btnOK: TcxButton
    Left = 7
    Top = 236
    Width = 84
    Height = 29
    Cancel = True
    Caption = ' '#30906#23450
    Enabled = False
    TabOrder = 1
    OnClick = btnOKClick
    Glyph.Data = {
      36040000424D3604000000000000360000002800000010000000100000000100
      2000000000000004000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000004040410000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000003F3F3FFF3F3F3FFF00000000040404100000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000003F3F3FFF70E584FF70E584FF3F3F3FFF000000000404
      0410000000000000000000000000000000000000000000000000000000000000
      0000000000003F3F3FFF70E584FF70E584FF70E584FF70E584FF3F3F3FFF0000
      0000040404100000000000000000000000000000000000000000000000000000
      00003F3F3FFF70E584FF70E584FF70E584FF70E584FF70E584FF70E584FF3F3F
      3FFF000000000000000000000000000000000000000000000000000000003F3F
      3FFF70E584FF70E584FF70E584FF3F3F3FFF424943FF70E584FF70E584FF70E5
      84FF3F3F3FFF00000000000000000000000000000000000000003F3F3FFF70E5
      84FF70E584FF70E584FF3F3F3FFF00000000000000003F3F3FFF6DDB80FF70E5
      84FF70E584FF3F3F3FFF00000000000000000000000000000000000000003F3F
      3FFF70E584FF3F3F3FFF000000000404041000000000000000003F3F3FFF70E5
      84FF70E584FF70E584FF3F3F3FFF000000000000000000000000000000000000
      00003F3F3FFF0000000004040410000000000000000000000000000000003F3F
      3FFF70E584FF70E584FF70E584FF3F3F3FFF0000000000000000000000000000
      0000000000000404041000000000000000000000000000000000000000000000
      00003F3F3FFF70E584FF70E584FF70E584FF3F3F3FFF00000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000003F3F3FFF70E584FF70E584FF70E584FF3F3F3FFF000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000003F3F3FFF70E584FF3F3F3FFF04040410000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000003F3F3FFF0C0C0C3000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000}
  end
  object btnExit: TcxButton
    Left = 764
    Top = 236
    Width = 84
    Height = 29
    Cancel = True
    Caption = #32080#26463
    TabOrder = 2
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
  object dsCSV: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 562
    Top = 358
  end
  object dsView: TDataSource
    DataSet = dsCSV
    Left = 442
    Top = 358
  end
end
