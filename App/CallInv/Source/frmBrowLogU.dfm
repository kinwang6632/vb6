object frmLog: TfrmLog
  Left = 113
  Top = 179
  BorderStyle = bsDialog
  Caption = '['#20659#36865#36890#30693'] '#35352#37636#28687#28768
  ClientHeight = 268
  ClientWidth = 750
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 750
    Height = 268
    Align = alClient
    TabOrder = 0
    object cxGrid: TcxGrid
      Left = 1
      Top = 1
      Width = 748
      Height = 266
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
      object cxGridBrow: TcxGridDBTableView
        NavigatorButtons.ConfirmDelete = False
        FilterBox.CustomizeDialog = False
        DataController.DataSource = dsInvLog
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
        OptionsView.GroupByBox = False
        object cxGridBrowType: TcxGridDBColumn
          Caption = #30332#36865#23660#24615
          DataBinding.FieldName = 'TYPE'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          OnGetDisplayText = cxGridBrowTypeGetDisplayText
          HeaderAlignmentHorz = taCenter
          Options.Editing = False
          Width = 90
        end
        object cxGridBrowCURRENTSTATE: TcxGridDBColumn
          Caption = #26368#26032#26597#35426#29376#24907
          DataBinding.FieldName = 'CURRENTSTATENAME'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          HeaderAlignmentHorz = taCenter
          Width = 121
        end
        object cxGridBrowType2: TcxGridDBColumn
          Caption = #32879#32097#36039#35338
          DataBinding.FieldName = 'TYPE2'
          Width = 223
        end
        object cxGridBrowSendType: TcxGridDBColumn
          Caption = #38928#32004#30332#36865#26178#38291
          DataBinding.FieldName = 'SENDTIME'
          PropertiesClassName = 'TcxDateEditProperties'
          Properties.Alignment.Horz = taCenter
          HeaderAlignmentHorz = taCenter
          Width = 128
        end
        object cxGridBrowREADTIME: TcxGridDBColumn
          Caption = #23458#25142#35712#21462#26178#38291
          DataBinding.FieldName = 'READTIME'
          Width = 119
        end
        object cxGridBrowSendStats: TcxGridDBColumn
          Caption = #30332#36865#29376#24907
          DataBinding.FieldName = 'SENDSTATE'
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          OnGetDisplayText = cxGridBrowSendStatsGetDisplayText
          HeaderAlignmentHorz = taCenter
          Width = 92
        end
        object cxGridBrowSendCnt: TcxGridDBColumn
          Caption = #30332#36865#27425#25976
          DataBinding.FieldName = 'SENDCNT'
          HeaderAlignmentHorz = taCenter
          Width = 84
        end
        object cxGridBrowReturnCode: TcxGridDBColumn
          Caption = #22238#25033#32080#26524#20195#30908
          DataBinding.FieldName = 'RETURNCODE'
          HeaderAlignmentHorz = taCenter
          Width = 114
        end
        object cxGridBrowReturnMsg: TcxGridDBColumn
          Caption = #22238#25033#32080#26524#35338#24687
          DataBinding.FieldName = 'RETURNMSG'
          HeaderAlignmentHorz = taCenter
          Width = 403
        end
      end
      object cxGridLevel: TcxGridLevel
        GridView = cxGridBrow
      end
    end
  end
  object dsInvLog: TDataSource
    DataSet = dtmComm.adoLog
    Left = 48
    Top = 72
  end
end
