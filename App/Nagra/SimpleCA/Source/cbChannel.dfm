object fmChannel: TfmChannel
  Left = 478
  Top = 208
  BorderStyle = bsDialog
  Caption = #38971#36947#36984#25799
  ClientHeight = 389
  ClientWidth = 393
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 393
    Height = 41
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      393
      41)
    object Bevel1: TBevel
      Left = 0
      Top = 38
      Width = 393
      Height = 3
      Align = alBottom
      Shape = bsBottomLine
    end
    object btnOK: TButton
      Left = 114
      Top = 7
      Width = 75
      Height = 25
      Anchors = [akRight, akBottom]
      Caption = #30906#23450
      Default = True
      ModalResult = 1
      TabOrder = 0
    end
    object btnCancel: TButton
      Left = 198
      Top = 7
      Width = 75
      Height = 25
      Anchors = [akRight, akBottom]
      Caption = #21462#28040
      ModalResult = 2
      TabOrder = 1
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 41
    Width = 393
    Height = 348
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 1
    object Bevel2: TBevel
      Left = 0
      Top = 0
      Width = 393
      Height = 3
      Align = alTop
      Shape = bsSpacer
    end
    object Bevel3: TBevel
      Left = 0
      Top = 345
      Width = 393
      Height = 3
      Align = alBottom
      Shape = bsSpacer
    end
    object ChGrid: TcxGrid
      Left = 0
      Top = 3
      Width = 393
      Height = 342
      Align = alClient
      TabOrder = 0
      LookAndFeel.Kind = lfFlat
      object ChGridView: TcxGridDBTableView
        NavigatorButtons.ConfirmDelete = False
        DataController.DataSource = CD024DataSource
        DataController.KeyFieldNames = 'CODENO'
        DataController.Summary.DefaultGroupSummaryItems = <>
        DataController.Summary.FooterSummaryItems = <>
        DataController.Summary.SummaryGroups = <>
        OptionsBehavior.PullFocusing = True
        OptionsCustomize.ColumnFiltering = False
        OptionsData.Deleting = False
        OptionsData.Editing = False
        OptionsData.Inserting = False
        OptionsSelection.CellSelect = False
        OptionsSelection.MultiSelect = True
        OptionsView.GroupByBox = False
        OptionsView.Indicator = True
        object ViewCODENO: TcxGridDBColumn
          Caption = #20195#30908
          DataBinding.FieldName = 'CODENO'
          Width = 60
        end
        object ViewDESCRIPTION: TcxGridDBColumn
          Caption = #38971#36947#21517#31281
          DataBinding.FieldName = 'DESCRIPTION'
          Width = 159
        end
        object ViewCHANNELID: TcxGridDBColumn
          Caption = 'Product Id'
          DataBinding.FieldName = 'CHANNELID'
          Width = 112
        end
      end
      object ChGridLevel: TcxGridLevel
        GridView = ChGridView
      end
    end
  end
  object CD024Reader: TADOQuery
    CacheSize = 1000
    Connection = fmMain.ComConnection
    CursorType = ctStatic
    Parameters = <>
    Left = 256
    Top = 104
  end
  object CD024DataSource: TDataSource
    DataSet = CD024Reader
    Left = 288
    Top = 104
  end
end
