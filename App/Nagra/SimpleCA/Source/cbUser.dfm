object fmUser: TfmUser
  Left = 370
  Top = 107
  BorderStyle = bsDialog
  Caption = #20351#29992#32773#35373#23450
  ClientHeight = 517
  ClientWidth = 535
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
  DesignSize = (
    535
    517)
  PixelsPerInch = 96
  TextHeight = 14
  object Bevel1: TBevel
    Left = 8
    Top = 49
    Width = 514
    Height = 3
    Anchors = [akLeft, akTop, akRight]
    Shape = bsTopLine
  end
  object Label1: TLabel
    Left = 44
    Top = 64
    Width = 60
    Height = 14
    Caption = #20351#29992#32773#24115#34399
    Font.Charset = ANSI_CHARSET
    Font.Color = clRed
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label2: TLabel
    Left = 242
    Top = 63
    Width = 48
    Height = 14
    Caption = #30331#20837#23494#30908
    Font.Charset = ANSI_CHARSET
    Font.Color = clRed
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label3: TLabel
    Left = 44
    Top = 89
    Width = 60
    Height = 14
    Caption = #20351#29992#32773#22995#21517
    Font.Charset = ANSI_CHARSET
    Font.Color = clRed
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label4: TLabel
    Left = 8
    Top = 121
    Width = 96
    Height = 14
    Caption = #21487#30332#36865#25351#20196#31995#32113#21488
    Font.Charset = ANSI_CHARSET
    Font.Color = clRed
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label5: TLabel
    Left = 44
    Top = 225
    Width = 60
    Height = 14
    Caption = #20351#29992#21040#26399#26085
  end
  object Label6: TLabel
    Left = 227
    Top = 224
    Width = 84
    Height = 14
    Caption = #38283#38971#36947#38480#21046#27425#25976
  end
  object Label7: TLabel
    Left = 379
    Top = 224
    Width = 84
    Height = 14
    Caption = #38283#38971#36947#38480#21046#22825#25976
  end
  object btnInsert: TButton
    Left = 37
    Top = 8
    Width = 75
    Height = 25
    Caption = #26032#22686
    TabOrder = 0
    OnClick = btnInsertClick
  end
  object btnUpdate: TButton
    Left = 113
    Top = 8
    Width = 75
    Height = 25
    Caption = #20462#25913
    TabOrder = 1
    OnClick = btnUpdateClick
  end
  object btnDelete: TButton
    Left = 189
    Top = 8
    Width = 75
    Height = 25
    Caption = #21034#38500
    TabOrder = 2
    OnClick = btnDeleteClick
  end
  object btnCancel: TButton
    Left = 265
    Top = 8
    Width = 75
    Height = 25
    Caption = #21462#28040
    TabOrder = 3
    OnClick = btnCancelClick
  end
  object btnSave: TButton
    Left = 341
    Top = 8
    Width = 75
    Height = 25
    Caption = #23384#27284
    TabOrder = 4
    OnClick = btnSaveClick
  end
  object btnExit: TButton
    Left = 418
    Top = 8
    Width = 75
    Height = 25
    Caption = #32080#26463
    TabOrder = 5
    OnClick = btnExitClick
  end
  object txtUserId: TcxTextEdit
    Left = 114
    Top = 61
    ParentFont = False
    Properties.CharCase = ecUpperCase
    Properties.MaxLength = 10
    Style.StyleController = fmMain.cxEditStyleController1
    TabOrder = 6
    Width = 121
  end
  object txtPassword: TcxTextEdit
    Left = 294
    Top = 61
    ParentFont = False
    Properties.MaxLength = 10
    Style.StyleController = fmMain.cxEditStyleController1
    TabOrder = 7
    Width = 121
  end
  object txtUserName: TcxTextEdit
    Left = 114
    Top = 89
    ParentFont = False
    Properties.MaxLength = 20
    Style.StyleController = fmMain.cxEditStyleController1
    TabOrder = 8
    Width = 121
  end
  object txtUserLock: TcxCheckBox
    Left = 250
    Top = 89
    Caption = #37782#23450#24115#34399
    ParentFont = False
    Style.StyleController = fmMain.cxEditStyleController1
    TabOrder = 9
    Width = 81
  end
  object txtCompStr: TcxCheckListBox
    Left = 113
    Top = 119
    Width = 368
    Height = 94
    Columns = 4
    IntegralHeight = True
    Items = <
      item
        Text = '222'
      end
      item
        Text = '111'
      end
      item
        Text = '333'
      end
      item
        Text = '4444'
      end
      item
        Text = '5555'
      end
      item
      end
      item
      end
      item
      end
      item
      end
      item
      end
      item
      end
      item
      end
      item
      end
      item
      end
      item
      end
      item
      end
      item
      end>
    ParentFont = False
    Style.StyleController = fmMain.cxEditStyleController1
    TabOrder = 10
  end
  object txtExpireDate: TcxDateEdit
    Left = 112
    Top = 221
    ParentFont = False
    Style.StyleController = fmMain.cxEditStyleController1
    TabOrder = 11
    Width = 107
  end
  object txtChLimitCount: TcxSpinEdit
    Left = 317
    Top = 221
    ParentFont = False
    Properties.MaxValue = 999.000000000000000000
    Style.StyleController = fmMain.cxEditStyleController1
    TabOrder = 12
    Width = 49
  end
  object txtChLimitDay: TcxSpinEdit
    Left = 468
    Top = 221
    ParentFont = False
    Properties.MaxValue = 999.000000000000000000
    Style.StyleController = fmMain.cxEditStyleController1
    TabOrder = 13
    Width = 49
  end
  object txtAdmin: TcxCheckBox
    Left = 401
    Top = 89
    Caption = #31995#32113#31649#29702#21729
    ParentFont = False
    Style.StyleController = fmMain.cxEditStyleController1
    TabOrder = 14
    Width = 91
  end
  object UserGrid: TcxGrid
    Left = 7
    Top = 257
    Width = 518
    Height = 249
    TabOrder = 15
    LookAndFeel.Kind = lfFlat
    object UserGridView: TcxGridDBTableView
      NavigatorButtons.ConfirmDelete = False
      DataController.DataSource = UserDataSource
      DataController.KeyFieldNames = 'USERID'
      DataController.Summary.DefaultGroupSummaryItems = <>
      DataController.Summary.FooterSummaryItems = <>
      DataController.Summary.SummaryGroups = <>
      OptionsCustomize.ColumnFiltering = False
      OptionsData.Deleting = False
      OptionsData.Editing = False
      OptionsData.Inserting = False
      OptionsSelection.CellSelect = False
      OptionsView.GroupByBox = False
      OptionsView.Indicator = True
      object UserGridViewADMIN: TcxGridDBColumn
        Caption = #31995#32113#31649#29702#21729
        DataBinding.FieldName = 'ADMIN'
        PropertiesClassName = 'TcxCheckBoxProperties'
        Properties.DisplayChecked = 'Y'
        Properties.DisplayUnchecked = 'N'
        Properties.NullStyle = nssUnchecked
        Properties.ValueChecked = 'Y'
        Properties.ValueUnchecked = 'N'
        Width = 71
      end
      object UserGridViewUSERID: TcxGridDBColumn
        Caption = #20351#29992#32773#24115#34399
        DataBinding.FieldName = 'USERID'
        Width = 88
      end
      object UserGridViewUSERNAME: TcxGridDBColumn
        Caption = #20351#29992#32773#22995#21517
        DataBinding.FieldName = 'USERNAME'
        Width = 117
      end
      object UserGridViewPASSWORD: TcxGridDBColumn
        Caption = #23494#30908
        DataBinding.FieldName = 'PASSWORD'
        PropertiesClassName = 'TcxTextEditProperties'
        Properties.EchoMode = eemPassword
        Properties.PasswordChar = '*'
        Width = 128
      end
      object UserGridViewUSERLOCK: TcxGridDBColumn
        Caption = #24115#34399#37782#23450
        DataBinding.FieldName = 'USERLOCK'
        PropertiesClassName = 'TcxCheckBoxProperties'
        Properties.DisplayChecked = 'Y'
        Properties.DisplayUnchecked = 'N'
        Properties.NullStyle = nssUnchecked
        Properties.ValueChecked = 'Y'
        Properties.ValueUnchecked = 'N'
        Width = 71
      end
      object UserGridViewCOMPSTR: TcxGridDBColumn
        Caption = #21487#30332#36865#25351#20196#31995#32113#21488
        DataBinding.FieldName = 'COMPSTR'
        Width = 285
      end
      object UserGridViewEXPIREDATE: TcxGridDBColumn
        Caption = #20351#29992#21040#26399#26085
        DataBinding.FieldName = 'EXPIREDATE'
        Width = 111
      end
      object UserGridViewCHLIMITCOUNT: TcxGridDBColumn
        Caption = #38283#38971#36947#38480#21046#27425#25976
        DataBinding.FieldName = 'CHLIMITCOUNT'
        Width = 103
      end
      object UserGridViewCHLIMITDAY: TcxGridDBColumn
        Caption = #38283#38971#36947#38480#21046#22825#25976
        DataBinding.FieldName = 'CHLIMITDAY'
      end
    end
    object UserGridLevel: TcxGridLevel
      GridView = UserGridView
    end
  end
  object UserDataSet: TADOQuery
    Connection = fmMain.ComConnection
    CursorType = ctStatic
    AfterScroll = UserDataSetAfterScroll
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM CASIMPLE ORDER BY USERID')
    Left = 92
    Top = 328
  end
  object UserDataSource: TDataSource
    DataSet = UserDataSet
    Left = 124
    Top = 328
  end
  object DataWriter: TADOQuery
    Connection = fmMain.ComConnection
    Parameters = <>
    Left = 92
    Top = 364
  end
end
