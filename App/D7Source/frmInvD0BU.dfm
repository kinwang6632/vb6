object frmInvD0B: TfrmInvD0B
  Left = 206
  Top = 202
  BorderStyle = bsDialog
  Caption = '[D0B] '#25490#31243#35373#23450
  ClientHeight = 327
  ClientWidth = 468
  Color = clBtnFace
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
  object pnlEdit: TPanel
    Left = 0
    Top = 35
    Width = 468
    Height = 188
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 14
      Top = 10
      Width = 64
      Height = 19
      Caption = #20844#21496#20195#30908
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object chkEmail: TcxCheckBox
      Left = 9
      Top = 37
      Caption = '1.'#38651#23376#37109#20214
      ParentFont = False
      Properties.Alignment = taLeftJustify
      Properties.DisplayChecked = '1'
      Properties.DisplayUnchecked = '0'
      Properties.NullStyle = nssUnchecked
      Properties.ValueChecked = 1
      Properties.ValueUnchecked = 0
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.StyleController = dtmMain.cxEditStyle
      Style.TextColor = clBlack
      Style.IsFontAssigned = True
      TabOrder = 0
      OnClick = chkEmailClick
      Width = 92
    end
    object edtMailSet: TcxCurrencyEdit
      Left = 107
      Top = 38
      EditValue = 0
      Properties.Alignment.Horz = taRightJustify
      Properties.AssignedValues.DisplayFormat = True
      Properties.MaxValue = 999999.000000000000000000
      Properties.Nullable = False
      Properties.NullString = '0'
      TabOrder = 1
      Width = 47
    end
    object lblTVMail: TcxLabel
      Left = 159
      Top = 66
      Caption = #20998#37912
      ParentFont = False
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.IsFontAssigned = True
    end
    object chkMobile: TcxCheckBox
      Left = 216
      Top = 37
      Caption = '2.'#25163#27231#34399#30908
      ParentFont = False
      Properties.Alignment = taLeftJustify
      Properties.DisplayChecked = '1'
      Properties.DisplayUnchecked = '0'
      Properties.NullStyle = nssUnchecked
      Properties.ValueChecked = 1
      Properties.ValueUnchecked = 0
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.StyleController = dtmMain.cxEditStyle
      Style.TextColor = clBlack
      Style.IsFontAssigned = True
      TabOrder = 3
      OnClick = chkMobileClick
      Width = 97
    end
    object lblMobile: TcxLabel
      Left = 364
      Top = 39
      Caption = #20998#37912
      ParentFont = False
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.IsFontAssigned = True
    end
    object edtMobileSet: TcxCurrencyEdit
      Left = 314
      Top = 38
      EditValue = 0
      Properties.Alignment.Horz = taRightJustify
      Properties.AssignedValues.DisplayFormat = True
      Properties.MaxValue = 999999.000000000000000000
      Properties.Nullable = False
      Properties.NullString = '0'
      TabOrder = 4
      Width = 47
    end
    object chkTVMail: TcxCheckBox
      Left = 9
      Top = 65
      Caption = '3.TVMail'
      ParentFont = False
      Properties.Alignment = taLeftJustify
      Properties.DisplayChecked = '1'
      Properties.DisplayUnchecked = '0'
      Properties.NullStyle = nssUnchecked
      Properties.ValueChecked = 1
      Properties.ValueUnchecked = 0
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.StyleController = dtmMain.cxEditStyle
      Style.TextColor = clBlack
      Style.IsFontAssigned = True
      TabOrder = 5
      OnClick = chkTVMailClick
      Width = 86
    end
    object chkCM: TcxCheckBox
      Left = 216
      Top = 64
      Caption = '4.CM'#23566#27969
      ParentFont = False
      Properties.Alignment = taLeftJustify
      Properties.DisplayChecked = '1'
      Properties.DisplayUnchecked = '0'
      Properties.NullStyle = nssUnchecked
      Properties.ValueChecked = 1
      Properties.ValueUnchecked = 0
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.StyleController = dtmMain.cxEditStyle
      Style.TextColor = clBlack
      Style.IsFontAssigned = True
      TabOrder = 7
      OnClick = chkCMClick
      Width = 86
    end
    object edtTVMailSet: TcxCurrencyEdit
      Left = 107
      Top = 66
      EditValue = 0
      Properties.Alignment.Horz = taRightJustify
      Properties.AssignedValues.DisplayFormat = True
      Properties.MaxValue = 999999.000000000000000000
      Properties.Nullable = False
      Properties.NullString = '0'
      TabOrder = 6
      Width = 47
    end
    object lblCM: TcxLabel
      Left = 364
      Top = 66
      Caption = #20998#37912
      ParentFont = False
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.IsFontAssigned = True
    end
    object cxLabel5: TcxLabel
      Left = 18
      Top = 95
      Caption = #20778#20808#38918#24207
      ParentFont = False
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.IsFontAssigned = True
    end
    object edtOrder: TcxTextEdit
      Left = 88
      Top = 96
      ParentFont = False
      Properties.MaxLength = 4
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -11
      Style.Font.Name = 'MS Sans Serif'
      Style.Font.Style = []
      Style.IsFontAssigned = True
      TabOrder = 11
      OnKeyPress = edtOrderKeyPress
      Width = 49
    end
    object rdgInvUse: TcxRadioGroup
      Left = 18
      Top = 125
      Caption = #25424#36104#35387#35352
      Properties.Columns = 3
      Properties.Items = <
        item
          Caption = #25424#36104
        end
        item
          Caption = #38750#25424#36104
        end
        item
          Caption = #20840#37096
        end>
      ItemIndex = 2
      TabOrder = 12
      Height = 50
      Width = 435
    end
    object cxLabel6: TcxLabel
      Left = 146
      Top = 96
      Caption = #21516#26178#21152#36865
      ParentFont = False
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.IsFontAssigned = True
    end
    object edtAddGive: TcxTextEdit
      Left = 212
      Top = 95
      ParentFont = False
      Properties.MaxLength = 4
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -11
      Style.Font.Name = 'MS Sans Serif'
      Style.Font.Style = []
      Style.IsFontAssigned = True
      TabOrder = 14
      OnKeyPress = edtAddGiveKeyPress
      Width = 49
    end
    object edtCMSet: TcxCurrencyEdit
      Left = 314
      Top = 66
      EditValue = 0
      Properties.Alignment.Horz = taRightJustify
      Properties.AssignedValues.DisplayFormat = True
      Properties.MaxValue = 999999.000000000000000000
      Properties.Nullable = False
      Properties.NullString = '0'
      TabOrder = 15
      Width = 47
    end
    object cboComId: TcxComboBox
      Left = 84
      Top = 10
      ParentFont = False
      Properties.DropDownListStyle = lsFixedList
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.IsFontAssigned = True
      TabOrder = 16
      Width = 61
    end
    object lblCMSend: TcxLabel
      Left = 266
      Top = 96
      Caption = #23566#27969#36890#30693#32080#26463#25490#31243
      ParentFont = False
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.IsFontAssigned = True
    end
    object edtEndCMNotify: TcxCurrencyEdit
      Left = 392
      Top = 96
      EditValue = 0
      Properties.Alignment.Horz = taRightJustify
      Properties.AssignedValues.DisplayFormat = True
      Properties.DecimalPlaces = 0
      Properties.MaxValue = 999999.000000000000000000
      Properties.NullString = '0'
      TabOrder = 18
      Width = 47
    end
    object lblCMSend2: TcxLabel
      Left = 442
      Top = 96
      Caption = #22825
      ParentFont = False
      Style.Font.Charset = ANSI_CHARSET
      Style.Font.Color = clWindowText
      Style.Font.Height = -13
      Style.Font.Name = 'Tahoma'
      Style.Font.Style = []
      Style.IsFontAssigned = True
    end
  end
  object lblEmail: TcxLabel
    Left = 159
    Top = 74
    Caption = #20998#37912
    ParentFont = False
    Style.Font.Charset = ANSI_CHARSET
    Style.Font.Color = clWindowText
    Style.Font.Height = -13
    Style.Font.Name = 'Tahoma'
    Style.Font.Style = []
    Style.IsFontAssigned = True
  end
  object pnlGrid: TPanel
    Left = 0
    Top = 223
    Width = 468
    Height = 247
    Align = alTop
    TabOrder = 2
    object cxGrid1: TcxGrid
      Left = 1
      Top = 1
      Width = 466
      Height = 87
      Align = alTop
      TabOrder = 0
      object cxGrid1DBTableView1: TcxGridDBTableView
        NavigatorButtons.ConfirmDelete = False
        DataController.DataSource = DataSource1
        DataController.Summary.DefaultGroupSummaryItems = <>
        DataController.Summary.FooterSummaryItems = <>
        DataController.Summary.SummaryGroups = <>
        OptionsCustomize.ColumnFiltering = False
        OptionsCustomize.ColumnGrouping = False
        OptionsCustomize.ColumnMoving = False
        OptionsCustomize.ColumnSorting = False
        OptionsData.Deleting = False
        OptionsData.Editing = False
        OptionsData.Inserting = False
        OptionsView.GroupByBox = False
        object gvCompId: TcxGridDBColumn
          Caption = #20844#21496#21029
          DataBinding.FieldName = 'CompId'
          Width = 66
        end
        object gvEmailNotify: TcxGridDBColumn
          Caption = 'Email'#36890#30693#25490#31243
          DataBinding.FieldName = 'EmailNotify'
          Width = 101
        end
        object gvSmsNotify: TcxGridDBColumn
          Caption = #25163#27231#36890#30693#25490#31243
          DataBinding.FieldName = 'SmsNotify'
          Width = 89
        end
        object gvTVMAILNotify: TcxGridDBColumn
          Caption = 'TVMAIL'#36890#30693#25490#31243
          DataBinding.FieldName = 'TVMAILNotify'
          Width = 101
        end
        object gvCMNotify: TcxGridDBColumn
          Caption = 'CM'#23566#27969#36890#30693#25490#31243
          DataBinding.FieldName = 'CMNotify'
          Width = 108
        end
        object gvPriorityOrder: TcxGridDBColumn
          Caption = #20778#20808#38918#24207
          DataBinding.FieldName = 'PriorityOrder'
          Width = 78
        end
        object gvAddGive: TcxGridDBColumn
          Caption = #21516#26178#21152#36865
          DataBinding.FieldName = 'AddGive'
        end
        object gvINVUSEID: TcxGridDBColumn
          Caption = #25424#36104#35387#35352
          DataBinding.FieldName = 'INVUSEID'
          Width = 74
        end
        object gvUpdEn: TcxGridDBColumn
          Caption = #30064#21205#20154#21729
          DataBinding.FieldName = 'UpdEn'
          Width = 96
        end
        object gvUpdTime: TcxGridDBColumn
          Caption = #30064#21205#26178#38291
          DataBinding.FieldName = 'UpdTime'
          Width = 160
        end
      end
      object cxGrid1Level1: TcxGridLevel
        GridView = cxGrid1DBTableView1
      end
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 0
    Width = 468
    Height = 35
    Align = alTop
    TabOrder = 3
    object btnAppend: TBitBtn
      Left = 9
      Top = 4
      Width = 71
      Height = 26
      Caption = #26032#22686
      TabOrder = 0
      OnClick = btnAppendClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        33333333FF33333333FF333993333333300033377F3333333777333993333333
        300033F77FFF3333377739999993333333333777777F3333333F399999933333
        33003777777333333377333993333333330033377F3333333377333993333333
        3333333773333333333F333333333333330033333333F33333773333333C3333
        330033333337FF3333773333333CC333333333FFFFF77FFF3FF33CCCCCCCCCC3
        993337777777777F77F33CCCCCCCCCC3993337777777777377333333333CC333
        333333333337733333FF3333333C333330003333333733333777333333333333
        3000333333333333377733333333333333333333333333333333}
      NumGlyphs = 2
    end
    object btnEdit: TBitBtn
      Left = 82
      Top = 4
      Width = 71
      Height = 26
      Caption = #20462#25913
      TabOrder = 1
      OnClick = btnEditClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000120B0000120B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333000000
        000033333377777777773333330FFFFFFFF03FF3FF7FF33F3FF700300000FF0F
        00F077F777773F737737E00BFBFB0FFFFFF07773333F7F3333F7E0BFBF000FFF
        F0F077F3337773F3F737E0FBFBFBF0F00FF077F3333FF7F77F37E0BFBF00000B
        0FF077F3337777737337E0FBFBFBFBF0FFF077F33FFFFFF73337E0BF0000000F
        FFF077FF777777733FF7000BFB00B0FF00F07773FF77373377373330000B0FFF
        FFF03337777373333FF7333330B0FFFF00003333373733FF777733330B0FF00F
        0FF03333737F37737F373330B00FFFFF0F033337F77F33337F733309030FFFFF
        00333377737FFFFF773333303300000003333337337777777333}
      NumGlyphs = 2
    end
    object btnDelete: TBitBtn
      Left = 155
      Top = 4
      Width = 71
      Height = 26
      Caption = #21034#38500
      TabOrder = 2
      OnClick = btnDeleteClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        333333333333333333FF33333333333330003333333333333777333333333333
        300033FFFFFF3333377739999993333333333777777F3333333F399999933333
        3300377777733333337733333333333333003333333333333377333333333333
        3333333333333333333F333333333333330033333F33333333773333C3333333
        330033337F3333333377333CC3333333333333F77FFFFFFF3FF33CCCCCCCCCC3
        993337777777777F77F33CCCCCCCCCC399333777777777737733333CC3333333
        333333377F33333333FF3333C333333330003333733333333777333333333333
        3000333333333333377733333333333333333333333333333333}
      NumGlyphs = 2
    end
    object btnCancel: TBitBtn
      Left = 228
      Top = 4
      Width = 71
      Height = 26
      Caption = #21462#28040
      TabOrder = 3
      OnClick = btnCancelClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000120B0000120B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00500005000555
        555557777F777555F55500000000555055557777777755F75555005500055055
        555577F5777F57555555005550055555555577FF577F5FF55555500550050055
        5555577FF77577FF555555005050110555555577F757777FF555555505099910
        555555FF75777777FF555005550999910555577F5F77777775F5500505509990
        3055577F75F77777575F55005055090B030555775755777575755555555550B0
        B03055555F555757575755550555550B0B335555755555757555555555555550
        BBB35555F55555575F555550555555550BBB55575555555575F5555555555555
        50BB555555555555575F555555555555550B5555555555555575}
      NumGlyphs = 2
    end
    object btnSave: TBitBtn
      Left = 301
      Top = 4
      Width = 71
      Height = 26
      Caption = #20786#23384
      TabOrder = 4
      OnClick = btnSaveClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000120B0000120B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00555555555555
        555555555555555555555555555555555555555555FF55555555555559055555
        55555555577FF5555555555599905555555555557777F5555555555599905555
        555555557777FF5555555559999905555555555777777F555555559999990555
        5555557777777FF5555557990599905555555777757777F55555790555599055
        55557775555777FF5555555555599905555555555557777F5555555555559905
        555555555555777FF5555555555559905555555555555777FF55555555555579
        05555555555555777FF5555555555557905555555555555777FF555555555555
        5990555555555555577755555555555555555555555555555555}
      NumGlyphs = 2
    end
    object btnExit: TBitBtn
      Left = 379
      Top = 3
      Width = 71
      Height = 26
      Caption = #32080#26463
      TabOrder = 5
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
  end
  object adoINV042: TADOQuery
    Connection = dtmMain.InvConnection
    AfterScroll = adoINV042AfterScroll
    Parameters = <>
    Left = 424
    Top = 58
  end
  object DataSource1: TDataSource
    DataSet = adoINV042
    OnDataChange = DataSource1DataChange
    Left = 400
    Top = 98
  end
end
