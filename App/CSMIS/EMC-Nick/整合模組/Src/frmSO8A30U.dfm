object frmSO8A30: TfrmSO8A30
  Left = -4
  Top = 34
  Width = 798
  Height = 548
  Caption = #25286#24115#20316#26989#20043#21508#29986#21697#36039#26009#27284
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poScreenCenter
  WindowState = wsMaximized
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object Groupbox1: TGroupBox
    Left = 0
    Top = 0
    Width = 790
    Height = 45
    Align = alTop
    TabOrder = 0
    object btnAppend: TBitBtn
      Left = 414
      Top = 10
      Width = 70
      Height = 30
      Caption = #26032#22686
      ParentShowHint = False
      ShowHint = True
      TabOrder = 5
      OnClick = btnAppendClick
      NumGlyphs = 2
    end
    object btnEdit: TBitBtn
      Left = 484
      Top = 10
      Width = 70
      Height = 30
      Caption = #20462#25913
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
      OnClick = btnEditClick
      NumGlyphs = 2
    end
    object btnDelete: TBitBtn
      Left = 554
      Top = 10
      Width = 70
      Height = 30
      Caption = #21034#38500
      ParentShowHint = False
      ShowHint = True
      TabOrder = 7
      OnClick = btnDeleteClick
      NumGlyphs = 2
    end
    object btnCloseCancel: TBitBtn
      Left = 624
      Top = 10
      Width = 70
      Height = 30
      Caption = '(&X)'#32080#26463
      ParentShowHint = False
      ShowHint = True
      TabOrder = 8
      OnClick = btnCloseCancelClick
      NumGlyphs = 2
    end
    object btnSave: TBitBtn
      Left = 344
      Top = 10
      Width = 70
      Height = 30
      Caption = #20786#23384
      Enabled = False
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
      OnClick = btnSaveClick
      NumGlyphs = 2
    end
    object btnNext: TBitBtn
      Left = 203
      Top = 10
      Width = 70
      Height = 30
      Caption = #19979#19968#31558
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnClick = btnNextClick
    end
    object btnPrior: TBitBtn
      Left = 133
      Top = 10
      Width = 70
      Height = 30
      Caption = #19978#19968#31558
      TabOrder = 1
      OnClick = btnPriorClick
    end
    object btnLast: TBitBtn
      Left = 273
      Top = 10
      Width = 70
      Height = 30
      Caption = #26368#24460#19968#31558
      TabOrder = 3
      OnClick = btnLastClick
    end
    object BtnFirst: TBitBtn
      Left = 63
      Top = 10
      Width = 70
      Height = 30
      Caption = #31532#19968#31558
      TabOrder = 0
      OnClick = BtnFirstClick
    end
  end
  object gbxSo112: TGroupBox
    Left = 0
    Top = 45
    Width = 790
    Height = 65
    Align = alTop
    Enabled = False
    TabOrder = 1
    object Label2: TLabel
      Left = 30
      Top = 14
      Width = 98
      Height = 13
      AutoSize = False
      Caption = #29986#21697#20195#30908#33287#21517#31281
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label3: TLabel
      Left = 280
      Top = 14
      Width = 113
      Height = 13
      AutoSize = False
      Caption = #38971#36947#21830#20195#30908#21450#21517#31281
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label5: TLabel
      Left = 694
      Top = 17
      Width = 26
      Height = 13
      Alignment = taRightJustify
      AutoSize = False
      Caption = #39006#22411
      FocusControl = dedType
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label7: TLabel
      Left = 534
      Top = 15
      Width = 39
      Height = 13
      Alignment = taRightJustify
      AutoSize = False
      Caption = #21407#23450#20729
      FocusControl = dedPrice
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label4: TLabel
      Left = 74
      Top = 40
      Width = 41
      Height = 13
      AutoSize = False
      Caption = #21443#32771#34399
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label9: TLabel
      Left = 282
      Top = 42
      Width = 95
      Height = 13
      AutoSize = False
      Caption = #21855#29992#22522#26412#20844#24335
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label11: TLabel
      Left = 460
      Top = 41
      Width = 113
      Height = 13
      Alignment = taRightJustify
      AutoSize = False
      Caption = #38971#36947#21830#25286#24115#30334#20998#27604
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label16: TLabel
      Left = 639
      Top = 44
      Width = 82
      Height = 13
      Alignment = taRightJustify
      AutoSize = False
      Caption = 'So'#25286#24115#30334#20998#27604
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object dlcProductId: TDBLookupComboBox
      Left = 129
      Top = 9
      Width = 149
      Height = 24
      DataField = 'PRODUCT_ID'
      DataSource = dsrSo112M
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      KeyField = 'CODENO'
      ListField = 'CODENO_DESCRIPTION'
      ListSource = dsrCd024Look
      TabOrder = 0
      OnCloseUp = dlcProductIdCloseUp
    end
    object dedType: TDBEdit
      Tag = 3
      Left = 724
      Top = 11
      Width = 44
      Height = 24
      DataField = 'TYPE'
      DataSource = dsrSo112M
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      TabOrder = 3
    end
    object dedPrice: TDBEdit
      Tag = 5
      Left = 580
      Top = 9
      Width = 55
      Height = 24
      DataField = 'PRICE'
      DataSource = dsrSo112M
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      TabOrder = 2
      OnExit = dedPriceExit
    end
    object dedRef_no: TDBEdit
      Left = 130
      Top = 36
      Width = 54
      Height = 24
      DataField = 'REF_NO'
      DataSource = dsrSo112M
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      TabOrder = 4
      OnExit = dedRef_noExit
    end
    object dlcPRoviderId: TDBLookupComboBox
      Left = 384
      Top = 9
      Width = 124
      Height = 24
      DataField = 'PROVIDER_ID'
      DataSource = dsrSo112M
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      KeyField = 'PROVIDER_ID'
      ListField = 'PROVIDER_ID_NAME'
      ListSource = dsrSo113Look
      TabOrder = 1
      OnCloseUp = dlcPRoviderIdCloseUp
    end
    object DBCheckBox1: TDBCheckBox
      Left = 264
      Top = 40
      Width = 15
      Height = 17
      Caption = 'DBCheckBox1'
      DataField = 'USEBASEFORMULA'
      DataSource = dsrSo112M
      TabOrder = 5
      ValueChecked = 'Y'
      ValueUnchecked = 'N'
    end
    object DBEdit1: TDBEdit
      Left = 580
      Top = 37
      Width = 57
      Height = 24
      DataField = 'PROVIDER_PERCENT'
      DataSource = dsrSo112M
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      TabOrder = 6
    end
    object DBEdit2: TDBEdit
      Left = 725
      Top = 38
      Width = 45
      Height = 24
      DataField = 'SO_PERCENT'
      DataSource = dsrSo112M
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      TabOrder = 7
    end
  end
  object pnlSo112: TPanel
    Left = 0
    Top = 110
    Width = 790
    Height = 210
    Align = alClient
    BevelOuter = bvLowered
    Caption = 'pnlSo112'
    TabOrder = 2
    object dgrSo112M: TDBGrid
      Left = 1
      Top = 1
      Width = 788
      Height = 208
      Align = alClient
      DataSource = dsrSo112M
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgConfirmDelete, dgCancelOnExit]
      ParentFont = False
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
      OnCellClick = dgrSo112MCellClick
      OnDblClick = dgrSo112MDblClick
      OnTitleClick = dgrSo112MTitleClick
      Columns = <
        item
          Expanded = False
          FieldName = 'PRODUCT_ID'
          Visible = False
        end
        item
          Expanded = False
          FieldName = 'PRODUCT_NAME'
          Width = 94
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'PROVIDER_ID'
          Visible = False
        end
        item
          Expanded = False
          FieldName = 'PROVIDER_NAME'
          Width = 109
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'PRICE'
          Width = 88
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'TYPE'
          Width = 55
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'REF_NO'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'OPERATOR'
          Width = 160
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'UPDTIME'
          Visible = True
        end>
    end
  end
  object pnlSo111: TPanel
    Left = 0
    Top = 320
    Width = 790
    Height = 201
    Align = alBottom
    BevelOuter = bvLowered
    TabOrder = 3
    object DBTEXT1: TDBText
      Left = 660
      Top = 87
      Width = 54
      Height = 21
      Alignment = taCenter
      Color = clBtnFace
      DataField = 'SUBSCRIBER_COUNT_PERCENT1'
      DataSource = dsrSo111D
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object DBTEXT2: TDBText
      Left = 660
      Top = 109
      Width = 54
      Height = 21
      Alignment = taCenter
      Color = clBtnFace
      DataField = 'SUBSCRIBER_COUNT_PERCENT2'
      DataSource = dsrSo111D
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object DBTEXT3: TDBText
      Left = 660
      Top = 131
      Width = 54
      Height = 21
      Alignment = taCenter
      Color = clBtnFace
      DataField = 'SUBSCRIBER_COUNT_PERCENT3'
      DataSource = dsrSo111D
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object DBTEXT4: TDBText
      Left = 660
      Top = 152
      Width = 54
      Height = 21
      Alignment = taCenter
      Color = clBtnFace
      DataField = 'SUBSCRIBER_COUNT_PERCENT4'
      DataSource = dsrSo111D
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object DBTEXT5: TDBText
      Left = 660
      Top = 175
      Width = 54
      Height = 21
      Alignment = taCenter
      Color = clBtnFace
      DataField = 'SUBSCRIBER_COUNT_PERCENT5'
      DataSource = dsrSo111D
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object DBTEXT6: TDBText
      Left = 730
      Top = 86
      Width = 56
      Height = 21
      Alignment = taCenter
      Color = clBtnFace
      DataField = 'AMOUNT_PERCENT1'
      DataSource = dsrSo111D
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object DBTEXT7: TDBText
      Left = 730
      Top = 108
      Width = 56
      Height = 21
      Alignment = taCenter
      Color = clBtnFace
      DataField = 'AMOUNT_PERCENT2'
      DataSource = dsrSo111D
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object DBTEXT8: TDBText
      Left = 730
      Top = 129
      Width = 56
      Height = 21
      Alignment = taCenter
      Color = clBtnFace
      DataField = 'AMOUNT_PERCENT3'
      DataSource = dsrSo111D
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object DBTEXT9: TDBText
      Left = 730
      Top = 151
      Width = 56
      Height = 21
      Alignment = taCenter
      Color = clBtnFace
      DataField = 'AMOUNT_PERCENT4'
      DataSource = dsrSo111D
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object DBTEXT10: TDBText
      Left = 730
      Top = 173
      Width = 56
      Height = 21
      Alignment = taCenter
      Color = clBtnFace
      DataField = 'AMOUNT_PERCENT5'
      DataSource = dsrSo111D
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object Label8: TLabel
      Left = 644
      Top = 87
      Width = 15
      Height = 16
      Caption = '(1)'
    end
    object Label10: TLabel
      Left = 643
      Top = 108
      Width = 15
      Height = 16
      Caption = '(2)'
    end
    object Label13: TLabel
      Left = 643
      Top = 130
      Width = 15
      Height = 16
      Caption = '(3)'
    end
    object Label14: TLabel
      Left = 643
      Top = 175
      Width = 15
      Height = 16
      Caption = '(5)'
    end
    object Label15: TLabel
      Left = 643
      Top = 152
      Width = 15
      Height = 16
      Caption = '(4)'
    end
    object Label1: TLabel
      Left = 654
      Top = 61
      Width = 58
      Height = 17
      AutoSize = False
      Caption = #25142#25976#32026#36317
      FocusControl = DBTEXT1
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Label12: TLabel
      Left = 720
      Top = 60
      Width = 74
      Height = 17
      AutoSize = False
      Caption = #37329#38989'-'#30334#20998#27604
      FocusControl = DBTEXT6
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object dgrSo111D: TDBGrid
      Left = 1
      Top = 42
      Width = 640
      Height = 158
      Align = alLeft
      DataSource = dsrSo111D
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      ReadOnly = True
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -13
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
      OnDblClick = dgrSo111DDblClick
      OnEditButtonClick = dgrSo111DEditButtonClick
      Columns = <
        item
          Expanded = False
          FieldName = 'SO_PROVIDER_ID'
          Visible = False
        end
        item
          ButtonStyle = cbsNone
          Expanded = False
          FieldName = 'SO_PROVIDER_DESC'
          Width = 117
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'PRODUCT_ID'
          Visible = False
        end
        item
          Expanded = False
          FieldName = 'PRODUCT_NAME'
          Visible = False
        end
        item
          ButtonStyle = cbsNone
          Expanded = False
          FieldName = 'FORMULA_NAME_LU'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'FORMULA_ID'
          Visible = False
        end
        item
          ButtonStyle = cbsNone
          Expanded = False
          FieldName = 'COMPUTE_ISSUE_NAME'
          Width = 70
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'COMPUTE_ISSUE'
          Visible = False
        end
        item
          ButtonStyle = cbsNone
          Expanded = False
          FieldName = 'RANGE_UNIT_NAME'
          Width = 85
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'RANGE_UNIT'
          Visible = False
        end
        item
          Expanded = False
          FieldName = 'OPERATOR'
          Width = 60
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'UPDTIME'
          Visible = True
        end>
    end
    object Panel1: TPanel
      Left = 1
      Top = 1
      Width = 788
      Height = 41
      Align = alTop
      Caption = 'Panel1'
      TabOrder = 1
      object BitBtn2: TBitBtn
        Left = 58
        Top = 5
        Width = 70
        Height = 30
        Caption = #31532#19968#31558
        TabOrder = 0
        OnClick = BitBtn2Click
      end
      object BitBtn9: TBitBtn
        Left = 128
        Top = 5
        Width = 70
        Height = 30
        Caption = #19978#19968#31558
        TabOrder = 1
        OnClick = BitBtn9Click
      end
      object BitBtn10: TBitBtn
        Left = 198
        Top = 5
        Width = 70
        Height = 30
        Caption = #19979#19968#31558
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = BitBtn10Click
      end
      object BitBtn11: TBitBtn
        Left = 268
        Top = 5
        Width = 70
        Height = 30
        Caption = #26368#24460#19968#31558
        TabOrder = 3
        OnClick = BitBtn11Click
      end
      object btnDAppend: TBitBtn
        Left = 338
        Top = 5
        Width = 75
        Height = 30
        Caption = #26032#22686
        TabOrder = 4
        OnClick = btnDAppendClick
      end
      object btnDEdit: TBitBtn
        Left = 413
        Top = 5
        Width = 75
        Height = 30
        Caption = #20462#25913
        TabOrder = 5
        OnClick = btnDEditClick
      end
      object btnDDelete: TBitBtn
        Left = 488
        Top = 5
        Width = 75
        Height = 30
        Caption = #21034#38500
        TabOrder = 6
        OnClick = btnDDeleteClick
      end
    end
  end
  object dsrSo113Cd039: TDataSource
    DataSet = dtmMain2.cdsSo113Cd039
    Left = 259
    Top = 219
  end
  object dsrSo112M: TDataSource
    DataSet = dtmMain2.cdsSo112M
    Left = 496
    Top = 56
  end
  object dsrSo111D: TDataSource
    DataSet = dtmMain2.cdsSo111D
    Left = 32
    Top = 400
  end
  object dsrSo113Look: TDataSource
    DataSet = dtmMain2.cdsSo113Look
    Left = 456
    Top = 56
  end
  object dsrSo110: TDataSource
    DataSet = dtmMain2.cdsSo110
    Left = 312
    Top = 160
  end
  object dsrCd024CT: TDataSource
    DataSet = dtmMain2.cdsCd024CT
    Left = 216
    Top = 152
  end
  object DataSource1: TDataSource
    DataSet = dtmMain2.cdsSo110
    Left = 400
    Top = 136
  end
  object dsrCd024Look: TDataSource
    DataSet = dtmMain2.cdsCd024Look
    Left = 144
    Top = 112
  end
  object dsrSo110Look: TDataSource
    DataSet = dtmMain2.cdsSo110Look
    Left = 504
    Top = 104
  end
end
