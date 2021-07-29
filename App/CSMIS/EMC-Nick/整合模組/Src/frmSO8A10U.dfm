object frmSO8A10: TfrmSO8A10
  Left = 36
  Top = 59
  Width = 676
  Height = 489
  Caption = #25286#24115#20844#24335#20195#30908
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 0
    Top = 0
    Width = 668
    Height = 146
    Align = alTop
    Enabled = False
    TabOrder = 0
    object Label1: TLabel
      Left = 24
      Top = 26
      Width = 33
      Height = 23
      Alignment = taRightJustify
      AutoSize = False
      Caption = #20195#30908
      FocusControl = dedFormulaId
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 18
      Top = 56
      Width = 38
      Height = 13
      Alignment = taRightJustify
      AutoSize = False
      Caption = #21517#31281
      FocusControl = DBEdit2
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label4: TLabel
      Left = 85
      Top = 111
      Width = 58
      Height = 13
      AutoSize = False
      Caption = #26159#21542#20572#29992
      Color = clBtnFace
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object Label3: TLabel
      Left = 25
      Top = 87
      Width = 30
      Height = 13
      Alignment = taRightJustify
      AutoSize = False
      Caption = #35498#26126
      FocusControl = DBEdit1
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label5: TLabel
      Left = 184
      Top = 25
      Width = 89
      Height = 16
      Alignment = taRightJustify
      AutoSize = False
      Caption = #20844#24335#21443#32771#20195#30908
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object dedFormulaId: TDBEdit
      Left = 67
      Top = 22
      Width = 56
      Height = 24
      DataField = 'FORMULA_ID'
      DataSource = dsrSo110
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      ParentFont = False
      TabOrder = 0
    end
    object DBEdit2: TDBEdit
      Left = 66
      Top = 51
      Width = 587
      Height = 24
      DataField = 'FORMULA_NAME'
      DataSource = dsrSo110
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      ParentFont = False
      TabOrder = 1
    end
    object DBCheckBox1: TDBCheckBox
      Left = 66
      Top = 109
      Width = 15
      Height = 17
      DataField = 'STOPFLAG'
      DataSource = dsrSo110
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      ValueChecked = '1'
      ValueUnchecked = '0'
    end
    object DBEdit1: TDBEdit
      Left = 66
      Top = 82
      Width = 587
      Height = 24
      DataField = 'FORMULA_DESC'
      DataSource = dsrSo110
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      ParentFont = False
      TabOrder = 2
    end
    object CobRefNo: TComboBox
      Left = 282
      Top = 21
      Width = 183
      Height = 24
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      ItemHeight = 16
      ParentFont = False
      TabOrder = 4
      OnChange = CobRefNoChange
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 146
    Width = 668
    Height = 271
    Align = alTop
    BevelOuter = bvLowered
    TabOrder = 1
    object dgr1: TDBGrid
      Left = 1
      Top = 1
      Width = 666
      Height = 249
      Align = alClient
      DataSource = dsrSo110
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      ReadOnly = True
      TabOrder = 1
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
      OnCellClick = dgr1CellClick
      OnDblClick = dgr1DblClick
      OnEnter = dgr1Enter
      Columns = <
        item
          Expanded = False
          FieldName = 'FORMULA_ID'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'FORMULA_NAME'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'FORMULA_DESC'
          Width = 169
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'REFNO_NAME'
          Width = 229
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'STOPFLAG'
          Visible = False
        end
        item
          Expanded = False
          FieldName = 'STOPFLAG_NAME'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'OPERATOR'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'UPDTIME'
          Visible = True
        end>
    end
    object StatusBar1: TStatusBar
      Left = 1
      Top = 250
      Width = 666
      Height = 20
      Color = clCream
      Panels = <
        item
          Width = 692
        end>
      SimplePanel = False
    end
  end
  object GroupBox2: TGroupBox
    Left = 0
    Top = 415
    Width = 668
    Height = 41
    TabOrder = 2
    object btnSave: TBitBtn
      Left = 22
      Top = 11
      Width = 73
      Height = 25
      Caption = 'F2.'#20786#23384
      Enabled = False
      TabOrder = 0
      OnClick = btnSaveClick
      NumGlyphs = 2
    end
    object btnPrint: TBitBtn
      Left = 268
      Top = 11
      Width = 75
      Height = 25
      Caption = 'F5.'#21015#21360
      TabOrder = 1
      Visible = False
      OnClick = btnPrintClick
      NumGlyphs = 2
    end
    object btnInsert: TBitBtn
      Left = 102
      Top = 11
      Width = 75
      Height = 25
      Caption = 'F6.'#26032#22686
      TabOrder = 2
      OnClick = btnInsertClick
      NumGlyphs = 2
    end
    object btnEdit: TBitBtn
      Left = 188
      Top = 11
      Width = 75
      Height = 25
      Caption = 'F11.'#20462#25913
      TabOrder = 3
      OnClick = btnEditClick
      NumGlyphs = 2
    end
    object btnDelete: TBitBtn
      Left = 450
      Top = 11
      Width = 75
      Height = 25
      Caption = #21034#38500
      TabOrder = 4
      OnClick = btnDeleteClick
      NumGlyphs = 2
    end
    object btnCloseCancel: TBitBtn
      Left = 533
      Top = 10
      Width = 75
      Height = 25
      Caption = '(&X)'#32080#26463
      TabOrder = 5
      OnClick = btnCloseCancelClick
      NumGlyphs = 2
    end
    object Edit1: TEdit
      Left = 616
      Top = 12
      Width = 36
      Height = 21
      Color = clLime
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      TabOrder = 6
    end
  end
  object dsrSo110: TDataSource
    DataSet = dtmMain2.cdsSo110
    OnDataChange = dsrSo110DataChange
    Left = 24
    Top = 155
  end
  object dsrRef_No: TDataSource
    DataSet = dtmMain2.cdsRefNo
    Left = 392
    Top = 8
  end
end
