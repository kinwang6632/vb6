object frmSO8A20: TfrmSO8A20
  Left = 298
  Top = 168
  Width = 676
  Height = 499
  Caption = #38971#36947#21830#20195#30908
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = #32048#26126#39636
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
    Height = 162
    Align = alTop
    Enabled = False
    TabOrder = 0
    object Label1: TLabel
      Left = 40
      Top = 11
      Width = 31
      Height = 13
      AutoSize = False
      Caption = #20195#30908
      Color = clBtnFace
      FocusControl = dedProviderId
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object Label2: TLabel
      Left = 39
      Top = 34
      Width = 30
      Height = 13
      AutoSize = False
      Caption = #21517#31281
      FocusControl = DBEdit2
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label3: TLabel
      Left = 26
      Top = 56
      Width = 43
      Height = 13
      AutoSize = False
      Caption = #32879#32097#20154
      FocusControl = DBEdit3
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label4: TLabel
      Left = 346
      Top = 57
      Width = 28
      Height = 13
      AutoSize = False
      Caption = #38651#35441
      FocusControl = DBEdit4
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label5: TLabel
      Left = 40
      Top = 80
      Width = 32
      Height = 13
      AutoSize = False
      Caption = #22320#22336
      FocusControl = DBEdit5
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label6: TLabel
      Left = 40
      Top = 104
      Width = 28
      Height = 13
      AutoSize = False
      Caption = #20633#35387
      FocusControl = DBEdit6
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label7: TLabel
      Left = 96
      Top = 130
      Width = 61
      Height = 13
      AutoSize = False
      Caption = #26159#21542#20572#29992
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object Label8: TLabel
      Left = 162
      Top = 12
      Width = 30
      Height = 13
      AutoSize = False
      Caption = #23660#24615
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
    end
    object dedProviderId: TDBEdit
      Left = 74
      Top = 7
      Width = 56
      Height = 21
      DataField = 'PROVIDER_ID'
      DataSource = dsrSo113CT
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 0
    end
    object DBEdit2: TDBEdit
      Left = 74
      Top = 30
      Width = 586
      Height = 21
      DataField = 'PROVIDER_NAME'
      DataSource = dsrSo113CT
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 1
    end
    object DBEdit3: TDBEdit
      Left = 74
      Top = 54
      Width = 264
      Height = 21
      DataField = 'CONTACTEE'
      DataSource = dsrSo113CT
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 2
    end
    object DBEdit4: TDBEdit
      Left = 377
      Top = 54
      Width = 282
      Height = 21
      DataField = 'TEL'
      DataSource = dsrSo113CT
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 3
    end
    object DBEdit5: TDBEdit
      Left = 74
      Top = 78
      Width = 586
      Height = 21
      DataField = 'ADDRESS'
      DataSource = dsrSo113CT
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 4
    end
    object DBEdit6: TDBEdit
      Left = 74
      Top = 102
      Width = 586
      Height = 21
      DataField = 'NOTES'
      DataSource = dsrSo113CT
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 5
    end
    object DBCheckBox1: TDBCheckBox
      Left = 74
      Top = 128
      Width = 17
      Height = 17
      DataField = 'STOPFLAG'
      DataSource = dsrSo113CT
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 6
      ValueChecked = '1'
      ValueUnchecked = '0'
    end
    object CobAttribute: TComboBox
      Left = 200
      Top = 8
      Width = 213
      Height = 21
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #32048#26126#39636
      Font.Style = []
      ItemHeight = 13
      ParentFont = False
      TabOrder = 7
      OnChange = CobAttributeChange
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 162
    Width = 668
    Height = 263
    Align = alTop
    BevelOuter = bvLowered
    TabOrder = 1
    object dgr1: TDBGrid
      Left = 1
      Top = 1
      Width = 666
      Height = 241
      Align = alClient
      DataSource = dsrSo113CT
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      ReadOnly = True
      TabOrder = 1
      TitleFont.Charset = ANSI_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -13
      TitleFont.Name = #32048#26126#39636
      TitleFont.Style = []
      OnCellClick = dgr1CellClick
      OnDblClick = dgr1DblClick
      OnEnter = dgr1Enter
      Columns = <
        item
          Expanded = False
          FieldName = 'PROVIDER_ID'
          Width = 70
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'PROVIDER_NAME'
          Width = 99
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'ATTRIBUTE'
          Visible = False
        end
        item
          Expanded = False
          FieldName = 'ATTRIBUTE_NAME'
          Width = 172
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'CONTACTEE'
          Width = 94
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'TEL'
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'ADDRESS'
          Width = 325
          Visible = True
        end
        item
          Expanded = False
          FieldName = 'NOTES'
          Width = 205
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
          Width = 85
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
      Top = 242
      Width = 666
      Height = 20
      Color = clCream
      Panels = <
        item
          Width = 692
        end>
    end
  end
  object GroupBox2: TGroupBox
    Left = 0
    Top = 429
    Width = 668
    Height = 42
    Align = alBottom
    TabOrder = 2
    object btnSave: TBitBtn
      Left = 13
      Top = 11
      Width = 75
      Height = 25
      Caption = 'F2.'#20786#23384
      Enabled = False
      TabOrder = 0
      OnClick = btnSaveClick
      NumGlyphs = 2
    end
    object btnPrint: TBitBtn
      Left = 256
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
      Left = 93
      Top = 11
      Width = 75
      Height = 25
      Caption = 'F6.'#26032#22686
      TabOrder = 2
      OnClick = btnInsertClick
      NumGlyphs = 2
    end
    object btnEdit: TBitBtn
      Left = 176
      Top = 11
      Width = 75
      Height = 25
      Caption = 'F11.'#20462#25913
      TabOrder = 3
      OnClick = btnEditClick
      NumGlyphs = 2
    end
    object btnDelete: TBitBtn
      Left = 429
      Top = 11
      Width = 75
      Height = 25
      Caption = #21034#38500
      TabOrder = 4
      OnClick = btnDeleteClick
      NumGlyphs = 2
    end
    object btnCloseCancel: TBitBtn
      Left = 512
      Top = 11
      Width = 75
      Height = 25
      Caption = '(&X)'#32080#26463
      TabOrder = 5
      OnClick = btnCloseCancelClick
      NumGlyphs = 2
    end
    object Edit1: TEdit
      Left = 594
      Top = 12
      Width = 36
      Height = 21
      Color = clLime
      ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
      TabOrder = 6
    end
  end
  object dsrSo113CT: TDataSource
    DataSet = dtmMain2.cdsSo113CT
    OnDataChange = dsrSo113CTDataChange
    Left = 272
    Top = 112
  end
  object DataSource1: TDataSource
    DataSet = dtmMain2.cdsAttribute
    Left = 24
    Top = 16
  end
end
