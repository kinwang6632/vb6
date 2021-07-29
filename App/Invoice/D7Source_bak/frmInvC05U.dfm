object frmInvC05: TfrmInvC05
  Left = 366
  Top = 161
  BorderStyle = bsDialog
  Caption = 'frmInvC05'
  ClientHeight = 491
  ClientWidth = 509
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
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 509
    Height = 32
    Align = alTop
    Color = clSkyBlue
    TabOrder = 0
    object Label5: TLabel
      Left = 10
      Top = 8
      Width = 119
      Height = 16
      Caption = #32113#19968#30332#31080#26126#32048#34920
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 32
    Width = 509
    Height = 459
    Align = alClient
    TabOrder = 1
    DesignSize = (
      509
      459)
    object Bevel1: TBevel
      Left = 97
      Top = 143
      Width = 396
      Height = 5
      Shape = bsTopLine
    end
    object Label1: TLabel
      Left = 19
      Top = 137
      Width = 72
      Height = 14
      Caption = #34389#29702#30332#31080#23383#36556
    end
    object GroupBox1: TGroupBox
      Left = 18
      Top = 9
      Width = 476
      Height = 120
      Caption = #30332#31080#36039#26009
      TabOrder = 0
      object Label2: TLabel
        Left = 33
        Top = 23
        Width = 52
        Height = 14
        Caption = #30332#31080#24180#26376':'
        Transparent = True
      end
      object Label10: TLabel
        Left = 33
        Top = 51
        Width = 52
        Height = 14
        Caption = #30332#31080#23383#36556':'
        Transparent = True
      end
      object Label7: TLabel
        Left = 33
        Top = 81
        Width = 52
        Height = 14
        Alignment = taRightJustify
        Caption = #30332#31080#26684#24335':'
        Transparent = True
      end
      object edtYearMonth: TcxMaskEdit
        Left = 92
        Top = 21
        Properties.MaskKind = emkRegExprEx
        Properties.EditMask = '([123][0-9])? [0-9][0-9] / (0?[1-9] | 1[012])'
        Properties.MaxLength = 0
        Properties.OnValidate = edtYearMonthPropertiesValidate
        Style.StyleController = dtmMain.cxEditStyle
        TabOrder = 0
        Width = 100
      end
      object edtPrefix: TcxMaskEdit
        Left = 92
        Top = 49
        Properties.CharCase = ecUpperCase
        Properties.MaskKind = emkRegExpr
        Properties.EditMask = '[a-zA-Z][a-zA-Z]'
        Properties.MaxLength = 0
        Style.StyleController = dtmMain.cxEditStyle
        TabOrder = 1
        Width = 70
      end
      object cmbFormat: TcxComboBox
        Left = 92
        Top = 78
        Properties.DropDownListStyle = lsEditFixedList
        Properties.Items.Strings = (
          '1.'#38651#23376#24335
          '2.'#25163#24037#20108#32879
          '3.'#25163#24037#19977#32879)
        Style.StyleController = dtmMain.cxEditStyle
        TabOrder = 2
        Width = 143
      end
      object cxGroupBox1: TcxGroupBox
        Left = 267
        Top = 22
        Caption = ' '#38750#29151#26989#20154#30332#31080#31237#38989#35336#31639#26041#24335'  '
        TabOrder = 3
        Height = 47
        Width = 185
        object rdo1: TcxRadioButton
          Left = 26
          Top = 22
          Width = 60
          Height = 17
          Caption = #20998#38283
          TabOrder = 0
        end
        object rdo2: TcxRadioButton
          Left = 92
          Top = 22
          Width = 60
          Height = 17
          Caption = #21512#20341
          TabOrder = 1
        end
      end
      object btnLoadChooseInv: TcxButton
        Left = 361
        Top = 79
        Width = 90
        Height = 26
        Caption = #23383#36556#26597#35426
        TabOrder = 4
        OnClick = btnLoadChooseInvClick
      end
    end
    object InvTree: TcxTreeList
      Left = 19
      Top = 161
      Width = 473
      Height = 142
      Styles.Background = dtmMain.cxGridBackGroundStyle
      Styles.Content = dtmMain.cxGridBackGroundStyle
      Styles.Inactive = dtmMain.cxGridInActiveStyle
      Styles.Selection = dtmMain.cxGridActiveStyle
      Bands = <
        item
        end>
      BufferedPaint = False
      OptionsBehavior.Sorting = False
      OptionsData.Deleting = False
      OptionsView.GridLines = tlglBoth
      OptionsView.Indicator = True
      OptionsView.ShowRoot = False
      TabOrder = 1
      object TreeColumn1: TcxTreeListColumn
        PropertiesClassName = 'TcxCheckBoxProperties'
        Properties.DisplayChecked = 'Y'
        Properties.DisplayUnchecked = 'N'
        Properties.ImmediatePost = True
        Properties.NullStyle = nssUnchecked
        Properties.ValueChecked = 'Y'
        Properties.ValueUnchecked = 'N'
        Caption.AlignHorz = taCenter
        Caption.Text = #36984#21462
        DataBinding.ValueType = 'String'
        Width = 63
        Position.ColIndex = 0
        Position.RowIndex = 0
        Position.BandIndex = 0
      end
      object TreeColumn2: TcxTreeListColumn
        PropertiesClassName = 'TcxTextEditProperties'
        Properties.Alignment.Horz = taRightJustify
        Caption.AlignHorz = taRightJustify
        Caption.Text = #38918#24207
        DataBinding.ValueType = 'String'
        Options.Editing = False
        Options.Focusing = False
        Width = 58
        Position.ColIndex = 1
        Position.RowIndex = 0
        Position.BandIndex = 0
      end
      object TreeColumn3: TcxTreeListColumn
        Caption.AlignHorz = taCenter
        Caption.Text = #23383#36556
        DataBinding.ValueType = 'String'
        Options.Editing = False
        Options.Focusing = False
        Width = 60
        Position.ColIndex = 2
        Position.RowIndex = 0
        Position.BandIndex = 0
      end
      object TreeColumn4: TcxTreeListColumn
        Caption.Text = #36215#22987#34399
        DataBinding.ValueType = 'String'
        Options.Editing = False
        Options.Focusing = False
        Width = 120
        Position.ColIndex = 3
        Position.RowIndex = 0
        Position.BandIndex = 0
      end
      object TreeColumn5: TcxTreeListColumn
        Caption.Text = #25130#27490#34399
        DataBinding.ValueType = 'String'
        Options.Editing = False
        Options.Focusing = False
        Width = 120
        Position.ColIndex = 4
        Position.RowIndex = 0
        Position.BandIndex = 0
      end
      object TreeColumn6: TcxTreeListColumn
        Visible = False
        Caption.Text = #24050#20351#29992#34399#30908
        DataBinding.ValueType = 'String'
        Position.ColIndex = 5
        Position.RowIndex = 0
        Position.BandIndex = 0
      end
    end
    object cxGroupBox3: TcxGroupBox
      Left = 19
      Top = 311
      Caption = #21015#21360#24335
      TabOrder = 2
      Height = 92
      Width = 474
      object chkOnlyPrintSummary: TcxCheckBox
        Left = 19
        Top = 58
        Caption = #21482#21360#32113#35336#32080#26524
        Properties.OnChange = chkOnlyPrintSummaryPropertiesChange
        TabOrder = 0
        Width = 104
      end
      object cxCheckBox2: TcxCheckBox
        Left = 118
        Top = 58
        Caption = #21482#21360#20351#29992#30332#31080
        TabOrder = 1
        Width = 104
      end
      object PBar: TcxProgressBar
        Left = 231
        Top = 57
        Style.StyleController = dtmMain.cxEditStyle
        TabOrder = 2
        Width = 230
      end
      object Panel3: TPanel
        Left = 14
        Top = 18
        Width = 125
        Height = 32
        BevelOuter = bvNone
        TabOrder = 3
        object rdoReport: TcxRadioButton
          Left = 7
          Top = 8
          Width = 49
          Height = 17
          Caption = #22577#34920
          Checked = True
          TabOrder = 0
          TabStop = True
          OnClick = rdoReportClick
        end
        object rdoFile: TcxRadioButton
          Left = 61
          Top = 8
          Width = 57
          Height = 17
          Caption = #25991#23383#27284
          TabOrder = 1
          OnClick = rdoReportClick
        end
      end
      object txtFile: TcxButtonEdit
        Left = 147
        Top = 23
        Properties.Buttons = <
          item
            Default = True
            Kind = bkEllipsis
          end>
        Properties.OnButtonClick = txtFilePropertiesButtonClick
        Style.StyleController = dtmMain.cxEditStyle
        TabOrder = 4
        Width = 314
      end
    end
    object btnOk: TcxButton
      Left = 165
      Top = 418
      Width = 85
      Height = 29
      Anchors = [akTop, akRight]
      Caption = #30906#23450
      Default = True
      TabOrder = 3
      OnClick = btnOKClick
      Glyph.Data = {
        36060000424D3606000000000000360000002800000020000000100000000100
        18000000000000060000F00A0000F00A00000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FF008B00008200038805008A00FF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF97796797796797796797
        7967FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FF00A300008E0013AC2715B32B06940E009300FF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF977967977967FF00FFFF00FF97
        7967977967FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        00B10000940016AE2E17B23114AD2A13B12906920B009500FF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FFFF
        00FF977967977967FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF00A900
        009A0019B2331CB63A18B33434BC4D17B03013B02A06910B009500FF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FF977967977967FF00FFFF00FFFF00FFFF00FFFF00FFFF00FF0CA316
        1FB73E20BA421FB84011A824018E0162C77119B23213B12B06910B009500FF00
        FFFF00FFFF00FFFF00FFFF00FF977967FF00FFFF00FFFF00FFB89888977967FF
        00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FFFF00FFFF00FF28B338
        42C8651FBB4615B32C00A40000920000930061C87119B13113B22C06900B0095
        00FF00FFFF00FFFF00FFFF00FF977967FF00FFFF00FFB8988897796797796797
        7967FF00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FFFF00FF0EB924
        43C45636BC4E00AF08FF00FFFF00FF00AA0000920063CA7217B13114B12C0691
        09009B00FF00FFFF00FFFF00FF977967FF00FFFF00FF977967977967FF00FFB8
        9888977967FF00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FFFF00FF
        10D2240DBE20009700FF00FFFF00FFFF00FF00A80000930163CA7217B13114B3
        2D069009008F00FF00FFFF00FFFF00FF977967977967977967FF00FFFF00FFFF
        00FFB89888977967FF00FFFF00FFFF00FF977967977967FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF00AD0000940064CB7515B1
        3116B330028904FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFB89888977967FF00FFFF00FFFF00FF977967FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF00AB0001970366CD
        782ABC46008A02FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFB89888977967FF00FFFF00FF977967FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF009A000094
        01049306008400FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FF977967977967977967977967FF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
      NumGlyphs = 2
    end
    object btnClose: TcxButton
      Left = 264
      Top = 418
      Width = 85
      Height = 29
      Anchors = [akTop, akRight]
      Caption = #32080#26463
      ModalResult = 2
      TabOrder = 4
      OnClick = btnCloseClick
      Glyph.Data = {
        36030000424D3603000000000000360000002800000010000000100000000100
        18000000000000030000120B0000120B00000000000000000000FF00FFFF00FF
        FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00
        FFFF00FFFF00FFFF00FFFF00FF0000A40206B0030AAE0000A6000098FF00FFFF
        00FFFF00FFFF00FF0000A700009A0000A7FF00FFFF00FFFF00FFFF00FF0000A9
        1844F6194DF81031D20102AB0000B6FF00FFFF00FF0000B10928D7092ED70313
        B30000ACFF00FFFF00FFFF00FF0103B32451F91F52FF1D4FFF1744E8040BB000
        00B00000AC0D2EDD1142F90D3DF50B3BF0041ABC0000A5FF00FFFF00FF0000AE
        1832DB285BFF2456FF2253FF1B4BF1050DB10F30DD164AFE1344F91041F60E3E
        F60A3CF000009FFF00FFFF00FF0000BE1F37DD3A6FFF2C5EFF295AFF2657FF20
        52FC1C4FFF194AFD1646FA1445FA0F3DF2020AB10000A8FF00FFFF00FFFF00FF
        0000C8121DC83D6AFB3567FF2C5DFF2859FF2253FF1D4EFF1A4DFF123DED0002
        AC0000BAFF00FFFF00FFFF00FFFF00FFFF00FF0000CC0000B62E4EE73668FF2E
        5EFF2859FF2254FF163DEA0000A80000ABFF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FFFF00FF0000BF253FDF3B6DFF3464FF2E5EFF2759FF1B46EA0001AC0000
        A9FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF0203C84B7CFF4170FF3B
        6BFF396CFF2D5EFF2558FF1336D70000B6FF00FFFF00FFFF00FFFF00FFFF00FF
        FF00FF0000D9263CDB5080FF4575FF3662FA0D14C33C6DFF2A5BFF2053FD0B1D
        C40000C0FF00FFFF00FFFF00FFFF00FFFF00FF0000CB527CFA5081FF4B7DFF0B
        13C90000BB0E15C7386AFF2456FF1A4AF20207B30000B5FF00FFFF00FFFF00FF
        FF00FF131CDD6A9CFF5788FF2B46E70000CDFF00FF0000CD0F1BCB3065FF1F51
        FF1439DD0000B1FF00FFFF00FFFF00FFFF00FF0000DE3A52E45782FB0101D0FF
        00FFFF00FFFF00FF0000CC1426D6265AFF0F2EE30103B8FF00FFFF00FFFF00FF
        FF00FFFF00FF0000CF0000C00000CEFF00FFFF00FFFF00FFFF00FF0000C40001
        B80000B5000077FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
        00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF}
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = #25991#23383#27284#26696'(*.txt)|*.txt|'#25152#26377#27284#26696'(*.*)|*.*'
    Left = 56
    Top = 256
  end
end
