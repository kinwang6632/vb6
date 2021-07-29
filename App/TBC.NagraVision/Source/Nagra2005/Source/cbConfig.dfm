object fmConfig: TfmConfig
  Left = 335
  Top = 112
  BorderStyle = bsDialog
  Caption = #36984#38917
  ClientHeight = 418
  ClientWidth = 692
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnKeyDown = FormKeyDown
  PixelsPerInch = 96
  TextHeight = 14
  object Bevel1: TBevel
    Left = 0
    Top = 374
    Width = 692
    Height = 3
    Align = alBottom
    Shape = bsBottomLine
  end
  object ConfigPage: TcxPageControl
    Left = 0
    Top = 0
    Width = 692
    Height = 374
    ActivePage = cxCAS
    Align = alClient
    LookAndFeel.Kind = lfOffice11
    Style = 9
    TabOrder = 0
    OnChange = ConfigPageChange
    ClientRectBottom = 374
    ClientRectRight = 692
    ClientRectTop = 21
    object cxCAS: TcxTabSheet
      Caption = 'CAS'#20027#27231'   '
      ImageIndex = 0
      object lblCASIP: TLabel
        Left = 27
        Top = 71
        Width = 28
        Height = 14
        Caption = #20027#27231':'
        Transparent = True
      end
      object lblControlPort: TLabel
        Left = 186
        Top = 71
        Width = 88
        Height = 14
        Caption = #20659#36865#25351#20196#36890#35338#22496':'
        Transparent = True
      end
      object lblFeedbackPort: TLabel
        Left = 380
        Top = 71
        Width = 88
        Height = 14
        Caption = #22238#20659#36039#26009#36890#35338#22496':'
        Transparent = True
      end
      object lblSourceId: TLabel
        Left = 85
        Top = 100
        Width = 57
        Height = 14
        Caption = 'Source Id:'
        Transparent = True
      end
      object lblDestId: TLabel
        Left = 260
        Top = 100
        Width = 44
        Height = 14
        Caption = 'Dest Id:'
        Transparent = True
      end
      object lblMopPpid: TLabel
        Left = 412
        Top = 100
        Width = 56
        Height = 14
        Caption = 'Mop PPId:'
        Transparent = True
      end
      object lblSendCommandDelay: TLabel
        Left = 318
        Top = 258
        Width = 220
        Height = 14
        Caption = #27599#20659#36865#19968#31558#20302#38542#25351#20196#24460#24310#36978'                '#31186
        Transparent = True
      end
      object lblCasRetryFrquence: TLabel
        Left = 28
        Top = 189
        Width = 300
        Height = 14
        Caption = #28961#27861#33287#20027#27231#36899#32218#25110#20027#27231#24537#30860#26178', '#27599'                '#31186#37325#35430#36899#32218
        Transparent = True
      end
      object lblCasSendErrMax: TLabel
        Left = 28
        Top = 223
        Width = 240
        Height = 14
        Caption = #20659#36865#25351#20196', '#30332#29983'                '#27425#37679#35492#21363#20572#27490#22519#34892
        Transparent = True
      end
      object lblCasRecvErrMax: TLabel
        Left = 318
        Top = 224
        Width = 240
        Height = 14
        Caption = #22238#20659#36039#26009', '#30332#29983'                '#27425#37679#35492#21363#20572#27490#22519#34892
        Transparent = True
      end
      object lblCasCommCheck: TLabel
        Left = 28
        Top = 259
        Width = 262
        Height = 14
        Caption = #27599'                 '#31186#20659#36865#19968#27425' Commnucation Check'
        Transparent = True
      end
      object Label1: TLabel
        Left = 27
        Top = 101
        Width = 48
        Height = 14
        Caption = #20659#36865#25351#20196
      end
      object Label3: TLabel
        Left = 84
        Top = 124
        Width = 57
        Height = 14
        Caption = 'Source Id:'
        Transparent = True
      end
      object Label4: TLabel
        Left = 260
        Top = 127
        Width = 44
        Height = 14
        Caption = 'Dest Id:'
        Transparent = True
      end
      object Label5: TLabel
        Left = 412
        Top = 127
        Width = 56
        Height = 14
        Caption = 'Mop PPId:'
        Transparent = True
      end
      object Label6: TLabel
        Left = 27
        Top = 128
        Width = 48
        Height = 14
        Caption = #22238#20659#36039#26009
      end
      object Label7: TLabel
        Left = 28
        Top = 158
        Width = 136
        Height = 14
        Caption = #20659#36865#25351#20196#20132#26131#24207#34399#35672#21029#30908':'
        Transparent = True
      end
      object Label8: TLabel
        Left = 272
        Top = 158
        Width = 136
        Height = 14
        Caption = #22238#20659#36039#26009#20132#26131#24207#34399#35672#21029#30908':'
        Transparent = True
      end
      object Label9: TLabel
        Left = 348
        Top = 189
        Width = 136
        Height = 14
        Caption = #20302#38542#25351#20196#20659#36865#35336#25976#22120#19978#38480':'
        Transparent = True
      end
      object txtCASIP: TcxMaskEdit
        Left = 59
        Top = 68
        Enabled = False
        ParentFont = False
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 0
        Text = 'txtCASIP'
        Width = 116
      end
      object txtControlPort: TcxMaskEdit
        Left = 280
        Top = 68
        Enabled = False
        ParentFont = False
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 1
        Text = 'txtControlPort'
        Width = 70
      end
      object txtFeedbackPort: TcxMaskEdit
        Left = 474
        Top = 68
        Enabled = False
        ParentFont = False
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 2
        Text = 'txtFeedbackPort'
        Width = 70
      end
      object txtControlSourceId: TcxMaskEdit
        Left = 147
        Top = 96
        Enabled = False
        ParentFont = False
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 3
        Text = 'txtFeedbackPort'
        Width = 86
      end
      object txtControlDestId: TcxMaskEdit
        Left = 309
        Top = 96
        Enabled = False
        ParentFont = False
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 4
        Text = 'txtFeedbackPort'
        Width = 84
      end
      object txtControlMopPpid: TcxMaskEdit
        Left = 474
        Top = 96
        Enabled = False
        ParentFont = False
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 5
        Text = 'txtFeedbackPort'
        Width = 84
      end
      object chkEnableControlSend: TcxCheckBox
        Left = 27
        Top = 291
        Caption = #21855#29992#25351#20196#20659#36865':'
        ParentFont = False
        Properties.Alignment = taRightJustify
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 14
        Width = 102
      end
      object chkEnableFeedbackRecv: TcxCheckBox
        Left = 147
        Top = 291
        Caption = #21855#29992#22238#20659#27231#21046':'
        ParentFont = False
        Properties.Alignment = taRightJustify
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 15
        Width = 102
      end
      object txtSendCommandDelay: TcxSpinEdit
        Left = 469
        Top = 255
        ParentFont = False
        Properties.AssignedValues.MaxValue = True
        Properties.Increment = 0.100000000000000000
        Properties.LargeIncrement = 1.000000000000000000
        Properties.MinValue = 10.000000000000000000
        Properties.ValueType = vtFloat
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 13
        Value = 0.500000000000000000
        Width = 50
      end
      object chkUseSimulator: TcxCheckBox
        Left = 267
        Top = 291
        Caption = #20351#29992#27169#25836#22120':'
        Enabled = False
        ParentFont = False
        Properties.Alignment = taRightJustify
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 16
        Width = 90
      end
      object txtCasRetryFrquence: TcxSpinEdit
        Left = 211
        Top = 185
        ParentFont = False
        Properties.MaxValue = 600.000000000000000000
        Properties.MinValue = 3.000000000000000000
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 9
        Value = 3
        Width = 52
      end
      object txtCasSendErrMax: TcxSpinEdit
        Left = 113
        Top = 221
        ParentFont = False
        Properties.MaxValue = 10.000000000000000000
        Properties.MinValue = 1.000000000000000000
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 10
        Value = 3
        Width = 52
      end
      object txtCasRecvErrMax: TcxSpinEdit
        Left = 403
        Top = 220
        ParentFont = False
        Properties.MaxValue = 10.000000000000000000
        Properties.MinValue = 1.000000000000000000
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 11
        Value = 3
        Width = 52
      end
      object txtCasCommCheck: TcxSpinEdit
        Left = 50
        Top = 255
        ParentFont = False
        Properties.MaxValue = 600.000000000000000000
        Properties.MinValue = 1.000000000000000000
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 12
        Value = 3
        Width = 52
      end
      object HeaderPanel: TPanel
        Left = 9
        Top = 5
        Width = 675
        Height = 42
        BevelOuter = bvNone
        Color = clMedGray
        TabOrder = 17
        object HeaderImage: TImage
          Left = 11
          Top = 6
          Width = 32
          Height = 32
          AutoSize = True
          Transparent = True
        end
      end
      object txtFeedbackSourceId: TcxMaskEdit
        Left = 147
        Top = 123
        Enabled = False
        ParentFont = False
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 6
        Text = 'txtFeedbackPort'
        Width = 86
      end
      object txtFeedbackDestId: TcxMaskEdit
        Left = 309
        Top = 123
        Enabled = False
        ParentFont = False
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 7
        Text = 'txtFeedbackPort'
        Width = 84
      end
      object txtFeedbackMopPpid: TcxMaskEdit
        Left = 474
        Top = 123
        Enabled = False
        ParentFont = False
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 8
        Text = 'txtFeedbackPort'
        Width = 84
      end
      object txtControlTransId: TcxMaskEdit
        Left = 169
        Top = 155
        Enabled = False
        ParentFont = False
        Properties.MaskKind = emkRegExpr
        Properties.EditMask = '(\d)|(\d\d)'
        Properties.MaxLength = 0
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 18
        Width = 91
      end
      object txtFeedbackTransId: TcxMaskEdit
        Left = 415
        Top = 155
        Enabled = False
        ParentFont = False
        Properties.MaskKind = emkRegExpr
        Properties.EditMask = '(\d)|(\d\d)'
        Properties.MaxLength = 0
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 19
        Width = 92
      end
      object txtCasCmdMaxCounter: TcxSpinEdit
        Left = 490
        Top = 185
        ParentFont = False
        Properties.AssignedValues.MinValue = True
        Properties.MaxValue = 999999.000000000000000000
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 20
        Value = 1000
        Width = 75
      end
    end
    object cxCommon: TcxTabSheet
      Caption = #19968#33324#35373#23450'   '
      ImageIndex = 1
      object lbDbRetryFrequence: TLabel
        Left = 24
        Top = 99
        Width = 276
        Height = 14
        Caption = #31995#32113#21488#36039#26009#24235#26039#32218#26178', '#31561#20505'                  '#31186#37325#35430#36899#32218':'
        Transparent = True
      end
      object lblBusyTimeStart: TLabel
        Left = 24
        Top = 126
        Width = 125
        Height = 14
        Caption = #24537#30860#26178#27573' --> '#36215#22987#26178#38291':'
        Transparent = True
      end
      object lblBusyTimeEnd: TLabel
        Left = 253
        Top = 126
        Width = 52
        Height = 14
        Caption = #32080#27490#26178#38291':'
        Transparent = True
      end
      object lblBusyTimeReadFrquence: TLabel
        Left = 661
        Top = 79
        Width = 180
        Height = 14
        Caption = #27599'                  '#31186#35712#21462#31995#32113#21488#36039#26009
      end
      object lblNormalTime: TLabel
        Left = 24
        Top = 154
        Width = 257
        Height = 14
        Caption = #19968#33324#26178#27573' --> '#27599'                  '#31186#35712#21462#31995#32113#21488#36039#26009','
        Transparent = True
      end
      object lblProcessRecordCount: TLabel
        Left = 288
        Top = 154
        Width = 180
        Height = 14
        Caption = #27599#27425#34389#29702'                  '#31558#39640#38542#25351#20196
        Transparent = True
      end
      object Label2: TLabel
        Left = 406
        Top = 126
        Width = 184
        Height = 14
        Caption = #27599'                  '#31186#35712#21462#31995#32113#21488#36039#26009','
        Transparent = True
      end
      object lblDbWriteErrorWhenSocketFail: TLabel
        Left = 24
        Top = 181
        Width = 204
        Height = 14
        Caption = #28961#27861#20659#36865#25351#20196#26178', '#21508#31995#32113#21488#24453#34389#29702#36039#26009':'
        Transparent = True
      end
      object lblProcessIPPV: TLabel
        Left = 24
        Top = 208
        Width = 74
        Height = 14
        Caption = #34389#29702'PPV'#36039#26009':'
        Transparent = True
      end
      object lblProcessBatch: TLabel
        Left = 24
        Top = 234
        Width = 76
        Height = 14
        Caption = #34389#29702#25209#27425#36039#26009':'
        Transparent = True
      end
      object lblBatchOperator: TLabel
        Left = 312
        Top = 234
        Width = 88
        Height = 14
        Caption = #25209#27425#36039#26009#25805#20316#21729':'
        Transparent = True
      end
      object chkDbMultiThread: TcxCheckBox
        Left = 19
        Top = 70
        Caption = #21508#31995#32113#21488#20351#29992#29544#31435#22519#34892#32210':'
        Enabled = False
        ParentColor = False
        ParentFont = False
        Properties.Alignment = taRightJustify
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 0
        Width = 162
      end
      object txtDbRetryFrequence: TcxSpinEdit
        Left = 171
        Top = 96
        ParentFont = False
        Properties.MaxValue = 600.000000000000000000
        Properties.MinValue = 10.000000000000000000
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 1
        Value = 10
        Width = 58
      end
      object txtBusyTimeStart: TcxTimeEdit
        Left = 160
        Top = 123
        EditValue = 0.000000000000000000
        ParentFont = False
        Properties.TimeFormat = tfHourMin
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 2
        Width = 85
      end
      object txtBusyTimeEnd: TcxTimeEdit
        Left = 311
        Top = 123
        EditValue = 0.000000000000000000
        ParentFont = False
        Properties.TimeFormat = tfHourMin
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 3
        Width = 85
      end
      object txtlBusyTimeReadFrquence: TcxSpinEdit
        Left = 426
        Top = 123
        ParentFont = False
        Properties.MaxValue = 60.000000000000000000
        Properties.MinValue = 2.000000000000000000
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 4
        Value = 3
        Width = 58
      end
      object txtNormalTime: TcxSpinEdit
        Left = 117
        Top = 151
        ParentFont = False
        Properties.MaxValue = 60.000000000000000000
        Properties.MinValue = 2.000000000000000000
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 5
        Value = 3
        Width = 58
      end
      object txtProcessRecordCount: TcxSpinEdit
        Left = 342
        Top = 151
        ParentFont = False
        Properties.MaxValue = 1000.000000000000000000
        Properties.MinValue = 1.000000000000000000
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 6
        Value = 20
        Width = 58
      end
      object Panel2: TPanel
        Left = 238
        Top = 178
        Width = 185
        Height = 24
        BevelOuter = bvNone
        TabOrder = 7
        object chkWriteNo: TcxRadioButton
          Left = 3
          Top = 3
          Width = 63
          Height = 17
          Caption = #19981#34389#29702
          Checked = True
          TabOrder = 0
          TabStop = True
          LookAndFeel.Kind = lfFlat
        end
        object chkWriteYes: TcxRadioButton
          Left = 68
          Top = 3
          Width = 109
          Height = 17
          Caption = #30452#25509#27161#31034#25104#37679#35492
          TabOrder = 1
          LookAndFeel.Kind = lfFlat
        end
      end
      object Panel3: TPanel
        Left = 108
        Top = 205
        Width = 194
        Height = 24
        BevelOuter = bvNone
        ParentBackground = False
        TabOrder = 8
        object chkProcessIPPV_A: TcxRadioButton
          Left = 3
          Top = 3
          Width = 50
          Height = 17
          Caption = #20840#37096
          Checked = True
          TabOrder = 0
          TabStop = True
          LookAndFeel.Kind = lfFlat
        end
        object chkProcessIPPV_N: TcxRadioButton
          Left = 55
          Top = 3
          Width = 60
          Height = 17
          Caption = #19981#34389#29702
          TabOrder = 1
          LookAndFeel.Kind = lfFlat
        end
        object chkProcessIPPV_O: TcxRadioButton
          Left = 120
          Top = 3
          Width = 60
          Height = 17
          Caption = #21482#34389#29702
          TabOrder = 2
          LookAndFeel.Kind = lfFlat
        end
      end
      object Panel4: TPanel
        Left = 108
        Top = 230
        Width = 194
        Height = 24
        BevelOuter = bvNone
        TabOrder = 9
        object ChkProcessBatch_A: TcxRadioButton
          Left = 3
          Top = 3
          Width = 50
          Height = 17
          Caption = #20840#37096
          Checked = True
          TabOrder = 0
          TabStop = True
          LookAndFeel.Kind = lfFlat
        end
        object ChkProcessBatch_N: TcxRadioButton
          Left = 55
          Top = 3
          Width = 60
          Height = 17
          Caption = #19981#34389#29702
          TabOrder = 1
          LookAndFeel.Kind = lfFlat
        end
        object ChkProcessBatch_O: TcxRadioButton
          Left = 120
          Top = 3
          Width = 60
          Height = 17
          Caption = #21482#34389#29702
          TabOrder = 2
          LookAndFeel.Kind = lfFlat
        end
      end
      object txtBatchOperator: TcxTextEdit
        Left = 403
        Top = 231
        ParentFont = False
        Style.StyleController = StyleModule.cxEditStyle
        TabOrder = 10
        Text = 'CABLESOFT'
        Width = 126
      end
    end
    object cxSoDb: TcxTabSheet
      Caption = #31995#32113#21488#36039#26009#24235'   '
      ImageIndex = 2
      object lblTest: TLabel
        Left = 247
        Top = 63
        Width = 409
        Height = 33
        AutoSize = False
        Caption = 'lblTest'
        Transparent = True
        WordWrap = True
      end
      object SoTree: TcxTreeList
        Left = 12
        Top = 104
        Width = 666
        Height = 235
        Bands = <
          item
            FixedKind = tlbfLeft
          end
          item
          end>
        BufferedPaint = True
        Images = StyleModule.SoTreeImageList
        LookAndFeel.Kind = lfOffice11
        OptionsBehavior.GoToNextCellOnEnter = True
        OptionsBehavior.Sorting = False
        OptionsView.CellEndEllipsis = True
        OptionsView.CategorizedColumn = SoTreeColumn3
        OptionsView.Indicator = True
        OptionsView.PaintStyle = tlpsCategorized
        Styles.Content = StyleModule.cxStyle1
        TabOrder = 0
        OnEdited = SoTreeEdited
        OnFocusedNodeChanged = SoTreeFocusedNodeChanged
        object SoTreeColumn1: TcxTreeListColumn
          PropertiesClassName = 'TcxCheckBoxProperties'
          Properties.ImmediatePost = True
          Properties.NullStyle = nssUnchecked
          Properties.ValueChecked = 'Y'
          Properties.ValueUnchecked = 'N'
          Caption.AlignHorz = taCenter
          Caption.Text = #36984#21462
          DataBinding.ValueType = 'String'
          Position.ColIndex = 0
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
        object SoTreeColumn2: TcxTreeListColumn
          PropertiesClassName = 'TcxTextEditProperties'
          Caption.Text = #20844#21496#20195#30908
          DataBinding.ValueType = 'String'
          Width = 64
          Position.ColIndex = 2
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
        object SoTreeColumn3: TcxTreeListColumn
          PropertiesClassName = 'TcxTextEditProperties'
          Caption.Text = #31995#32113#21488#21517#31281
          DataBinding.ValueType = 'String'
          Width = 94
          Position.ColIndex = 3
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
        object SoTreeColumn4: TcxTreeListColumn
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taCenter
          Caption.Text = #32178#36335#32232#34399
          DataBinding.ValueType = 'String'
          Width = 66
          Position.ColIndex = 0
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
        object SoTreeColumn5: TcxTreeListColumn
          PropertiesClassName = 'TcxTextEditProperties'
          Visible = False
          Caption.Text = #24115#34399
          DataBinding.ValueType = 'String'
          Width = 75
          Position.ColIndex = 3
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
        object SoTreeColumn6: TcxTreeListColumn
          PropertiesClassName = 'TcxTextEditProperties'
          Visible = False
          Caption.Text = #23494#30908
          DataBinding.ValueType = 'String'
          Width = 78
          Position.ColIndex = 4
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
        object SoTreeColumn7: TcxTreeListColumn
          PropertiesClassName = 'TcxTextEditProperties'
          Visible = False
          Caption.Text = #36039#26009#24235#21029#21517
          DataBinding.ValueType = 'String'
          Width = 144
          Position.ColIndex = 2
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
        object SoTreeColumn8: TcxTreeListColumn
          Visible = False
          Caption.Text = 'SO_TYPE'
          DataBinding.ValueType = 'String'
          Position.ColIndex = 6
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
        object SoTreeColumn9: TcxTreeListColumn
          Visible = False
          Caption.Text = 'SO_POS'
          DataBinding.ValueType = 'String'
          Position.ColIndex = 5
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
      end
      object btnDeleteSo: TcxButton
        Left = 99
        Top = 64
        Width = 65
        Height = 26
        Hint = #21034#38500#36984#21462#30340#31995#32113#21488
        Caption = #21034#38500
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = btnDeleteSoClick
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          A4140206B0CB030AAEE10000A67B000098070000000000000000000000000000
          00000000A78500009A8B0000A72A000000000000000000000000000000000000
          A98D1844F6FF194DF8FF1031D2FF0102ABD00000B62400000000000000010000
          B1930928D7FF092ED7FF0313B3FF0000AC8F0000000000000000000000000103
          B3D02451F9FF1F52FFFF1D4FFFFF1744E8FF040BB0EC0000B03A0000AC970D2E
          DDFF1142F9FF0D3DF5FF0B3BF0FF041ABCFF0000A55D00000000000000000000
          AE421832DBFF285BFFFF2456FFFF2253FFFF1B4BF1FF050DB1FB0F30DDFF164A
          FEFF1344F9FF1041F6FF0E3EF6FF0A3CF0FF00009FBF00000000000000000000
          BE511F37DDFF3A6FFFFF2C5EFFFF295AFFFF2657FFFF2052FCFF1C4FFFFF194A
          FDFF1646FAFF1445FAFF0F3DF2FF020AB1E70000A84300000000000000000000
          00010000C866121DC8F03D6AFBFF3567FFFF2C5DFFFF2859FFFF2253FFFF1D4E
          FFFF1A4DFFFF123DEDFF0002ACCB0000BA1A0000000000000000000000000000
          0000000000000000CC1C0000B6BA2E4EE7FF3668FFFF2E5EFFFF2859FFFF2254
          FFFF163DEAFF0000A8BE0000AB0B000000000000000000000000000000000000
          000000000000000000000000BF4D253FDFFF3B6DFFFF3464FFFF2E5EFFFF2759
          FFFF1B46EAFF0001ACCF0000A908000000000000000000000000000000000000
          000000000000000000000203C8C14B7CFFFF4170FFFF3B6BFFFF396CFFFF2D5E
          FFFF2558FFFF1336D7FF0000B693000000000000000000000000000000000000
          0000000000000000D92C263CDBFE5080FFFF4575FFFF3662FAFF0D14C3FE3C6D
          FFFF2A5BFFFF2053FDFF0B1DC4FD0000C0470000000000000000000000000000
          0000000000000000CB8B527CFAFF5081FFFF4B7DFFFF0B13C9EA0000BB420E15
          C7E9386AFFFF2456FFFF1A4AF2FF0207B3E30000B50F00000000000000000000
          000000000000131CDDE06A9CFFFF5788FFFF2B46E7FF0000CD5C000000000000
          CD320F1BCBF23065FFFF1F51FFFF1439DDFF0000B19C00000000000000000000
          0000000000000000DE913A52E4FE5782FBFF0101D0C200000000000000000000
          00000000CC471426D6FA265AFFFF0F2EE3FF0103B8CF00000000000000000000
          000000000000000000000000CF3C0000C0A40000CE3000000000000000000000
          0000000000000000C4630001B8BC0000B5490000770200000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
        LookAndFeel.Kind = lfFlat
      end
      object btnAddSo: TcxButton
        Left = 28
        Top = 64
        Width = 65
        Height = 26
        Hint = #26032#22686#31995#32113#21488
        Caption = #26032#22686
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = btnAddSoClick
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000038307CE0E8E18FF0B8A15FF0A8814FF0985
          12FF007703C70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000048609CE51DA7BFF3ACF69FF39CD67FF32C2
          5BFF017A04C70000000000000000000000000000000000000000000000000000
          000000000000000000000000000005870ACE5CE084FF3ED46EFF3DD36DFF34C5
          5FFF027C05C70000000000000000000000000000000000000000000000000000
          000000000000000000000000000006890CCE64E48AFF41D771FF3FD56FFF37C8
          61FF027D06C70000000000000000000000000000000000000000000000000893
          10D4089211D4079010D4078E0FD4088E10F056E282FF44DA74FF41D872FF39CB
          64FF038209EC037E07CE037C06CE027A05CE017804CE000000000000000021A7
          2DFF60E487FF4BDD79FF4ADC78FF49DC77FF4AE07AFF46DD77FF44DB75FF40D6
          70FF3BCD67FF39CB65FF36C862FF33C45EFF0A8413FF000000000000000022A9
          2FFF76F099FF5EEA8AFF5AE888FF56E684FF53E481FF4EE17DFF47DE78FF44DB
          75FF41D872FF3FD56FFF3DD36DFF39CD67FF0B8715FF000000000000000024AB
          30FF7EF39FFF68EE91FF64ED8EFF60EA8BFF59E886FF54E482FF4EE17DFF46DD
          77FF44DA74FF41D771FF3ED46EFF3ACF69FF0C8916FF000000000000000025AD
          32FF91F7ABFF8DF6A8FF8BF5A6FF89F4A5FF7AF09BFF59E886FF53E481FF4ADF
          7AFF5BE486FF67E58CFF5EE186FF53DB7CFF0F8D1AFF0000000000000000099D
          14CE0A9B14CE099913CE099713CE0B9615EF89F4A5FF60EA8BFF56E684FF45D8
          72FF078D0FEC06890DCE06860CCE05840BCE048209CE00000000000000000000
          0000000000000000000000000000099713CE8BF5A6FF64ED8EFF5AE888FF45D7
          71FF068B0DC70000000000000000000000000000000000000000000000000000
          00000000000000000000000000000A9914CE8DF6A8FF68EE91FF5EEA8AFF46D8
          72FF068E0EC70000000000000000000000000000000000000000000000000000
          00000000000000000000000000000B9A15CE91F7ABFF7FF39FFF76F099FF58DF
          7FFF078F0FC70000000000000000000000000000000000000000000000000000
          00000000000000000000000000000A9C16CE26AD33FF25AB32FF23A930FF21A6
          2EFF079110C70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
        LookAndFeel.Kind = lfFlat
      end
      object btnTest: TcxButton
        Left = 170
        Top = 64
        Width = 65
        Height = 26
        Hint = #28204#35430#31995#32113#21488#36899#32080
        Caption = #28204#35430
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
        OnClick = btnTestClick
        Glyph.Data = {
          36050000424D3605000000000000360400002800000010000000100000000100
          0800000000000001000000000000000000000001000000000000000000000000
          80000080000000808000800000008000800080800000C0C0C000808080000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00080808001717
          1700272727003737370047474700575757006767670077777700878787009797
          9700A7A7A700B7B7B700C7C7C700D7D7D700E7E7E700F7F7F70000004C000000
          7F000000B2000000E5001919FF004C4CFE007F7FFF00B2B2FF0000104C00001B
          7F000026B2000031E500194AFF004C72FE007F9AFF00B2C2FF0000204C000036
          7F00004CB2000062E500197BFF004C99FE007FB6FF00B2D3FF0000314C000051
          7F000072B2000093E50019ADFF004CBFFE007FD1FF00B2E3FF0000414C00006D
          7F000099B20000C4E50019DEFF004CE5FE007FECFF00B2F4FF00004C4700007F
          760000B2A50000E5D50019FFEE004CFEF2007FFFF500B2FFF900004C3600007F
          5B0000B27F0000E5A30019FFBD004CFECC007FFFDA00B2FFE900004C2600007F
          3F0000B2590000E5720019FF8C004CFEA5007FFFBF00B2FFD800004C1500007F
          240000B2330000E5410019FF5B004CFE7F007FFFA300B2FFC800004C0500007F
          090000B20C0000E5100019FF29004CFE59007FFF8800B2FFB7000A4C0000127F
          000019B2000020E500003AFF190066FE4C0091FF7F00BDFFB2001B4C00002D7F
          00003FB2000051E500006BFF19008CFE4C00ADFF7F00CDFFB2002B4C0000487F
          000065B2000083E500009CFF1900B2FE4C00C8FF7F00DEFFB2003C4C0000647F
          00008CB20000B4E50000CDFF1900D8FE4C00E3FF7F00EEFFB2004C4C00007F7F
          0000B2B20000E5E50000FFFF1900FEFE4C00FFFF7F00FFFFB2004C3C00007F64
          0000B28C0000E5B40000FFCD1900FED84C00FFE37F00FFEEB2004C2B00007F48
          0000B2660000E5830000FF9C1900FEB24C00FFC87F00FFDEB2004C1B00007F2D
          0000B23F0000E5510000FF6B1900FE8C4C00FFAD7F00FFCDB2004C0A00007F12
          0000B2190000E5200000FF3A1900FE654C00FF917F00FFBDB2004C0005007F00
          0900B2000C00E5001000FF192900FE4C5900FF7F8800FFB2B7004C0015007F00
          2400B2003200E5004100FF195B00FE4C7F00FF7FA300FFB2C8004C0026007F00
          3F00B2005900E5007200FF198C00FE4CA500FF7FBF00FFB2D8004C0036007F00
          5B00B2007F00E500A300FF19BD00FE4CCC00FF7FDA00FFB2E9004C0047007F00
          7600B200A500E500D500FF19EE00FE4CF200FF7FF500FFB2F90041004C006D00
          7F009900B200C400E500DE19FF00E54CFE00EC7FFF00F4B2FF0031004C005100
          7F007200B2009300E500AD19FF00BF4CFE00D17FFF00E3B2FF0020004C003600
          7F004C00B2006200E5007B19FF00994CFE00B67FFF00D3B2FF0010004C001B00
          7F002600B2003100E5004A19FF00724CFE009A7FFF00C2B2FF000F0F0F0F0F0F
          0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0000001419
          0F0F00000F0F0F0F0F0F0F0F0F0F0F0F0F0F4600000F0F0F0F0F0F000000140F
          0F0F074600000F0F0F0F0F0F0F0F0F0F0F0F19074600000F0F0F0F0000001419
          0F0F0F19074600000F0F0F0F0F0F0F0F0F16000000074600000F0F000000140F
          0F1907464646464600000F0F0F0F0F0F0F0F19074600000F0F0F0F0000001419
          0F0F0F19074600000F0F0F0F0F0F0F0F0F0F0F19070F4600000F0F0000001419
          0F0F0F0F19070F4600000F0F0F0F0F0F0F0F0F0F0F19070F46000F0F0F0F0F0F
          0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F0F}
        LookAndFeel.Kind = lfFlat
      end
    end
    object cxHighCmd: TcxTabSheet
      Caption = #39640#38542#25351#20196#23565#29031#34920'   '
      ImageIndex = 3
      object HighCmdTree: TcxTreeList
        Left = 12
        Top = 100
        Width = 666
        Height = 235
        Bands = <
          item
            FixedKind = tlbfLeft
          end
          item
          end>
        BufferedPaint = True
        Images = StyleModule.CmdTreeImageList
        LookAndFeel.Kind = lfOffice11
        OptionsBehavior.GoToNextCellOnEnter = True
        OptionsView.CellEndEllipsis = True
        OptionsView.GridLines = tlglBoth
        OptionsView.Indicator = True
        OptionsView.TreeLineStyle = tllsNone
        Styles.Content = StyleModule.cxStyle1
        TabOrder = 0
        OnEdited = SoTreeEdited
        OnFocusedNodeChanged = SoTreeFocusedNodeChanged
        object HighCmdTreeColumn1: TcxTreeListColumn
          Caption.Text = #39640#38542#25351#20196
          DataBinding.ValueType = 'String'
          Width = 77
          Position.ColIndex = 0
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
        object HighCmdTreeColumn2: TcxTreeListColumn
          Caption.Text = #21517#31281
          DataBinding.ValueType = 'String'
          Width = 169
          Position.ColIndex = 2
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
        object HighCmdTreeColumn3: TcxTreeListColumn
          PropertiesClassName = 'TcxComboBoxProperties'
          Properties.Alignment.Horz = taCenter
          Properties.DropDownListStyle = lsFixedList
          Properties.Items.Strings = (
            'SUB'
            'PPV'
            'CBK'
            'NOR')
          Caption.AlignHorz = taCenter
          Caption.Text = #25351#20196#21029
          DataBinding.ValueType = 'String'
          Width = 58
          Position.ColIndex = 3
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
        object HighCmdTreeColumn4: TcxTreeListColumn
          Caption.Text = #23565#25033#20302#38542#25351#20196
          DataBinding.ValueType = 'String'
          Width = 155
          Position.ColIndex = 0
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
        object HighCmdTreeColumn5: TcxTreeListColumn
          Caption.Text = #22810#32068#25286#35299
          DataBinding.ValueType = 'String'
          Width = 70
          Position.ColIndex = 3
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
        object HighCmdTreeColumn6: TcxTreeListColumn
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taRightJustify
          Caption.AlignHorz = taRightJustify
          Caption.Text = #25351#20196#21029#20778#20808#27402
          DataBinding.ValueType = 'String'
          Width = 80
          Position.ColIndex = 7
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
        object HighCmdTreeColumn7: TcxTreeListColumn
          PropertiesClassName = 'TcxTextEditProperties'
          Properties.Alignment.Horz = taRightJustify
          Caption.AlignHorz = taRightJustify
          Caption.Text = #20659#36865#20778#20808#27402
          DataBinding.ValueType = 'String'
          Width = 80
          Position.ColIndex = 8
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
        object HighCmdTreeColumn8: TcxTreeListColumn
          Caption.Text = #23565#25033'IRD'#25351#20196
          DataBinding.ValueType = 'String'
          Width = 120
          Position.ColIndex = 2
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
        object HighCmdTreeColumn9: TcxTreeListColumn
          Caption.Text = #26032#27231#20302#38542#25351#20196#25976
          DataBinding.ValueType = 'String'
          Width = 100
          Position.ColIndex = 5
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
        object HighCmdTreeColumn10: TcxTreeListColumn
          Caption.Text = #33290#27231#20302#38542#25351#20196#25976
          DataBinding.ValueType = 'String'
          Width = 100
          Position.ColIndex = 6
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
        object HighCmdTreeColumn11: TcxTreeListColumn
          Caption.Text = 'Vod Mopp Id'
          DataBinding.ValueType = 'String'
          Width = 100
          Position.ColIndex = 4
          Position.RowIndex = 0
          Position.BandIndex = 1
        end
      end
      object btnAddHighCmd: TcxButton
        Left = 28
        Top = 64
        Width = 65
        Height = 26
        Hint = #26032#22686#39640#38542#25351#20196
        Caption = #26032#22686
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = btnAddHighCmdClick
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000038307CE0E8E18FF0B8A15FF0A8814FF0985
          12FF007703C70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000048609CE51DA7BFF3ACF69FF39CD67FF32C2
          5BFF017A04C70000000000000000000000000000000000000000000000000000
          000000000000000000000000000005870ACE5CE084FF3ED46EFF3DD36DFF34C5
          5FFF027C05C70000000000000000000000000000000000000000000000000000
          000000000000000000000000000006890CCE64E48AFF41D771FF3FD56FFF37C8
          61FF027D06C70000000000000000000000000000000000000000000000000893
          10D4089211D4079010D4078E0FD4088E10F056E282FF44DA74FF41D872FF39CB
          64FF038209EC037E07CE037C06CE027A05CE017804CE000000000000000021A7
          2DFF60E487FF4BDD79FF4ADC78FF49DC77FF4AE07AFF46DD77FF44DB75FF40D6
          70FF3BCD67FF39CB65FF36C862FF33C45EFF0A8413FF000000000000000022A9
          2FFF76F099FF5EEA8AFF5AE888FF56E684FF53E481FF4EE17DFF47DE78FF44DB
          75FF41D872FF3FD56FFF3DD36DFF39CD67FF0B8715FF000000000000000024AB
          30FF7EF39FFF68EE91FF64ED8EFF60EA8BFF59E886FF54E482FF4EE17DFF46DD
          77FF44DA74FF41D771FF3ED46EFF3ACF69FF0C8916FF000000000000000025AD
          32FF91F7ABFF8DF6A8FF8BF5A6FF89F4A5FF7AF09BFF59E886FF53E481FF4ADF
          7AFF5BE486FF67E58CFF5EE186FF53DB7CFF0F8D1AFF0000000000000000099D
          14CE0A9B14CE099913CE099713CE0B9615EF89F4A5FF60EA8BFF56E684FF45D8
          72FF078D0FEC06890DCE06860CCE05840BCE048209CE00000000000000000000
          0000000000000000000000000000099713CE8BF5A6FF64ED8EFF5AE888FF45D7
          71FF068B0DC70000000000000000000000000000000000000000000000000000
          00000000000000000000000000000A9914CE8DF6A8FF68EE91FF5EEA8AFF46D8
          72FF068E0EC70000000000000000000000000000000000000000000000000000
          00000000000000000000000000000B9A15CE91F7ABFF7FF39FFF76F099FF58DF
          7FFF078F0FC70000000000000000000000000000000000000000000000000000
          00000000000000000000000000000A9C16CE26AD33FF25AB32FF23A930FF21A6
          2EFF079110C70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
        LookAndFeel.Kind = lfFlat
      end
      object btnDeleteHighCmd: TcxButton
        Left = 99
        Top = 64
        Width = 65
        Height = 26
        Hint = #21034#38500#36984#21462#30340#39640#38542#25351#20196
        Caption = #21034#38500
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
        OnClick = btnDeleteHighCmdClick
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          A4140206B0CB030AAEE10000A67B000098070000000000000000000000000000
          00000000A78500009A8B0000A72A000000000000000000000000000000000000
          A98D1844F6FF194DF8FF1031D2FF0102ABD00000B62400000000000000010000
          B1930928D7FF092ED7FF0313B3FF0000AC8F0000000000000000000000000103
          B3D02451F9FF1F52FFFF1D4FFFFF1744E8FF040BB0EC0000B03A0000AC970D2E
          DDFF1142F9FF0D3DF5FF0B3BF0FF041ABCFF0000A55D00000000000000000000
          AE421832DBFF285BFFFF2456FFFF2253FFFF1B4BF1FF050DB1FB0F30DDFF164A
          FEFF1344F9FF1041F6FF0E3EF6FF0A3CF0FF00009FBF00000000000000000000
          BE511F37DDFF3A6FFFFF2C5EFFFF295AFFFF2657FFFF2052FCFF1C4FFFFF194A
          FDFF1646FAFF1445FAFF0F3DF2FF020AB1E70000A84300000000000000000000
          00010000C866121DC8F03D6AFBFF3567FFFF2C5DFFFF2859FFFF2253FFFF1D4E
          FFFF1A4DFFFF123DEDFF0002ACCB0000BA1A0000000000000000000000000000
          0000000000000000CC1C0000B6BA2E4EE7FF3668FFFF2E5EFFFF2859FFFF2254
          FFFF163DEAFF0000A8BE0000AB0B000000000000000000000000000000000000
          000000000000000000000000BF4D253FDFFF3B6DFFFF3464FFFF2E5EFFFF2759
          FFFF1B46EAFF0001ACCF0000A908000000000000000000000000000000000000
          000000000000000000000203C8C14B7CFFFF4170FFFF3B6BFFFF396CFFFF2D5E
          FFFF2558FFFF1336D7FF0000B693000000000000000000000000000000000000
          0000000000000000D92C263CDBFE5080FFFF4575FFFF3662FAFF0D14C3FE3C6D
          FFFF2A5BFFFF2053FDFF0B1DC4FD0000C0470000000000000000000000000000
          0000000000000000CB8B527CFAFF5081FFFF4B7DFFFF0B13C9EA0000BB420E15
          C7E9386AFFFF2456FFFF1A4AF2FF0207B3E30000B50F00000000000000000000
          000000000000131CDDE06A9CFFFF5788FFFF2B46E7FF0000CD5C000000000000
          CD320F1BCBF23065FFFF1F51FFFF1439DDFF0000B19C00000000000000000000
          0000000000000000DE913A52E4FE5782FBFF0101D0C200000000000000000000
          00000000CC471426D6FA265AFFFF0F2EE3FF0103B8CF00000000000000000000
          000000000000000000000000CF3C0000C0A40000CE3000000000000000000000
          0000000000000000C4630001B8BC0000B5490000770200000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
        LookAndFeel.Kind = lfFlat
      end
    end
    object cxLowCmd: TcxTabSheet
      Caption = #20302#38542#25351#20196#23565#29031#34920'   '
      ImageIndex = 4
      object btnAddLowCmd: TcxButton
        Left = 28
        Top = 64
        Width = 65
        Height = 26
        Hint = #26032#22686#20302#38542#25351#20196
        Caption = #26032#22686
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = btnAddLowCmdClick
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000038307CE0E8E18FF0B8A15FF0A8814FF0985
          12FF007703C70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000048609CE51DA7BFF3ACF69FF39CD67FF32C2
          5BFF017A04C70000000000000000000000000000000000000000000000000000
          000000000000000000000000000005870ACE5CE084FF3ED46EFF3DD36DFF34C5
          5FFF027C05C70000000000000000000000000000000000000000000000000000
          000000000000000000000000000006890CCE64E48AFF41D771FF3FD56FFF37C8
          61FF027D06C70000000000000000000000000000000000000000000000000893
          10D4089211D4079010D4078E0FD4088E10F056E282FF44DA74FF41D872FF39CB
          64FF038209EC037E07CE037C06CE027A05CE017804CE000000000000000021A7
          2DFF60E487FF4BDD79FF4ADC78FF49DC77FF4AE07AFF46DD77FF44DB75FF40D6
          70FF3BCD67FF39CB65FF36C862FF33C45EFF0A8413FF000000000000000022A9
          2FFF76F099FF5EEA8AFF5AE888FF56E684FF53E481FF4EE17DFF47DE78FF44DB
          75FF41D872FF3FD56FFF3DD36DFF39CD67FF0B8715FF000000000000000024AB
          30FF7EF39FFF68EE91FF64ED8EFF60EA8BFF59E886FF54E482FF4EE17DFF46DD
          77FF44DA74FF41D771FF3ED46EFF3ACF69FF0C8916FF000000000000000025AD
          32FF91F7ABFF8DF6A8FF8BF5A6FF89F4A5FF7AF09BFF59E886FF53E481FF4ADF
          7AFF5BE486FF67E58CFF5EE186FF53DB7CFF0F8D1AFF0000000000000000099D
          14CE0A9B14CE099913CE099713CE0B9615EF89F4A5FF60EA8BFF56E684FF45D8
          72FF078D0FEC06890DCE06860CCE05840BCE048209CE00000000000000000000
          0000000000000000000000000000099713CE8BF5A6FF64ED8EFF5AE888FF45D7
          71FF068B0DC70000000000000000000000000000000000000000000000000000
          00000000000000000000000000000A9914CE8DF6A8FF68EE91FF5EEA8AFF46D8
          72FF068E0EC70000000000000000000000000000000000000000000000000000
          00000000000000000000000000000B9A15CE91F7ABFF7FF39FFF76F099FF58DF
          7FFF078F0FC70000000000000000000000000000000000000000000000000000
          00000000000000000000000000000A9C16CE26AD33FF25AB32FF23A930FF21A6
          2EFF079110C70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
        LookAndFeel.Kind = lfFlat
      end
      object btnDeleteLowCmd: TcxButton
        Left = 99
        Top = 64
        Width = 65
        Height = 26
        Hint = #21034#38500#36984#21462#30340#20302#38542#25351#20196
        Caption = #21034#38500
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = btnDeleteLowCmdClick
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          A4140206B0CB030AAEE10000A67B000098070000000000000000000000000000
          00000000A78500009A8B0000A72A000000000000000000000000000000000000
          A98D1844F6FF194DF8FF1031D2FF0102ABD00000B62400000000000000010000
          B1930928D7FF092ED7FF0313B3FF0000AC8F0000000000000000000000000103
          B3D02451F9FF1F52FFFF1D4FFFFF1744E8FF040BB0EC0000B03A0000AC970D2E
          DDFF1142F9FF0D3DF5FF0B3BF0FF041ABCFF0000A55D00000000000000000000
          AE421832DBFF285BFFFF2456FFFF2253FFFF1B4BF1FF050DB1FB0F30DDFF164A
          FEFF1344F9FF1041F6FF0E3EF6FF0A3CF0FF00009FBF00000000000000000000
          BE511F37DDFF3A6FFFFF2C5EFFFF295AFFFF2657FFFF2052FCFF1C4FFFFF194A
          FDFF1646FAFF1445FAFF0F3DF2FF020AB1E70000A84300000000000000000000
          00010000C866121DC8F03D6AFBFF3567FFFF2C5DFFFF2859FFFF2253FFFF1D4E
          FFFF1A4DFFFF123DEDFF0002ACCB0000BA1A0000000000000000000000000000
          0000000000000000CC1C0000B6BA2E4EE7FF3668FFFF2E5EFFFF2859FFFF2254
          FFFF163DEAFF0000A8BE0000AB0B000000000000000000000000000000000000
          000000000000000000000000BF4D253FDFFF3B6DFFFF3464FFFF2E5EFFFF2759
          FFFF1B46EAFF0001ACCF0000A908000000000000000000000000000000000000
          000000000000000000000203C8C14B7CFFFF4170FFFF3B6BFFFF396CFFFF2D5E
          FFFF2558FFFF1336D7FF0000B693000000000000000000000000000000000000
          0000000000000000D92C263CDBFE5080FFFF4575FFFF3662FAFF0D14C3FE3C6D
          FFFF2A5BFFFF2053FDFF0B1DC4FD0000C0470000000000000000000000000000
          0000000000000000CB8B527CFAFF5081FFFF4B7DFFFF0B13C9EA0000BB420E15
          C7E9386AFFFF2456FFFF1A4AF2FF0207B3E30000B50F00000000000000000000
          000000000000131CDDE06A9CFFFF5788FFFF2B46E7FF0000CD5C000000000000
          CD320F1BCBF23065FFFF1F51FFFF1439DDFF0000B19C00000000000000000000
          0000000000000000DE913A52E4FE5782FBFF0101D0C200000000000000000000
          00000000CC471426D6FA265AFFFF0F2EE3FF0103B8CF00000000000000000000
          000000000000000000000000CF3C0000C0A40000CE3000000000000000000000
          0000000000000000C4630001B8BC0000B5490000770200000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
        LookAndFeel.Kind = lfFlat
      end
      object LowCmdTree: TcxTreeList
        Left = 12
        Top = 100
        Width = 666
        Height = 235
        Bands = <
          item
          end>
        BufferedPaint = True
        Images = StyleModule.CmdTreeImageList
        LookAndFeel.Kind = lfOffice11
        OptionsBehavior.GoToNextCellOnEnter = True
        OptionsView.CellEndEllipsis = True
        OptionsView.GridLines = tlglBoth
        OptionsView.Indicator = True
        OptionsView.TreeLineStyle = tllsNone
        Styles.Content = StyleModule.cxStyle1
        TabOrder = 2
        OnEdited = SoTreeEdited
        OnFocusedNodeChanged = SoTreeFocusedNodeChanged
        object LowCmdTreeColumn1: TcxTreeListColumn
          Caption.Text = #20302#38542#25351#20196
          DataBinding.ValueType = 'String'
          Width = 112
          Position.ColIndex = 0
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
        object LowCmdTreeColumn2: TcxTreeListColumn
          Caption.Text = #21517#31281
          DataBinding.ValueType = 'String'
          Width = 347
          Position.ColIndex = 2
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
      end
    end
    object cxCmdError: TcxTabSheet
      Caption = #37679#35492#35338#24687#23565#29031#34920'   '
      ImageIndex = 5
      object btnAddError: TcxButton
        Left = 28
        Top = 64
        Width = 65
        Height = 26
        Hint = #26032#22686#37679#35492' '#35338#24687
        Caption = #26032#22686
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = btnAddErrorClick
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000038307CE0E8E18FF0B8A15FF0A8814FF0985
          12FF007703C70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000048609CE51DA7BFF3ACF69FF39CD67FF32C2
          5BFF017A04C70000000000000000000000000000000000000000000000000000
          000000000000000000000000000005870ACE5CE084FF3ED46EFF3DD36DFF34C5
          5FFF027C05C70000000000000000000000000000000000000000000000000000
          000000000000000000000000000006890CCE64E48AFF41D771FF3FD56FFF37C8
          61FF027D06C70000000000000000000000000000000000000000000000000893
          10D4089211D4079010D4078E0FD4088E10F056E282FF44DA74FF41D872FF39CB
          64FF038209EC037E07CE037C06CE027A05CE017804CE000000000000000021A7
          2DFF60E487FF4BDD79FF4ADC78FF49DC77FF4AE07AFF46DD77FF44DB75FF40D6
          70FF3BCD67FF39CB65FF36C862FF33C45EFF0A8413FF000000000000000022A9
          2FFF76F099FF5EEA8AFF5AE888FF56E684FF53E481FF4EE17DFF47DE78FF44DB
          75FF41D872FF3FD56FFF3DD36DFF39CD67FF0B8715FF000000000000000024AB
          30FF7EF39FFF68EE91FF64ED8EFF60EA8BFF59E886FF54E482FF4EE17DFF46DD
          77FF44DA74FF41D771FF3ED46EFF3ACF69FF0C8916FF000000000000000025AD
          32FF91F7ABFF8DF6A8FF8BF5A6FF89F4A5FF7AF09BFF59E886FF53E481FF4ADF
          7AFF5BE486FF67E58CFF5EE186FF53DB7CFF0F8D1AFF0000000000000000099D
          14CE0A9B14CE099913CE099713CE0B9615EF89F4A5FF60EA8BFF56E684FF45D8
          72FF078D0FEC06890DCE06860CCE05840BCE048209CE00000000000000000000
          0000000000000000000000000000099713CE8BF5A6FF64ED8EFF5AE888FF45D7
          71FF068B0DC70000000000000000000000000000000000000000000000000000
          00000000000000000000000000000A9914CE8DF6A8FF68EE91FF5EEA8AFF46D8
          72FF068E0EC70000000000000000000000000000000000000000000000000000
          00000000000000000000000000000B9A15CE91F7ABFF7FF39FFF76F099FF58DF
          7FFF078F0FC70000000000000000000000000000000000000000000000000000
          00000000000000000000000000000A9C16CE26AD33FF25AB32FF23A930FF21A6
          2EFF079110C70000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
        LookAndFeel.Kind = lfFlat
      end
      object btnDeleteError: TcxButton
        Left = 99
        Top = 64
        Width = 65
        Height = 26
        Hint = #21034#38500#36984#21462#30340#37679#35492#35338#24687
        Caption = #21034#38500
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
        OnClick = btnDeleteErrorClick
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          A4140206B0CB030AAEE10000A67B000098070000000000000000000000000000
          00000000A78500009A8B0000A72A000000000000000000000000000000000000
          A98D1844F6FF194DF8FF1031D2FF0102ABD00000B62400000000000000010000
          B1930928D7FF092ED7FF0313B3FF0000AC8F0000000000000000000000000103
          B3D02451F9FF1F52FFFF1D4FFFFF1744E8FF040BB0EC0000B03A0000AC970D2E
          DDFF1142F9FF0D3DF5FF0B3BF0FF041ABCFF0000A55D00000000000000000000
          AE421832DBFF285BFFFF2456FFFF2253FFFF1B4BF1FF050DB1FB0F30DDFF164A
          FEFF1344F9FF1041F6FF0E3EF6FF0A3CF0FF00009FBF00000000000000000000
          BE511F37DDFF3A6FFFFF2C5EFFFF295AFFFF2657FFFF2052FCFF1C4FFFFF194A
          FDFF1646FAFF1445FAFF0F3DF2FF020AB1E70000A84300000000000000000000
          00010000C866121DC8F03D6AFBFF3567FFFF2C5DFFFF2859FFFF2253FFFF1D4E
          FFFF1A4DFFFF123DEDFF0002ACCB0000BA1A0000000000000000000000000000
          0000000000000000CC1C0000B6BA2E4EE7FF3668FFFF2E5EFFFF2859FFFF2254
          FFFF163DEAFF0000A8BE0000AB0B000000000000000000000000000000000000
          000000000000000000000000BF4D253FDFFF3B6DFFFF3464FFFF2E5EFFFF2759
          FFFF1B46EAFF0001ACCF0000A908000000000000000000000000000000000000
          000000000000000000000203C8C14B7CFFFF4170FFFF3B6BFFFF396CFFFF2D5E
          FFFF2558FFFF1336D7FF0000B693000000000000000000000000000000000000
          0000000000000000D92C263CDBFE5080FFFF4575FFFF3662FAFF0D14C3FE3C6D
          FFFF2A5BFFFF2053FDFF0B1DC4FD0000C0470000000000000000000000000000
          0000000000000000CB8B527CFAFF5081FFFF4B7DFFFF0B13C9EA0000BB420E15
          C7E9386AFFFF2456FFFF1A4AF2FF0207B3E30000B50F00000000000000000000
          000000000000131CDDE06A9CFFFF5788FFFF2B46E7FF0000CD5C000000000000
          CD320F1BCBF23065FFFF1F51FFFF1439DDFF0000B19C00000000000000000000
          0000000000000000DE913A52E4FE5782FBFF0101D0C200000000000000000000
          00000000CC471426D6FA265AFFFF0F2EE3FF0103B8CF00000000000000000000
          000000000000000000000000CF3C0000C0A40000CE3000000000000000000000
          0000000000000000C4630001B8BC0000B5490000770200000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
        LookAndFeel.Kind = lfFlat
      end
      object ErrorTree: TcxTreeList
        Left = 16
        Top = 100
        Width = 666
        Height = 235
        Bands = <
          item
          end>
        BufferedPaint = True
        Images = StyleModule.CmdTreeImageList
        LookAndFeel.Kind = lfOffice11
        OptionsBehavior.GoToNextCellOnEnter = True
        OptionsView.CellEndEllipsis = True
        OptionsView.GridLines = tlglBoth
        OptionsView.Indicator = True
        OptionsView.TreeLineStyle = tllsNone
        Styles.Content = StyleModule.cxStyle1
        TabOrder = 2
        OnEdited = SoTreeEdited
        OnFocusedNodeChanged = SoTreeFocusedNodeChanged
        object ErrorTreeColumn1: TcxTreeListColumn
          PropertiesClassName = 'TcxComboBoxProperties'
          Properties.DropDownListStyle = lsFixedList
          Properties.Items.Strings = (
            'CODE'
            'EXT')
          Caption.Text = #39006#21029
          DataBinding.ValueType = 'String'
          Width = 106
          Position.ColIndex = 0
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
        object ErrorTreeColumn2: TcxTreeListColumn
          Caption.Text = #37679#35492#20195#30908
          DataBinding.ValueType = 'String'
          Width = 95
          Position.ColIndex = 2
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
        object ErrorTreeColumn3: TcxTreeListColumn
          Caption.Text = #35338#24687
          DataBinding.ValueType = 'String'
          Width = 370
          Position.ColIndex = 3
          Position.RowIndex = 0
          Position.BandIndex = 0
        end
      end
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 377
    Width = 692
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 1
    object btnConfirm: TcxButton
      Left = 508
      Top = 8
      Width = 82
      Height = 26
      Caption = #20786#23384
      Default = True
      ModalResult = 1
      TabOrder = 0
      LookAndFeel.Kind = lfFlat
      LookAndFeel.NativeStyle = True
    end
    object btnCancel: TcxButton
      Left = 600
      Top = 8
      Width = 82
      Height = 26
      Cancel = True
      Caption = #21462#28040
      ModalResult = 2
      TabOrder = 1
      LookAndFeel.Kind = lfFlat
      LookAndFeel.NativeStyle = True
    end
  end
end
