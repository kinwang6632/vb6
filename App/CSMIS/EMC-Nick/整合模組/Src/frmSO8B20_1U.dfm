object frmSO8B20_1: TfrmSO8B20_1
  Left = 419
  Top = 150
  Width = 420
  Height = 492
  Caption = 'frmSO8B20_1'
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
  object pctMain: TPageControl
    Left = 8
    Top = 16
    Width = 393
    Height = 441
    ActivePage = TabSheet1
    TabOrder = 0
    OnChange = pctMainChange
    object tbsQuery1: TTabSheet
      Caption = #20323#37329#32113#35336#34920
      TabVisible = False
      OnShow = tbsQuery1Show
      object Label2: TLabel
        Left = 36
        Top = 24
        Width = 67
        Height = 16
        AutoSize = False
        Caption = #20844#21496#21029#65306
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label3: TLabel
        Left = 20
        Top = 56
        Width = 79
        Height = 20
        AutoSize = False
        Caption = #27512#23660#24180#26376#65306
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label6: TLabel
        Left = 20
        Top = 80
        Width = 79
        Height = 20
        AutoSize = False
        Caption = #20154'    '#21729#65306
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object stxCom1: TStaticText
        Left = 102
        Top = 23
        Width = 53
        Height = 17
        Caption = 'stxCom1'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
      inline fraChineseYM1: TfraChineseYM
        Left = 102
        Top = 46
        Width = 100
        Height = 35
        TabOrder = 1
      end
      object chbPage: TCheckBox
        Left = 33
        Top = 248
        Width = 99
        Height = 17
        Caption = #20381#35336#31639#20381#25818#20998#38913
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
      end
      object rgpCompute: TRadioGroup
        Left = 19
        Top = 183
        Width = 222
        Height = 57
        Caption = ' '#35336#31639#20381#25818'  '
        Columns = 2
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ItemIndex = 0
        Items.Strings = (
          #20381#20844#21496#35336#31639
          #20381#37096#38272#35336#31639)
        ParentFont = False
        TabOrder = 4
      end
      object btnQueryMaster: TButton
        Left = 10
        Top = 279
        Width = 75
        Height = 25
        Caption = #26597#35426
        TabOrder = 6
        OnClick = btnQueryMasterClick
      end
      object btnMasterClose: TButton
        Left = 172
        Top = 279
        Width = 75
        Height = 25
        Caption = #38626#38283
        TabOrder = 7
        OnClick = btnMasterCloseClick
      end
      object lbxMen: TListBox
        Left = 102
        Top = 105
        Width = 121
        Height = 80
        ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
        ItemHeight = 13
        TabOrder = 3
      end
      object medMan: TMaskEdit
        Left = 102
        Top = 80
        Width = 121
        Height = 21
        ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
        TabOrder = 2
      end
      object btnReset: TButton
        Left = 90
        Top = 279
        Width = 75
        Height = 25
        Caption = #37325#35373
        TabOrder = 8
        OnClick = btnResetClick
      end
    end
    object tbsQuery2: TTabSheet
      Caption = #20323#29518#37329#22577#34920
      ImageIndex = 1
      OnShow = tbsQuery2Show
      object Label1: TLabel
        Left = 23
        Top = 8
        Width = 67
        Height = 16
        AutoSize = False
        Caption = #20844#21496#21029#65306
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label4: TLabel
        Left = 3
        Top = 29
        Width = 79
        Height = 20
        Alignment = taRightJustify
        AutoSize = False
        Caption = #27512#23660#24180#26376#65306
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label5: TLabel
        Left = 14
        Top = 57
        Width = 63
        Height = 13
        Alignment = taRightJustify
        Caption = 'SO '#21729' '#24037':'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label7: TLabel
        Left = 15
        Top = 157
        Width = 63
        Height = 13
        Alignment = taRightJustify
        Caption = #37559' '#21806' '#40670':'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label8: TLabel
        Left = 140
        Top = 154
        Width = 63
        Height = 13
        Alignment = taRightJustify
        Caption = #38364#20225#21729#24037':'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label9: TLabel
        Left = 259
        Top = 152
        Width = 63
        Height = 13
        Alignment = taRightJustify
        Caption = #25490#24207#26041#24335':'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label14: TLabel
        Left = 139
        Top = 55
        Width = 98
        Height = 13
        Alignment = taRightJustify
        Caption = #25910#36027#38971#36947#38917#30446#65306
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object stxCom2: TStaticText
        Left = 89
        Top = 7
        Width = 53
        Height = 17
        Caption = 'stxCom2'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 0
      end
      inline fraChineseYM2: TfraChineseYM
        Left = 84
        Top = 19
        Width = 100
        Height = 35
        TabOrder = 1
        inherited mseYM: TMaskEdit
          Left = 16
        end
      end
      object btnQueryDetail: TButton
        Left = 103
        Top = 360
        Width = 75
        Height = 23
        Caption = #26597#35426
        TabOrder = 15
        OnClick = btnQueryDetailClick
      end
      object btnDetailClose: TButton
        Left = 201
        Top = 358
        Width = 75
        Height = 25
        Caption = #38626#38283
        TabOrder = 16
        OnClick = btnDetailCloseClick
      end
      object btnWorker: TBitBtn
        Left = 107
        Top = 129
        Width = 21
        Height = 22
        Caption = '...'
        TabOrder = 4
        OnClick = btnWorkerClick
      end
      object lbxWorkerCode: TListBox
        Left = 12
        Top = 73
        Width = 93
        Height = 77
        ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
        ItemHeight = 13
        TabOrder = 3
      end
      object btnSales: TBitBtn
        Left = 108
        Top = 228
        Width = 21
        Height = 22
        Caption = '...'
        TabOrder = 8
        OnClick = btnSalesClick
      end
      object lbxSales: TListBox
        Left = 13
        Top = 172
        Width = 93
        Height = 77
        ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
        ItemHeight = 13
        TabOrder = 7
      end
      object lbxOtherCompCode: TListBox
        Left = 139
        Top = 170
        Width = 93
        Height = 77
        ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
        ItemHeight = 13
        TabOrder = 9
      end
      object btnOtherCompCode: TBitBtn
        Left = 234
        Top = 226
        Width = 21
        Height = 22
        Caption = '...'
        TabOrder = 10
        OnClick = btnOtherCompCodeClick
      end
      object btnOrder: TBitBtn
        Left = 352
        Top = 225
        Width = 22
        Height = 22
        Caption = '...'
        TabOrder = 12
        OnClick = btnOrderClick
      end
      object lbxOrderList: TListBox
        Left = 258
        Top = 170
        Width = 93
        Height = 77
        ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
        ItemHeight = 13
        MultiSelect = True
        TabOrder = 11
      end
      object rgpPayTypeRpt: TRadioGroup
        Left = 63
        Top = 255
        Width = 241
        Height = 41
        Caption = '  '#19968#27425'/'#20998#26399'  '
        Columns = 2
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ItemIndex = 1
        Items.Strings = (
          #19968#27425#32102#20184
          #20998#26399#32102#20184)
        ParentFont = False
        TabOrder = 13
      end
      object rgpClassRpt: TRadioGroup
        Left = 63
        Top = 302
        Width = 242
        Height = 51
        Caption = '  '#22577#34920#31278#39006'  '
        Columns = 3
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ItemIndex = 0
        Items.Strings = (
          #20323#37329#26126#32048
          #29518#37329#26126#32048
          #32113#35336#34920)
        ParentFont = False
        TabOrder = 14
        OnClick = rgpClassRptClick
        OnEnter = rgpClassRptEnter
      end
      object lbxChargeItem1: TListBox
        Left = 139
        Top = 71
        Width = 210
        Height = 77
        ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
        ItemHeight = 13
        TabOrder = 5
      end
      object btnItem1: TBitBtn
        Left = 354
        Top = 127
        Width = 21
        Height = 22
        Caption = '...'
        TabOrder = 6
        OnClick = btnItem1Click
      end
      object chkWorkerComm1: TCheckBox
        Left = 184
        Top = 32
        Width = 145
        Height = 17
        Caption = #21253#21547'BOX'#24037#31243#20154#21729#29518#37329
        TabOrder = 2
      end
    end
    object TabSheet1: TTabSheet
      Caption = #20323#29518#37329'Excel'
      ImageIndex = 2
      OnShow = TabSheet1Show
      object Label10: TLabel
        Left = 65
        Top = 43
        Width = 196
        Height = 13
        Caption = #25910#36027#38971#36947#38917#30446'('#20323#37329#25165#38656#40670#36984')'#65306
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label11: TLabel
        Left = 65
        Top = 11
        Width = 70
        Height = 13
        Caption = #27512#23660#24180#26376#65306
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label16: TLabel
        Left = 64
        Top = 347
        Width = 97
        Height = 13
        AutoSize = False
        Caption = 'Excel'#20786#23384#36335#24465':'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      inline fraChineseYM3: TfraChineseYM
        Left = 132
        Top = 2
        Width = 100
        Height = 34
        TabOrder = 0
        inherited mseYM: TMaskEdit
          Left = 4
          Top = 10
        end
      end
      object lbxChargeItem: TListBox
        Left = 61
        Top = 61
        Width = 234
        Height = 80
        ImeName = #20013#25991' ('#32321#39636') - '#26032#27880#38899
        ItemHeight = 13
        TabOrder = 1
      end
      object btnItem: TBitBtn
        Left = 298
        Top = 119
        Width = 21
        Height = 22
        Caption = '...'
        TabOrder = 2
        OnClick = btnItemClick
      end
      object btnQueryExcel: TButton
        Left = 60
        Top = 377
        Width = 75
        Height = 25
        Caption = #26597#35426
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 9
        OnClick = btnQueryExcelClick
      end
      object btnExcelExit: TButton
        Left = 234
        Top = 377
        Width = 75
        Height = 25
        Caption = #38626#38283
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 11
        OnClick = btnExcelExitClick
      end
      object btnReset1: TButton
        Left = 147
        Top = 377
        Width = 75
        Height = 25
        Caption = #37325#35373
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 10
        OnClick = btnReset1Click
      end
      object GroupBox1: TGroupBox
        Left = 59
        Top = 145
        Width = 129
        Height = 97
        Caption = '   '#20323#37329'   '
        TabOrder = 3
        object chbChannelComm_C: TCheckBox
          Left = 14
          Top = 72
          Width = 99
          Height = 17
          Caption = #25910#36027#21729#26126#32048
          Checked = True
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          State = cbChecked
          TabOrder = 2
        end
        object chbChannelComm_I: TCheckBox
          Left = 14
          Top = 48
          Width = 99
          Height = 17
          Caption = #20171#32057#20154#26126#32048
          Checked = True
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          State = cbChecked
          TabOrder = 1
        end
        object chbChannelComm_A: TCheckBox
          Left = 14
          Top = 24
          Width = 107
          Height = 17
          Caption = #20323#37329#25152#26377#26126#32048
          Checked = True
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          State = cbChecked
          TabOrder = 0
        end
      end
      object GroupBox2: TGroupBox
        Left = 195
        Top = 145
        Width = 129
        Height = 97
        Caption = '   '#29518#37329'   '
        TabOrder = 4
        object chbBoxComm_W: TCheckBox
          Left = 16
          Top = 71
          Width = 105
          Height = 17
          Caption = #24037#31243#20154#21729#26126#32048
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          TabOrder = 2
        end
        object chbBoxComm_I: TCheckBox
          Left = 16
          Top = 47
          Width = 105
          Height = 17
          Caption = #20171#32057#20154#26126#32048
          Checked = True
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          State = cbChecked
          TabOrder = 1
        end
        object chbBoxComm_A: TCheckBox
          Left = 16
          Top = 23
          Width = 105
          Height = 17
          Caption = #29518#37329#25152#26377#26126#32048
          Checked = True
          Font.Charset = ANSI_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = #32048#26126#39636
          Font.Style = []
          ParentFont = False
          State = cbChecked
          TabOrder = 0
        end
      end
      object rgpType: TRadioGroup
        Left = 61
        Top = 247
        Width = 262
        Height = 41
        Caption = '  '#19968#27425'/'#20998#26399'  '
        Columns = 2
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ItemIndex = 1
        Items.Strings = (
          #19968#27425#32102#20184
          #20998#26399#32102#20184)
        ParentFont = False
        TabOrder = 5
      end
      object chbStatisticExcel: TCheckBox
        Left = 62
        Top = 317
        Width = 123
        Height = 17
        Caption = #36681#20986#32113#35336#34920
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 7
      end
      object chkWorkerComm2: TCheckBox
        Left = 62
        Top = 293
        Width = 145
        Height = 17
        Caption = #21253#21547'BOX'#24037#31243#20154#21729#29518#37329
        TabOrder = 6
      end
      object edtPath1: TEdit
        Left = 160
        Top = 342
        Width = 193
        Height = 21
        ReadOnly = True
        TabOrder = 8
      end
    end
    object TabSheet2: TTabSheet
      Caption = #38928#20184#22577#34920
      ImageIndex = 3
      TabVisible = False
      object Label12: TLabel
        Left = 21
        Top = 31
        Width = 70
        Height = 13
        Caption = #22522#28310#24180#26376#65306
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      inline fraChineseYM4: TfraChineseYM
        Left = 88
        Top = 15
        Width = 100
        Height = 34
        TabOrder = 0
        inherited mseYM: TMaskEdit
          Left = 4
          Top = 10
        end
      end
      object rgpRptType: TRadioGroup
        Left = 21
        Top = 78
        Width = 172
        Height = 83
        Caption = '  '#22577#34920#31278#39006'  '
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ItemIndex = 0
        Items.Strings = (
          #24180#24230#38928#25910#32317#34920
          #30070#26376#26126#32048)
        ParentFont = False
        TabOrder = 1
      end
      object chbReport: TCheckBox
        Left = 22
        Top = 184
        Width = 107
        Height = 17
        Caption = #22577#34920
        Checked = True
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
        State = cbChecked
        TabOrder = 2
      end
      object chbExcel: TCheckBox
        Left = 126
        Top = 184
        Width = 99
        Height = 17
        Caption = 'Excel'
        Checked = True
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
        State = cbChecked
        TabOrder = 3
      end
      object Button1: TButton
        Left = 12
        Top = 233
        Width = 75
        Height = 25
        Caption = #26597#35426
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        OnClick = Button1Click
      end
      object Button2: TButton
        Left = 99
        Top = 233
        Width = 75
        Height = 25
        Caption = #37325#35373
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 5
        OnClick = Button2Click
      end
      object Button3: TButton
        Left = 186
        Top = 233
        Width = 75
        Height = 25
        Caption = #38626#38283
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 6
        OnClick = Button3Click
      end
    end
    object TabSheet3: TTabSheet
      Caption = #25892#25552#34920
      ImageIndex = 4
      OnShow = TabSheet3Show
      object Label13: TLabel
        Left = 25
        Top = 23
        Width = 70
        Height = 13
        Caption = #35336#31639#24180#24230#65306
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
      end
      object Label15: TLabel
        Left = 24
        Top = 125
        Width = 97
        Height = 13
        AutoSize = False
        Caption = 'Excel'#20786#23384#36335#24465':'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Button4: TButton
        Left = 28
        Top = 161
        Width = 75
        Height = 25
        Caption = #26597#35426
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        OnClick = Button4Click
      end
      object Button5: TButton
        Left = 139
        Top = 161
        Width = 75
        Height = 25
        Caption = #37325#35373
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 4
        OnClick = Button5Click
      end
      object Button6: TButton
        Left = 250
        Top = 161
        Width = 75
        Height = 25
        Caption = #38626#38283
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ParentFont = False
        TabOrder = 5
        OnClick = Button6Click
      end
      object medYear: TMaskEdit
        Left = 96
        Top = 16
        Width = 41
        Height = 21
        EditMask = '9999;1;_'
        MaxLength = 4
        TabOrder = 0
        Text = '    '
      end
      object rgpType2: TRadioGroup
        Left = 25
        Top = 54
        Width = 172
        Height = 51
        Caption = '  '#22577#34920#31278#39006'  '
        Columns = 3
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #32048#26126#39636
        Font.Style = []
        ItemIndex = 0
        Items.Strings = (
          #22577#34920
          'Excel')
        ParentFont = False
        TabOrder = 1
        OnClick = rgpClassRptClick
        OnEnter = rgpClassRptEnter
      end
      object edtPath2: TEdit
        Left = 120
        Top = 120
        Width = 241
        Height = 21
        ReadOnly = True
        TabOrder = 2
      end
    end
  end
end
