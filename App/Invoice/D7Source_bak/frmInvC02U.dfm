object frmInvC02: TfrmInvC02
  Left = 425
  Top = 246
  Width = 335
  Height = 221
  Caption = 'frmInvC02'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 327
    Height = 32
    Align = alTop
    Color = clSkyBlue
    TabOrder = 0
    object Label5: TLabel
      Left = 10
      Top = 8
      Width = 119
      Height = 16
      Caption = #30332#31080#20316#24290#19968#35261#34920
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
    Width = 327
    Height = 162
    Align = alClient
    TabOrder = 1
    object Label2: TLabel
      Left = 11
      Top = 86
      Width = 65
      Height = 13
      Caption = #30332#31080#26085#26399#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object Label3: TLabel
      Left = 11
      Top = 50
      Width = 65
      Height = 13
      Caption = #30332#31080#34399#30908#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object Label4: TLabel
      Left = 10
      Top = 15
      Width = 65
      Height = 13
      Alignment = taRightJustify
      Caption = #20316#24290#21407#22240#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object btnExit: TBitBtn
      Left = 200
      Top = 122
      Width = 73
      Height = 26
      Caption = #32080#26463
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 6
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
    object btnQuery: TBitBtn
      Left = 72
      Top = 122
      Width = 73
      Height = 26
      Caption = #26597#35426
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 5
      OnClick = btnQueryClick
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
    inline fraEInvDate: TfraYMD
      Left = 193
      Top = 64
      Width = 128
      Height = 49
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 4
      inherited mseYMD: TMaskEdit
        Left = 8
        Top = 16
        Width = 105
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        ParentFont = False
      end
    end
    inline fraSInvDate: TfraYMD
      Left = 83
      Top = 70
      Width = 118
      Height = 49
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      OnExit = fraSInvDateExit
      inherited mseYMD: TMaskEdit
        Left = 2
        Top = 10
        Width = 103
        Font.Charset = CHINESEBIG5_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        ParentFont = False
      end
    end
    object medSInvID: TMaskEdit
      Left = 87
      Top = 45
      Width = 103
      Height = 21
      EditMask = 'aaaaaaaaaa;1;_'
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      MaxLength = 10
      ParentFont = False
      TabOrder = 1
      Text = '          '
      OnExit = medSInvIDExit
    end
    object medEInvID: TMaskEdit
      Left = 200
      Top = 45
      Width = 105
      Height = 21
      EditMask = 'aaaaaaaaaa;1;_'
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      MaxLength = 10
      ParentFont = False
      TabOrder = 2
      Text = '          '
      OnExit = medEInvIDExit
    end
    object cobObsoleteReason: TComboBox
      Left = 86
      Top = 9
      Width = 105
      Height = 21
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ItemHeight = 13
      ItemIndex = 0
      ParentFont = False
      TabOrder = 0
      Text = #20840#37096
      Items.Strings = (
        #20840#37096
        #29151#26989#20154
        #38750#29151#26989#20154)
    end
    object Button2: TButton
      Left = 8
      Top = 99
      Width = 60
      Height = 25
      Caption = 'PreView'
      TabOrder = 7
      Visible = False
      OnClick = Button2Click
    end
  end
  object Button1: TButton
    Left = 8
    Top = 161
    Width = 60
    Height = 25
    Caption = 'Design'
    TabOrder = 2
    Visible = False
    OnClick = Button1Click
  end
  object frxReport1: TfrxReport
    Version = '3.15'
    DotMatrixReport = False
    IniFile = '\Software\Fast Reports'
    PreviewOptions.Buttons = [pbPrint, pbLoad, pbSave, pbExport, pbZoom, pbFind, pbOutline, pbPageSetup, pbTools, pbEdit, pbNavigator]
    PreviewOptions.Zoom = 1.000000000000000000
    PrintOptions.Printer = #38928#35373
    ReportOptions.CreateDate = 38309.441804097220000000
    ReportOptions.LastChange = 38309.441804097220000000
    ScriptLanguage = 'PascalScript'
    ScriptText.Strings = (
      'begin'
      ''
      'end.')
    OnGetValue = frxReport1GetValue
    Left = 32
    Top = 112
    Datasets = <>
    Variables = <>
    Style = <>
    object Page1: TfrxReportPage
      PaperWidth = 210.000000000000000000
      PaperHeight = 297.000000000000000000
      PaperSize = 9
      LeftMargin = 10.000000000000000000
      RightMargin = 10.000000000000000000
      TopMargin = 10.000000000000000000
      BottomMargin = 10.000000000000000000
    end
  end
  object frxDBDataset1: TfrxDBDataset
    UserName = 'frxDBDataset1'
    CloseDataSource = False
    DataSet = dtmMainJ.cdsObsoleteInvoiceData
    Left = 32
    Top = 56
  end
  object frxReport2: TfrxReport
    Version = '3.15'
    DotMatrixReport = False
    IniFile = '\Software\Fast Reports'
    PreviewOptions.Buttons = [pbPrint, pbLoad, pbSave, pbExport, pbZoom, pbFind, pbOutline, pbPageSetup, pbTools, pbEdit, pbNavigator]
    PreviewOptions.Zoom = 1.000000000000000000
    PrintOptions.Printer = #38928#35373
    ReportOptions.CreateDate = 38313.495332465310000000
    ReportOptions.LastChange = 38314.581440486100000000
    ScriptLanguage = 'PascalScript'
    ScriptText.Strings = (
      'begin'
      ''
      'end.')
    Left = 152
    Top = 96
    Datasets = <
      item
        DataSet = frxDBDataset1
        DataSetName = 'frxDBDataset1'
      end>
    Variables = <
      item
        Name = ' Label'
        Value = Null
      end
      item
        Name = 'lblCondition1'
        Value = Null
      end
      item
        Name = 'lblCondition2'
        Value = Null
      end
      item
        Name = 'qrlOperator'
        Value = Null
      end>
    Style = <>
    object Page1: TfrxReportPage
      PaperWidth = 210.000000000000000000
      PaperHeight = 297.000000000000000000
      PaperSize = 9
      LeftMargin = 10.000000000000000000
      RightMargin = 10.000000000000000000
      TopMargin = 10.000000000000000000
      BottomMargin = 10.000000000000000000
      object PageHeader1: TfrxPageHeader
        Height = 113.385900000000000000
        Top = 18.897650000000000000
        Width = 718.110700000000000000
        object Memo1: TfrxMemoView
          Left = 257.008040000000000000
          Top = 15.118120000000000000
          Width = 173.858380000000000000
          Height = 18.897650000000000000
          Memo.Strings = (
            #30332#31080#20316#24290#19968#35261#34920)
        end
        object Memo2: TfrxMemoView
          Left = 7.559060000000000000
          Top = 3.779530000000001000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          Memo.Strings = (
            '- rptInvC02_1 -')
        end
        object Memo22: TfrxMemoView
          Left = 555.590910000000000000
          Top = 15.118120000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          Memo.Strings = (
            #38913'   '#27425':[Page#]')
        end
      end
      object Header1: TfrxHeader
        Height = 92.047310000000000000
        Top = 192.756030000000000000
        Width = 718.110700000000000000
        object Memo3: TfrxMemoView
          Left = 15.118120000000000000
          Top = 11.338590000000010000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -13
          Font.Name = 'Arial'
          Font.Style = []
          Memo.Strings = (
            #26597#35426#26781#20214': ')
          ParentFont = False
        end
        object Memo4: TfrxMemoView
          Left = 117.165430000000000000
          Top = 11.338590000000010000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -13
          Font.Name = 'Arial'
          Font.Style = []
          Memo.Strings = (
            '[lblCondition1]')
          ParentFont = False
        end
        object Memo5: TfrxMemoView
          Left = 117.165430000000000000
          Top = 34.015770000000010000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -13
          Font.Name = 'Arial'
          Font.Style = []
          Memo.Strings = (
            '[lblCondition2]')
          ParentFont = False
        end
        object Memo6: TfrxMemoView
          Left = 41.574830000000000000
          Top = 68.031540000000010000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            #30332#31080#34399#30908)
        end
        object Memo7: TfrxMemoView
          Left = 143.622140000000000000
          Top = 68.031540000000010000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            #30332#31080#26085#26399)
        end
        object Memo8: TfrxMemoView
          Left = 275.905690000000000000
          Top = 68.031540000000010000
          Width = 68.031540000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          HAlign = haRight
          Memo.Strings = (
            #30332#31080#37329#38989)
        end
        object Memo9: TfrxMemoView
          Left = 355.393940000000000000
          Top = 68.031540000000010000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            #20316#24290#21407#22240)
        end
        object Memo10: TfrxMemoView
          Left = 461.220780000000000000
          Top = 68.031540000000010000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            #23458#25142#20195#30908)
        end
        object Memo11: TfrxMemoView
          Left = 567.268090000000000000
          Top = 68.031540000000010000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            #38283#31435#26041#24335)
        end
        object Line1: TfrxLineView
          Left = 37.795300000000000000
          Top = 89.267780000000010000
          Width = 631.181510000000000000
          Frame.Typ = [ftTop]
        end
      end
      object MasterData1: TfrxMasterData
        Height = 22.677180000000000000
        Top = 306.141930000000000000
        Width = 718.110700000000000000
        DataSet = frxDBDataset1
        DataSetName = 'frxDBDataset1'
        RowCount = 0
        object Memo12: TfrxMemoView
          Left = 41.574830000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            '[frxDBDataset1."InvID"]')
        end
        object Memo13: TfrxMemoView
          Left = 143.622140000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            '[frxDBDataset1."InvDate"]')
        end
        object Memo14: TfrxMemoView
          Left = 249.448980000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          HAlign = haRight
          Memo.Strings = (
            '[frxDBDataset1."InvAmount"]')
        end
        object Memo15: TfrxMemoView
          Left = 355.393940000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            '[frxDBDataset1."ObsoleteReason"]')
        end
        object Memo16: TfrxMemoView
          Left = 461.441250000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            '[frxDBDataset1."CustID"]')
        end
        object Memo17: TfrxMemoView
          Left = 567.268090000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            '[frxDBDataset1."HowToCreate"]')
        end
      end
      object PageFooter1: TfrxPageFooter
        Height = 49.133890000000000000
        Top = 389.291590000000000000
        Width = 718.110700000000000000
        object Memo18: TfrxMemoView
          Left = 37.795300000000000000
          Top = 15.118119999999980000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          Memo.Strings = (
            #21015#21360#26085#26399' : ')
        end
        object Memo19: TfrxMemoView
          Left = 136.063080000000000000
          Top = 15.118119999999980000
          Width = 71.811070000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DisplayFormat.DecimalSeparator = '.'
          DisplayFormat.FormatStr = 'yyyy/mm/dd hh:mm:ss'
          DisplayFormat.Kind = fkDateTime
          Memo.Strings = (
            '[Now]')
        end
        object Memo20: TfrxMemoView
          Left = 423.307360000000000000
          Top = 15.118119999999980000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          Memo.Strings = (
            #21360#34920#20154' : ')
        end
        object Memo21: TfrxMemoView
          Left = 517.795610000000000000
          Top = 15.118119999999980000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          Memo.Strings = (
            '[qrlOperator]')
        end
        object Line2: TfrxLineView
          Left = 37.795300000000000000
          Top = 3.779530000000022000
          Width = 631.181510000000000000
          Frame.Typ = [ftTop]
        end
      end
    end
  end
  object frxDesigner1: TfrxDesigner
    Restrictions = []
    Left = 64
    Top = 56
  end
end
