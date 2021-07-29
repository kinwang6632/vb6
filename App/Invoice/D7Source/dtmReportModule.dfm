object dtmReport: TdtmReport
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 475
  Top = 310
  Height = 314
  Width = 366
  object ADOMaster: TADOQuery
    CacheSize = 1000
    Connection = dtmMain.InvConnection
    Parameters = <>
    Prepared = True
    Left = 36
    Top = 104
  end
  object ADODetail: TADOQuery
    CacheSize = 100
    Connection = dtmMain.InvConnection
    Parameters = <>
    Prepared = True
    Left = 36
    Top = 160
  end
  object frxC03_1: TfrxDBDataset
    UserName = 'frxC03_1'
    CloseDataSource = True
    FieldAliases.Strings = (
      'BILLID='#25910#36027#21934#34399
      'CHARGEDATE='#25910#36027#26085#26399
      'DESCRIPTION='#25910#36027#38917#30446
      'TOTALAMOUNT='#32317#37329#38989
      'NOTES='#20633#35387
      'CUSTID='#23458#32232
      'CUSTNAME='#21517#31281
      'LOGTIME=Log'#26178#38291)
    DataSet = ADOMaster
    Left = 225
    Top = 24
  end
  object frxA01_1: TfrxDBDataset
    UserName = 'frxA01_1'
    CloseDataSource = True
    FieldAliases.Strings = (
      'SEQ='#27969#27700#34399
      'CUSTID='#23458#32232
      'TITLE='#25260#38957
      'CHARGEDATE='#25910#36027#26085
      'DESCRIPTION='#31237#21029#21517#31281
      'TAXRATE='#31237#29575
      'SALEAMOUNT='#37559#21806#38989
      'TAXAMOUNT='#31237#38989
      'INVAMOUNT='#30332#31080#38989
      'SHOULDBEASSIGNED='#26159#21542#38920#35201#38283#31435
      'TAXTYPE='#31237#21029
      'SHOULDBEASSIGNED2='#26159#21542#38920#35201#38283#31435'('#26126#32048')'
      'TOTALAMOUNT='#30332#31080#38989'('#26126#32048')'
      'TAXTYPE2='#31237#21029'('#26126#32048')')
    DataSet = ADOMaster
    Left = 112
    Top = 24
  end
  object frxA05_1: TfrxDBDataset
    UserName = 'frxA05_1'
    CloseDataSource = False
    FieldAliases.Strings = (
      'CHECKNO='#27298#26597#30908
      'INVID='#30332#31080#34399#30908
      'BUSINESSID='#32113#19968#32232#34399
      'INVDATE='#30332#31080#26085#26399
      'CUSTID='#23458#25142#20195#30908
      'CUSTSNAME='#23458#25142#31777#31281
      'INVTITLE='#30332#31080#25260#38957
      'ZIPCODE='#37109#36958#21312#34399
      'MAILADDR='#37109#23492#22320#22336
      'INVADDR='#30332#31080#22320#22336
      'INVFORMAT='#30332#31080#26684#24335
      'TAXTYPE='#31237#21029
      'SALEAMOUNT='#37559#21806#38989
      'TAXAMOUNT='#31237#38989
      'INVAMOUNT='#30332#31080#32317#37329#38989
      'DESCRIPTION1='#21697#21517'('#19968')'
      'STARTDATE1='#26377#25928#26085#36215'('#19968')'
      'ENDDATE1='#26377#25928#26085#36804'('#19968')'
      'QUANTITY1='#25976#37327'('#19968')'
      'UNITPRICE1='#21934#20729'('#19968')'
      'TOTALAMOUNT1='#32317#37329#38989'('#19968')'
      'DESCRIPTION2='#21697#21517'('#20108')'
      'STARTDATE2='#26377#25928#26085#36215'('#20108')'
      'ENDDATE2='#26377#25928#26085#36804'('#20108')'
      'QUANTITY2='#25976#37327'('#20108')'
      'UNITPRICE2='#21934#20729'('#20108')'
      'TOTALAMOUNT2='#32317#37329#38989'('#20108')'
      'DESCRIPTION3='#21697#21517'('#19977')'
      'STARTDATE3='#26377#25928#26085#36215'('#19977')'
      'ENDDATE3='#26377#25928#26085#36804'('#19977')'
      'QUANTITY3='#25976#37327'('#19977')'
      'UNITPRICE3='#21934#20729'('#19977')'
      'TOTALAMOUNT3='#32317#37329#38989'('#19977')'
      'DESCRIPTION4='#21697#21517'('#22235')'
      'STARTDATE4='#26377#25928#26085#36215'('#22235')'
      'ENDDATE4='#26377#25928#26085#36804'('#22235')'
      'QUANTITY4='#25976#37327'('#22235')'
      'UNITPRICE4='#21934#20729'('#22235')'
      'TOTALAMOUNT4='#32317#37329#38989'('#22235')'
      'DESCRIPTION5='#21697#21517'('#20116')'
      'STARTDATE5='#26377#25928#26085#36215'('#20116')'
      'ENDDATE5='#26377#25928#26085#36804'('#20116')'
      'QUANTITY5='#25976#37327'('#20116')'
      'UNITPRICE5='#21934#20729'('#20116')'
      'TOTALAMOUNT5='#32317#37329#38989'('#20116')'
      'MEMO1='#20633#35387
      'MAININVID='#20027#30332#31080#34399#30908
      'MAINSALEAMOUNT='#20027#30332#31080#32317#37559#21806#38989
      'MAINTAXAMOUNT='#20027#30332#31080#32317#31237#38989
      'MAININVAMOUNT='#20027#30332#31080#32317#37329#38989
      'APPLYFORMAT='#30003#22577#26684#24335
      'Creditcard4d='#33256#27331#20449#29992#21345#26411#22235#30908
      'QRCodeChineseInvDate=QRCODE'#27665#22283#24180#30332#31080#26085#26399
      'QRCodeInvDate=QRCode'#30332#31080#26085#26399#26178#38291
      'QRCodeSellerBUSINESSID=QRCode'#36067#26041#32113#19968#32232#34399
      'CODE39Data=CDDE39'#36039#26009#20839#23481
      'QRCodePngPath1=QRCode1'#22294#27284#36335#24465
      'QRCodePngPath2=QRCode2'#22294#27284#36335#24465
      'QRCodeInvID=QRCode'#30332#31080#34399#30908
      'QRCodeRanNum=QRCode'#38568#27231#30908
      'QRCodeTotalAmount=QRCode'#32317#35336#37329#38989
      'QRCodeBuyerBUSINESSID=QRCode'#36023#26041#32113#19968#32232#34399
      'QRCodeInvID=QRCode'#30332#31080#34399#30908
      'CanModify='#26159#21542#21487#20462#25913
      'QRCodeEFormat=QRCode'#26684#24335
      'CompSName='#20844#21496#31777#31281
      'COMPID='#20844#21496#21029
      'COMPNAME='#20844#21496#20840#21517
      'COMPBUSINESSID='#20844#21496#32113#32232
      'COMPADDR='#20844#21496#29151#26989#22320#22336
      'COMPTEL='#20844#21496#38651#35441
      'PRIZETYPE='#20013#29518#35352#37636)
    DataSet = cdsTempory
    Left = 170
    Top = 24
  end
  object cdsTempory: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 112
    Top = 103
  end
  object frxC05_1: TfrxDBDataset
    UserName = 'frxC05_1'
    CloseDataSource = True
    FieldAliases.Strings = (
      'PAGE='#38913#27425
      'RECNO='#24207#34399
      'INVID='#30332#31080#34399#30908
      'INVDATE='#30332#31080#26085#26399
      'BUSINESSID='#30332#31080#32113#32232
      'INVFORMAT='#30332#31080#26684#24335
      'TAXTYPE='#31237#21029
      'SALEAMOUNT='#37559#21806#38989
      'TAXAMOUNT='#31237#38989
      'INVAMOUNT='#30332#31080#37329#38989
      'ISOBSOLETE='#20570#24290#21542)
    DataSet = cdsTempory
    Left = 274
    Top = 24
  end
  object frxMainReport: TfrxReport
    Version = '3.15'
    DotMatrixReport = False
    IniFile = '\Software\Fast Reports'
    PreviewOptions.Buttons = [pbPrint, pbLoad, pbSave, pbExport, pbZoom, pbFind, pbOutline, pbPageSetup, pbTools, pbEdit, pbNavigator]
    PreviewOptions.Zoom = 1.000000000000000000
    PrintOptions.Printer = 'Default'
    ReportOptions.CreateDate = 38845.688991435200000000
    ReportOptions.LastChange = 42618.580198009260000000
    ScriptLanguage = 'PascalScript'
    ScriptText.Strings = (
      'procedure MasterData1OnBeforePrint(Sender: TfrxComponent);'
      'var'
      '  aText: String;'
      'begin'
      ''
      '  if ( <frxA07_2."'#30332#31080#31278#39006'">  = '#39'1'#39' ) then'
      '    aText := '#39#38651#23376#39
      '  else if ( <frxA07_2."'#30332#31080#31278#39006'">  = '#39'2'#39' ) then'
      '    aText := '#39#25163#20108#39
      '  else if ( <frxA07_2."'#30332#31080#31278#39006'">  = '#39'3'#39' ) then'
      '    aText := '#39#25163#19977#39';'
      '  Set( '#39#30332#31080#31278#39006'('#35498#26126')'#39', '#39#39#39#39' + aText + '#39#39#39#39' );'
      '  {}'
      '  {'
      '  if ( <frxA07_2."'#31237#21029'">  = '#39'1'#39' ) then'
      '    aText := '#39#25033#31237#39
      '  else if ( <frxA07_2."'#31237#21029'">  = '#39'2'#39' ) then'
      '    aText := '#39#38646#31237#39
      '  else if ( <frxA07_2."'#31237#21029'">  = '#39'3'#39' ) then'
      '    aText := '#39#20813#31237#39';'
      '  Set( '#39#31237#21029'('#35498#26126')'#39', '#39#39#39#39' + aText + '#39#39#39#39' ); }'
      'end;'
      ''
      'begin'
      ''
      'end.')
    Left = 44
    Top = 26
    Datasets = <
      item
        DataSet = frxA07_2
        DataSetName = 'frxA07_2'
      end>
    Variables = <
      item
        Name = ' '#26597#35426#26781#20214
        Value = Null
      end
      item
        Name = #20844#21496#21029
        Value = Null
      end
      item
        Name = ' '#20854#23427
        Value = Null
      end
      item
        Name = #30332#31080#31278#39006'('#35498#26126')'
        Value = Null
      end
      item
        Name = #31237#21029'('#35498#26126')'
        Value = Null
      end>
    Style = <>
    object Page1: TfrxReportPage
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #27161#26999#39636
      Font.Style = []
      Orientation = poLandscape
      PaperWidth = 297.000000000000000000
      PaperHeight = 210.000000000000000000
      PaperSize = 9
      LeftMargin = 10.000000000000000000
      RightMargin = 10.000000000000000000
      TopMargin = 10.000000000000000000
      BottomMargin = 10.000000000000000000
      object MasterData1: TfrxMasterData
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlack
        Font.Height = -15
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        Height = 574.488560000000000000
        ParentFont = False
        Top = 18.897650000000000000
        Width = 1046.929810000000000000
        OnBeforePrint = 'MasterData1OnBeforePrint'
        DataSet = frxA07_2
        DataSetName = 'frxA07_2'
        RowCount = 0
        StartNewPage = True
        object Shape1: TfrxShapeView
          Left = 15.118120000000000000
          Top = 94.488250000000000000
          Width = 480.000310000000000000
          Height = 105.826840000000000000
          Frame.Width = 2.000000000000000000
        end
        object Line3: TfrxLineView
          Left = 41.574830000000000000
          Top = 94.488250000000000000
          Height = 105.826840000000000000
          Frame.Typ = [ftLeft]
          Frame.Width = 2.000000000000000000
        end
        object Memo32: TfrxMemoView
          Left = 18.897650000000000000
          Top = 102.047310000000000000
          Width = 22.677180000000000000
          Height = 83.149660000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #21407
            #38283
            #31435
            #37559
            #36008)
          ParentFont = False
        end
        object Line4: TfrxLineView
          Left = 68.031540000000000000
          Top = 94.488250000000000000
          Height = 105.826840000000000000
          Frame.Typ = [ftLeft]
          Frame.Width = 2.000000000000000000
        end
        object Line5: TfrxLineView
          Left = 158.740260000000000000
          Top = 94.488250000000000000
          Height = 105.826840000000000000
          Frame.Typ = [ftLeft]
          Frame.Width = 2.000000000000000000
        end
        object Line6: TfrxLineView
          Left = 68.031540000000000000
          Top = 120.944960000000000000
          Width = 427.086890000000000000
          Frame.Typ = [ftTop]
          Frame.Width = 2.000000000000000000
        end
        object Line7: TfrxLineView
          Left = 68.031540000000000000
          Top = 158.740260000000000000
          Width = 427.086890000000000000
          Frame.Typ = [ftTop]
          Frame.Width = 2.000000000000000000
        end
        object Memo2: TfrxMemoView
          Left = 45.354360000000000000
          Top = 109.606370000000000000
          Width = 22.677180000000000000
          Height = 68.031540000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #30332
            #31080
            #21934
            #20301)
          ParentFont = False
        end
        object Memo3: TfrxMemoView
          Left = 79.370130000000000000
          Top = 98.267780000000000000
          Width = 68.031540000000000000
          Height = 15.118120000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #21517#12288#12288#31281)
          ParentFont = False
        end
        object Memo4: TfrxMemoView
          Left = 79.370130000000000000
          Top = 124.724490000000000000
          Width = 68.031540000000000000
          Height = 34.015770000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #29151#21033#20107#26989
            #32113#19968#32232#34399)
          ParentFont = False
        end
        object Memo5: TfrxMemoView
          Left = 79.370130000000000000
          Top = 162.519790000000000000
          Width = 68.031540000000000000
          Height = 34.015770000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #29151'  '#26989'  '#25152
            #22312'  '#22320'  '#22336)
          ParentFont = False
        end
        object Memo10: TfrxMemoView
          Left = 162.519790000000000000
          Top = 98.267780000000000000
          Width = 328.819110000000000000
          Height = 22.677180000000000000
          DataField = #20844#21496#20840#21517
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            '[frxA07_2."'#20844#21496#20840#21517'"]')
          ParentFont = False
        end
        object Memo1: TfrxMemoView
          Left = 162.519790000000000000
          Top = 132.283550000000000000
          Width = 249.448980000000000000
          Height = 18.897650000000000000
          DataField = #20844#21496#32113#32232
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          Memo.Strings = (
            '[frxA07_2."'#20844#21496#32113#32232'"]')
        end
        object Memo6: TfrxMemoView
          Left = 162.519790000000000000
          Top = 162.519790000000000000
          Width = 328.819110000000000000
          Height = 34.015770000000000000
          DataField = #20844#21496#29151#26989#22320#22336
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          Memo.Strings = (
            '[frxA07_2."'#20844#21496#29151#26989#22320#22336'"]')
        end
        object Memo7: TfrxMemoView
          Left = 510.236550000000000000
          Top = 102.047310000000000000
          Width = 321.260050000000000000
          Height = 18.897650000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -19
          Font.Name = #26032#32048#26126#39636
          Font.Style = [fsBold]
          Memo.Strings = (
            #29151#26989#20154#37559#36008#36864#22238#36914#36008#36864#20986#25110#25240#35731#35657#26126#21934)
          ParentFont = False
        end
        object Memo8: TfrxMemoView
          Left = 888.189550000000000000
          Top = 102.047310000000000000
          Width = 79.370130000000000000
          Height = 18.897650000000000000
          DataField = #21015#21360#20154#21729
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          Memo.Strings = (
            '[frxA07_2."'#21015#21360#20154#21729'"]')
        end
        object Memo9: TfrxMemoView
          Left = 1024.252630000000000000
          Top = 98.267780000000000000
          Width = 18.897650000000000000
          Height = 272.126160000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -11
          Font.Name = #26032#32048#26126#39636
          Font.Style = [fsBold]
          Memo.Strings = (
            #31532
            #19968
            #32879
            #65306
            #20132
            #20184
            #21407
            #37559
            #36008
            #20154
            #20316
            #28858
            #37559
            #38917
            #31237
            #38989
            #20043
            #25187
            #27454
            #24977
            #35657)
          ParentFont = False
        end
        object Memo11: TfrxMemoView
          Left = 510.236550000000000000
          Top = 170.078850000000000000
          Width = 64.252010000000000000
          Height = 15.118120000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #20013#33775#27665#22283)
          ParentFont = False
        end
        object Memo12: TfrxMemoView
          Left = 578.268090000000000000
          Top = 170.078850000000000000
          Width = 30.236240000000000000
          Height = 15.118120000000000000
          Memo.Strings = (
            '[IntToStr(StrToInt(<frxA07_2."'#25240#35731#21934#26085#26399'('#24180')">) - 1911)]')
        end
        object Memo13: TfrxMemoView
          Left = 604.724800000000000000
          Top = 170.078850000000000000
          Width = 18.897650000000000000
          Height = 18.897650000000000000
          Memo.Strings = (
            #24180)
        end
        object Memo14: TfrxMemoView
          Left = 627.401980000000000000
          Top = 170.078850000000000000
          Width = 22.677180000000000000
          Height = 18.897650000000000000
          DataField = #25240#35731#21934#26085#26399'('#26376')'
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          Memo.Strings = (
            '[frxA07_2."'#25240#35731#21934#26085#26399'('#26376')"]')
        end
        object Memo15: TfrxMemoView
          Left = 650.079160000000000000
          Top = 170.078850000000000000
          Width = 18.897650000000000000
          Height = 18.897650000000000000
          Memo.Strings = (
            #26376)
        end
        object Memo16: TfrxMemoView
          Left = 668.976810000000000000
          Top = 170.078850000000000000
          Width = 22.677180000000000000
          Height = 18.897650000000000000
          DataField = #25240#35731#21934#26085#26399'('#26085')'
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          Memo.Strings = (
            '[frxA07_2."'#25240#35731#21934#26085#26399'('#26085')"]')
        end
        object Memo17: TfrxMemoView
          Left = 691.653990000000000000
          Top = 170.078850000000000000
          Width = 18.897650000000000000
          Height = 18.897650000000000000
          Memo.Strings = (
            #26085)
        end
        object Memo18: TfrxMemoView
          Left = 737.008350000000000000
          Top = 143.622140000000000000
          Width = 49.133890000000000000
          Height = 15.118120000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #24207#34399#65306)
          ParentFont = False
        end
        object Memo19: TfrxMemoView
          Left = 737.008350000000000000
          Top = 170.078850000000000000
          Width = 49.133890000000000000
          Height = 15.118120000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #23458#32232#65306)
          ParentFont = False
        end
        object Memo20: TfrxMemoView
          Left = 782.362710000000000000
          Top = 170.078850000000000000
          Width = 79.370130000000000000
          Height = 18.897650000000000000
          DataField = #23458#32232
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          Memo.Strings = (
            '[frxA07_2."'#23458#32232'"]')
        end
        object Shape2: TfrxShapeView
          Left = 15.118120000000000000
          Top = 226.771800000000000000
          Width = 994.016390000000000000
          Height = 204.094620000000000000
          Frame.Width = 2.000000000000000000
        end
        object Line1: TfrxLineView
          Left = 15.118120000000000000
          Top = 253.228510000000000000
          Width = 823.937540000000000000
          Frame.Typ = [ftTop]
          Frame.Width = 2.000000000000000000
        end
        object Line2: TfrxLineView
          Left = 839.055660000000000000
          Top = 226.771800000000000000
          Height = 204.094620000000000000
          Frame.Typ = [ftLeft]
          Frame.Width = 2.000000000000000000
        end
        object Line8: TfrxLineView
          Left = 15.118120000000000000
          Top = 332.598640000000000000
          Width = 994.016390000000000000
          Frame.Typ = [ftTop]
          Frame.Width = 2.000000000000000000
        end
        object Line9: TfrxLineView
          Left = 15.118120000000000000
          Top = 298.582870000000000000
          Width = 994.016390000000000000
          Frame.Typ = [ftTop]
          Frame.Width = 2.000000000000000000
        end
        object Line10: TfrxLineView
          Left = 895.748610000000000000
          Top = 298.582870000000000000
          Height = 132.283550000000000000
          Frame.Typ = [ftLeft]
          Frame.Width = 2.000000000000000000
        end
        object Line11: TfrxLineView
          Left = 952.441560000000000000
          Top = 298.582870000000000000
          Height = 132.283550000000000000
          Frame.Typ = [ftLeft]
          Frame.Width = 2.000000000000000000
        end
        object Memo21: TfrxMemoView
          Left = 888.189550000000000000
          Top = 245.669450000000000000
          Width = 83.149660000000000000
          Height = 22.677180000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -16
          Font.Name = #26032#32048#26126#39636
          Font.Style = [fsBold]
          Memo.Strings = (
            #25171' ( V ) '#34389)
          ParentFont = False
        end
        object Memo22: TfrxMemoView
          Left = 888.189550000000000000
          Top = 268.346630000000000000
          Width = 83.149660000000000000
          Height = 22.677180000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -16
          Font.Name = #26032#32048#26126#39636
          Font.Style = [fsBold]
          Memo.Strings = (
            #20633'         '#35387)
          ParentFont = False
        end
        object Memo23: TfrxMemoView
          Left = 854.173780000000000000
          Top = 306.141930000000000000
          Width = 37.795300000000000000
          Height = 15.118120000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #25033#31237)
          ParentFont = False
        end
        object Memo24: TfrxMemoView
          Left = 903.307670000000000000
          Top = 306.141930000000000000
          Width = 49.133890000000000000
          Height = 15.118120000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #38646#31237#29575)
          ParentFont = False
        end
        object Memo25: TfrxMemoView
          Left = 963.780150000000000000
          Top = 306.141930000000000000
          Width = 37.795300000000000000
          Height = 15.118120000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #20813#31237)
          ParentFont = False
        end
        object Memo26: TfrxMemoView
          Left = 865.512370000000000000
          Top = 336.378170000000000000
          Width = 18.897650000000000000
          Height = 15.118120000000000000
          Memo.Strings = (
            '[IIF(<frxA07_2."'#31237#21029'">=1,'#39'V'#39','#39#39')]')
        end
        object Memo27: TfrxMemoView
          Left = 910.866730000000000000
          Top = 336.378170000000000000
          Width = 34.015770000000000000
          Height = 15.118120000000000000
          Memo.Strings = (
            '[IIF(<frxA07_2."'#31237#21029'">=2,'#39'V'#39','#39#39')]')
        end
        object Memo28: TfrxMemoView
          Left = 971.339210000000000000
          Top = 336.378170000000000000
          Width = 18.897650000000000000
          Height = 15.118120000000000000
          Memo.Strings = (
            '[IIF(<frxA07_2."'#31237#21029'">=3,'#39'V'#39','#39#39')]')
        end
        object Line12: TfrxLineView
          Left = 56.692950000000000000
          Top = 253.228510000000000000
          Height = 132.283550000000000000
          Frame.Typ = [ftLeft]
          Frame.Width = 2.000000000000000000
        end
        object Line13: TfrxLineView
          Left = 90.708720000000000000
          Top = 253.228510000000000000
          Height = 132.283550000000000000
          Frame.Typ = [ftLeft]
          Frame.Width = 2.000000000000000000
        end
        object Line14: TfrxLineView
          Left = 128.504020000000000000
          Top = 253.228510000000000000
          Height = 132.283550000000000000
          Frame.Typ = [ftLeft]
          Frame.Width = 2.000000000000000000
        end
        object Line15: TfrxLineView
          Left = 162.519790000000000000
          Top = 253.228510000000000000
          Height = 132.283550000000000000
          Frame.Typ = [ftLeft]
          Frame.Width = 2.000000000000000000
        end
        object Line16: TfrxLineView
          Left = 279.685220000000000000
          Top = 226.771800000000000000
          Height = 158.740260000000000000
          Frame.Typ = [ftLeft]
          Frame.Width = 2.000000000000000000
        end
        object Memo29: TfrxMemoView
          Left = 79.370130000000000000
          Top = 230.551330000000000000
          Width = 151.181200000000000000
          Height = 18.897650000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -19
          Font.Name = #26032#32048#26126#39636
          Font.Style = [fsBold]
          Memo.Strings = (
            #38283#12288#31435#12288#30332#12288#31080)
          ParentFont = False
        end
        object Memo30: TfrxMemoView
          Left = 22.677180000000000000
          Top = 260.787570000000000000
          Width = 18.897650000000000000
          Height = 30.236240000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #32879
            #24335)
          ParentFont = False
        end
        object Memo31: TfrxMemoView
          Left = 64.252010000000000000
          Top = 272.126160000000000000
          Width = 18.897650000000000000
          Height = 22.677180000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #24180)
          ParentFont = False
        end
        object Memo33: TfrxMemoView
          Left = 98.267780000000000000
          Top = 272.126160000000000000
          Width = 18.897650000000000000
          Height = 22.677180000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #26376)
          ParentFont = False
        end
        object Memo34: TfrxMemoView
          Left = 136.063080000000000000
          Top = 272.126160000000000000
          Width = 18.897650000000000000
          Height = 22.677180000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #26085)
          ParentFont = False
        end
        object Memo35: TfrxMemoView
          Left = 60.472480000000000000
          Top = 306.141930000000000000
          Width = 30.236240000000000000
          Height = 15.118120000000000000
          Memo.Strings = (
            '[IntToStr(StrToInt(<frxA07_2."'#25240#35731#21934#26085#26399'('#24180')">) - 1911)]')
        end
        object Memo36: TfrxMemoView
          Left = 98.267780000000000000
          Top = 306.141930000000000000
          Width = 22.677180000000000000
          Height = 18.897650000000000000
          DataField = #25240#35731#21934#26085#26399'('#26376')'
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          Memo.Strings = (
            '[frxA07_2."'#25240#35731#21934#26085#26399'('#26376')"]')
        end
        object Memo37: TfrxMemoView
          Left = 136.063080000000000000
          Top = 306.141930000000000000
          Width = 22.677180000000000000
          Height = 18.897650000000000000
          DataField = #25240#35731#21934#26085#26399'('#26085')'
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          Memo.Strings = (
            '[frxA07_2."'#25240#35731#21934#26085#26399'('#26085')"]')
        end
        object Line17: TfrxLineView
          Left = 207.874150000000000000
          Top = 253.228510000000000000
          Height = 132.283550000000000000
          Frame.Typ = [ftLeft]
          Frame.Width = 2.000000000000000000
        end
        object Memo38: TfrxMemoView
          Left = 170.078850000000000000
          Top = 272.126160000000000000
          Width = 34.015770000000000000
          Height = 22.677180000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #23383#36556)
          ParentFont = False
        end
        object Memo39: TfrxMemoView
          Left = 173.858380000000000000
          Top = 306.141930000000000000
          Width = 26.456710000000000000
          Height = 18.897650000000000000
          DataField = #38283#31435#30332#31080#23383#36556
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          Memo.Strings = (
            '[frxA07_2."'#38283#31435#30332#31080#23383#36556'"]')
        end
        object Memo40: TfrxMemoView
          Left = 215.433210000000000000
          Top = 272.126160000000000000
          Width = 60.472480000000000000
          Height = 22.677180000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #34399'       '#30908)
          ParentFont = False
        end
        object Memo41: TfrxMemoView
          Left = 215.433210000000000000
          Top = 306.141930000000000000
          Width = 56.692950000000000000
          Height = 18.897650000000000000
          DataField = #38283#31435#30332#31080#34399#30908
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          Memo.Strings = (
            '[frxA07_2."'#38283#31435#30332#31080#34399#30908'"]')
        end
        object Memo42: TfrxMemoView
          Left = 18.897650000000000000
          Top = 306.141930000000000000
          Width = 34.015770000000000000
          Height = 22.677180000000000000
          Memo.Strings = (
            '['#30332#31080#31278#39006'('#35498#26126')]')
        end
        object Memo43: TfrxMemoView
          Left = 430.866420000000000000
          Top = 230.551330000000000000
          Width = 275.905690000000000000
          Height = 18.897650000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -19
          Font.Name = #26032#32048#26126#39636
          Font.Style = [fsBold]
          Memo.Strings = (
            #36864#12288#36008#12288#25110#12288#25240#12288#35731#12288#20839#12288#23481)
          ParentFont = False
        end
        object Line18: TfrxLineView
          Left = 472.441250000000000000
          Top = 253.228510000000000000
          Height = 177.637910000000000000
          Frame.Typ = [ftLeft]
          Frame.Width = 2.000000000000000000
        end
        object Memo44: TfrxMemoView
          Left = 332.598640000000000000
          Top = 272.126160000000000000
          Width = 94.488250000000000000
          Height = 22.677180000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #21697#12288#12288'       '#21517)
          ParentFont = False
        end
        object Memo45: TfrxMemoView
          Left = 287.244280000000000000
          Top = 306.141930000000000000
          Width = 177.637910000000000000
          Height = 18.897650000000000000
          DataField = #30332#31080#21697#21517
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          Memo.Strings = (
            '[frxA07_2."'#30332#31080#21697#21517'"]')
        end
        object Line19: TfrxLineView
          Left = 521.575140000000000000
          Top = 253.228510000000000000
          Height = 177.637910000000000000
          Frame.Typ = [ftLeft]
          Frame.Width = 2.000000000000000000
        end
        object Line20: TfrxLineView
          Left = 574.488560000000000000
          Top = 253.228510000000000000
          Height = 177.637910000000000000
          Frame.Typ = [ftLeft]
          Frame.Width = 2.000000000000000000
        end
        object Memo46: TfrxMemoView
          Left = 480.000310000000000000
          Top = 272.126160000000000000
          Width = 34.015770000000000000
          Height = 22.677180000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #25976#37327)
          ParentFont = False
        end
        object Memo47: TfrxMemoView
          Left = 532.913730000000000000
          Top = 272.126160000000000000
          Width = 34.015770000000000000
          Height = 22.677180000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #21934#20729)
          ParentFont = False
        end
        object Line21: TfrxLineView
          Left = 574.488560000000000000
          Top = 275.905690000000000000
          Width = 264.567100000000000000
          Frame.Typ = [ftTop]
          Frame.Width = 2.000000000000000000
        end
        object Memo48: TfrxMemoView
          Left = 638.740570000000000000
          Top = 257.008040000000000000
          Width = 136.063080000000000000
          Height = 22.677180000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #36864#12288#20986#12288#25110#12288#25240#12288#35731)
          ParentFont = False
        end
        object Memo49: TfrxMemoView
          Left = 578.268090000000000000
          Top = 279.685220000000000000
          Width = 177.637910000000000000
          Height = 22.677180000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #37329#38989#65288#19981#21547#31237#20043#36914#36008#38989#65289)
          ParentFont = False
        end
        object Line22: TfrxLineView
          Left = 763.465060000000000000
          Top = 275.905690000000000000
          Height = 154.960730000000000000
          Frame.Typ = [ftLeft]
          Frame.Width = 2.000000000000000000
        end
        object Memo50: TfrxMemoView
          Left = 771.024120000000000000
          Top = 279.685220000000000000
          Width = 56.692950000000000000
          Height = 22.677180000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #29151' '#26989' '#31237)
          ParentFont = False
        end
        object Memo51: TfrxMemoView
          Left = 578.268090000000000000
          Top = 306.141930000000000000
          Width = 177.637910000000000000
          Height = 18.897650000000000000
          DataField = #37559#21806#38989
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          HAlign = haRight
          Memo.Strings = (
            '[frxA07_2."'#37559#21806#38989'"]')
        end
        object Memo52: TfrxMemoView
          Left = 767.244590000000000000
          Top = 306.141930000000000000
          Width = 68.031540000000000000
          Height = 18.897650000000000000
          DataField = #31237#38989
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          HAlign = haRight
          Memo.Strings = (
            '[frxA07_2."'#31237#38989'"]')
        end
        object Line23: TfrxLineView
          Left = 15.118120000000000000
          Top = 359.055350000000000000
          Width = 994.016390000000000000
          Frame.Typ = [ftTop]
          Frame.Width = 2.000000000000000000
        end
        object Line24: TfrxLineView
          Left = 15.118120000000000000
          Top = 385.512060000000000000
          Width = 994.016390000000000000
          Frame.Typ = [ftTop]
          Frame.Width = 2.000000000000000000
        end
        object Memo53: TfrxMemoView
          Left = 287.244280000000000000
          Top = 396.850650000000000000
          Width = 139.842610000000000000
          Height = 18.897650000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -19
          Font.Name = #26032#32048#26126#39636
          Font.Style = [fsBold]
          Memo.Strings = (
            #21512#12288#12288#12288#12288#35336)
          ParentFont = False
        end
        object Memo54: TfrxMemoView
          Left = 865.512370000000000000
          Top = 400.630180000000000000
          Width = 18.897650000000000000
          Height = 15.118120000000000000
          Memo.Strings = (
            '[IIF(<frxA07_2."'#31237#21029'">=1,'#39'V'#39','#39#39')]')
        end
        object Memo55: TfrxMemoView
          Left = 910.866730000000000000
          Top = 400.630180000000000000
          Width = 34.015770000000000000
          Height = 15.118120000000000000
          Memo.Strings = (
            '[IIF(<frxA07_2."'#31237#21029'">=2,'#39'V'#39','#39#39')]')
        end
        object Memo56: TfrxMemoView
          Left = 971.339210000000000000
          Top = 400.630180000000000000
          Width = 18.897650000000000000
          Height = 15.118120000000000000
          Memo.Strings = (
            '[IIF(<frxA07_2."'#31237#21029'">=3,'#39'V'#39','#39#39')]')
        end
        object Memo57: TfrxMemoView
          Left = 767.244590000000000000
          Top = 396.850650000000000000
          Width = 68.031540000000000000
          Height = 18.897650000000000000
          DataField = #31237#38989
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          HAlign = haRight
          Memo.Strings = (
            '[frxA07_2."'#31237#38989'"]')
        end
        object Memo58: TfrxMemoView
          Left = 574.488560000000000000
          Top = 396.850650000000000000
          Width = 177.637910000000000000
          Height = 18.897650000000000000
          DataField = #37559#21806#38989
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          HAlign = haRight
          Memo.Strings = (
            '[frxA07_2."'#37559#21806#38989'"]')
        end
        object Memo59: TfrxMemoView
          Left = 15.118120000000000000
          Top = 438.425480000000000000
          Width = 438.425480000000000000
          Height = 15.118120000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #26412#35657#26126#21934#25152#21015#37559#36008#36864#22238#36914#36008#36864#20986#25110#25240#35731#65292#30906#23526#23660#23526#65292#29305#27492#35657#26126#12290)
          ParentFont = False
        end
        object Memo60: TfrxMemoView
          Left = 15.118120000000000000
          Top = 461.102660000000000000
          Width = 98.267780000000000000
          Height = 37.795300000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #21407#36914#36008#29151#26989#20154
            '('#25110#21407#36023#21463#20154')')
          ParentFont = False
        end
        object Memo61: TfrxMemoView
          Left = 143.622140000000000000
          Top = 472.441250000000000000
          Width = 49.133890000000000000
          Height = 15.118120000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #21517#31281#65306)
          ParentFont = False
        end
        object Memo62: TfrxMemoView
          Left = 495.118430000000000000
          Top = 472.441250000000000000
          Width = 192.756030000000000000
          Height = 15.118120000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            '('#33995#31456')'#12288#12288#12288#12288#12288#22320#12288#22336#65306)
          ParentFont = False
        end
        object Memo63: TfrxMemoView
          Left = 15.118120000000000000
          Top = 506.457020000000000000
          Width = 147.401670000000000000
          Height = 15.118120000000000000
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #26032#32048#26126#39636
          Font.Style = []
          Memo.Strings = (
            #29151#26989#20107#26989#32113#19968#32232#34399#65306)
          ParentFont = False
        end
        object Memo64: TfrxMemoView
          Left = 170.078850000000000000
          Top = 506.457020000000000000
          Width = 207.874150000000000000
          Height = 18.897650000000000000
          DataField = #32113#32232
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          Memo.Strings = (
            '[frxA07_2."'#32113#32232'"]')
        end
        object Memo65: TfrxMemoView
          Left = 15.118120000000000000
          Top = 536.693260000000000000
          Width = 241.889920000000000000
          Height = 18.897650000000000000
          DataField = #30332#31080#25260#38957
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          Memo.Strings = (
            '[frxA07_2."'#30332#31080#25260#38957'"]')
        end
        object Memo66: TfrxMemoView
          Left = 279.685220000000000000
          Top = 536.693260000000000000
          Width = 355.275820000000000000
          Height = 18.897650000000000000
          DataField = #23458#25142#37109#23492#22320#22336
          DataSet = frxA07_2
          DataSetName = 'frxA07_2'
          Memo.Strings = (
            '[frxA07_2."'#23458#25142#37109#23492#22320#22336'"]')
        end
      end
    end
  end
  object frxA01_2: TfrxDBDataset
    UserName = 'frxA01_2'
    CloseDataSource = False
    FieldAliases.Strings = (
      'BILLID='#25910#36027#21934#34399
      'BILLIDITEMNO='#25910#36027#21934#34399#38917#27425
      'TAXTYPE='#31237#21029' '
      'CHARGEDATE='#25910#36027#26085#26399
      'ITEMID='#21697#21517#20195#30908
      'DESCRIPTION='#21697#21517' '
      'QUANTITY='#25976#37327
      'UNITPRICE='#21934#20729
      'TAXRATE='#29151#26989#31237#29575' '
      'TAXAMOUNT='#31237#38989' '
      'TOTALAMOUNT='#32317#37329#38989' '
      'STARTDATE='#26377#25928#36215#22987#26085
      'ENDDATE='#26377#25928#25130#27490#26085' '
      'CHARGEEN='#25910#36027#20154#21729
      'CUSTID='#23458#25142#32232#34399' '
      'CUSTNAME='#23458#25142#21517#31281
      'LOGTIME=Log'#26178#38291' '
      'UPTEN='#25805#20570#20154#21729)
    DataSet = ADOMaster
    Left = 171
    Top = 104
  end
  object frxC04_1: TfrxDBDataset
    UserName = 'frxC04_1'
    CloseDataSource = False
    FieldAliases.Strings = (
      'ISOBSOLETE='#20316#24290#21542
      'INVDATE='#30332#31080#26085#26399
      'AMOUNT_HOWTOCREATE_1='#38928#38283'('#37329#38989')'
      'AMOUNT_HOWTOCREATE_2='#24460#38283'('#37329#38989')'
      'AMOUNT_HOWTOCREATE_3='#29694#22580#38283#31435'('#37329#38989')'
      'AMOUNT_HOWTOCREATE_4='#19968#33324#38283#31435'('#37329#38989')'
      'COUNT_HOWTOCREATE_1='#38928#38283'('#31558#25976')'
      'COUNT_HOWTOCREATE_2='#24460#38283'('#31558#25976')'
      'COUNT_HOWTOCREATE_3='#29694#22580#38283#31435'('#31558#25976')'
      'COUNT_HOWTOCREATE_4='#19968#33324#38283#31435'('#31558#25976')'
      'TOTALAMOUNT='#32317#37329#38989
      'TOTALCOUNT='#32317#31558#25976)
    DataSet = cdsTempory
    Left = 225
    Top = 104
  end
  object frxA07_1: TfrxDBDataset
    UserName = 'frxA07_1'
    CloseDataSource = False
    FieldAliases.Strings = (
      'COMPID='#20844#21496#21029
      'COMPSNAME='#20844#21496#21517#31281
      'YEARMONTH='#30003#22577#24180#26376'('#19968')'
      'YEARMONTH2='#30003#22577#24180#26376'('#20108')'
      'PAPERNO='#35657#26126#21934#34399#30908
      'PAPERDATE='#35657#26126#21934#26085#26399
      'INVID='#30332#31080#34399#30908
      'INVDATE='#30332#31080#26085#26399
      'CUSTID='#23458#32232
      'CUSTNAME='#31777#31281
      'BUSINESSID='#32113#32232
      'INVFORMAT='#30332#31080#31278#39006
      'SALEAMOUNT='#37559#21806#38989
      'TAXAMOUNT='#31237#38989
      'INVAMOUNT='#30332#31080#37329#38989
      'TAXTYPE='#31237#21029
      'INVOICEKIND='#30332#31080#22411#24907
      'UPLOADFLAG='#19978#20659#35387#35352
      'ISOBSOLETE='#20316#24290)
    Left = 282
    Top = 72
  end
  object frxB04_1: TfrxDBDataset
    UserName = 'frxB04_1'
    CloseDataSource = False
    FieldAliases.Strings = (
      'CUSTID='#23458#32232
      'CUSTNAME='#23458#25142#21517#31281
      'CITEMCODE='#25910#36027#38917#30446#20195#34399
      'CITEMNAME='#25910#36027#38917#30446#21517#31281
      'REALDATE='#23526#25910#26085#26399
      'REALAMT='#23526#25910#37329#38989
      'SALEAMT='#37559#21806#38989
      'TAXAMT='#31237#38989
      'GUINO='#30332#31080#34399#30908
      'ERRMSG='#30064#24120#21407#22240)
    DataSet = cdsTempory
    Left = 168
    Top = 168
  end
  object frxC01_1: TfrxDBDataset
    UserName = 'frxC01_1'
    CloseDataSource = False
    FieldAliases.Strings = (
      'DESCRIPTION='#25910#36027#38917#30446
      'QUANTITY='#25976#37327
      'SALEAMOUNT='#37559#21806#38989
      'TAXAMOUNT='#31237#38989
      'TOTALAMOUNT='#30332#31080#37329#38989)
    Left = 225
    Top = 169
  end
  object frxC01_2: TfrxDBDataset
    UserName = 'frxC01_2'
    CloseDataSource = False
    FieldAliases.Strings = (
      'INVDATE='#30332#31080#26085#26399
      'INVID='#30332#31080#34399#30908
      'DESCRIPTION='#25910#36027#38917#30446
      'BILLID='#25910#36027#21934#34399
      'SALEAMOUNT='#37559#21806#38989
      'TAXAMOUNT='#31237#38989
      'TOTALAMOUNT='#30332#31080#37329#38989
      'CUSTID='#23458#32232
      'CUSTSNAME='#23458#25142#21517#31281
      'INVTITLE='#30332#31080#25260#38957
      'BUSINESSID='#32113#32232
      'CHARGEDATE='#25910#36027#26085#26399
      'UPLOADFLAG='#19978#20659#35387#35352
      'INVOICEKIND='#30332#31080#31278#39006
      'ISGIVE='#25424#36104#35387#35352
      'GIVEUNITDESC='#21463#36104#21934#20301)
    Left = 277
    Top = 169
  end
  object frxB06_1: TfrxDBDataset
    UserName = 'frxB06_1'
    CloseDataSource = True
    FieldAliases.Strings = (
      'YEAR='#23565#29518#30332#31080#20351#29992#24180
      'MONTH='#23565#29518#30332#31080#20351#29992#26376
      'YEARMONTH='#23565#29518#30332#31080#20351#29992#24180#26376
      'EXECUTEDATE='#22519#34892#23565#29518#26085
      'UPTTIME='#30064#21205#26178#38291
      'UPTEN='#30064#21205#20154#21729
      'UNUSUALPRIZE1='#29305#21029#29518#34399#30908'1'
      'UNUSUALPRIZE2='#29305#21029#29518#34399#30908'2'
      'UNUSUALPRIZE3='#29305#21029#29518#34399#30908'3'
      'UNUSUALPRIZE4='#29305#21029#29518#34399#30908'4'
      'UNUSUALPRIZE5='#29305#21029#29518#34399#30908'5'
      'SPECIALPRIZE1='#29305#29518#34399#30908'1'
      'SPECIALPRIZE2='#29305#29518#34399#30908'2'
      'SPECIALPRIZE3='#29305#29518#34399#30908'3'
      'SPECIALPRIZE4='#29305#29518#34399#30908'4'
      'SPECIALPRIZE5='#29305#29518#34399#30908'5'
      'FIRSTPRIZE1='#38957#29518#34399#30908'1'
      'FIRSTPRIZE2='#38957#29518#34399#30908'2'
      'FIRSTPRIZE3='#38957#29518#34399#30908'3'
      'FIRSTPRIZE4='#38957#29518#34399#30908'4'
      'FIRSTPRIZE5='#38957#29518#34399#30908'5'
      'EXTRAPRIZE1='#21152#38283#34399#30908'1'
      'EXTRAPRIZE2='#21152#38283#34399#30908'2'
      'EXTRAPRIZE3='#21152#38283#34399#30908'3'
      'EXTRAPRIZE4='#21152#38283#34399#30908'4'
      'EXTRAPRIZE5='#21152#38283#34399#30908'5'
      'COMPID='#20844#21496#21029
      'COMPNAME='#20844#21496#21517#31281
      'STARTMONTH='#36215#22987#26376#20221
      'ENDMONTH='#32066#27490#26376#20221)
    DataSet = cdsTempory
    Left = 120
    Top = 168
  end
  object frxDBDataset1: TfrxDBDataset
    UserName = 'frxA05_1'
    CloseDataSource = False
    FieldAliases.Strings = (
      'CHECKNO='#27298#26597#30908
      'INVID='#30332#31080#34399#30908
      'BUSINESSID='#32113#19968#32232#34399
      'INVDATE='#30332#31080#26085#26399
      'CUSTID='#23458#25142#20195#30908
      'CUSTSNAME='#23458#25142#31777#31281
      'INVTITLE='#30332#31080#25260#38957
      'ZIPCODE='#37109#36958#21312#34399
      'MAILADDR='#37109#23492#22320#22336
      'INVADDR='#30332#31080#22320#22336
      'INVFORMAT='#30332#31080#26684#24335
      'TAXTYPE='#31237#21029
      'SALEAMOUNT='#37559#21806#38989
      'TAXAMOUNT='#31237#38989
      'INVAMOUNT='#30332#31080#32317#37329#38989
      'DESCRIPTION1='#21697#21517'('#19968')'
      'STARTDATE1='#26377#25928#26085#36215'('#19968')'
      'ENDDATE1='#26377#25928#26085#36804'('#19968')'
      'QUANTITY1='#25976#37327'('#19968')'
      'UNITPRICE1='#21934#20729'('#19968')'
      'TOTALAMOUNT1='#32317#37329#38989'('#19968')'
      'DESCRIPTION2='#21697#21517'('#20108')'
      'STARTDATE2='#26377#25928#26085#36215'('#20108')'
      'ENDDATE2='#26377#25928#26085#36804'('#20108')'
      'QUANTITY2='#25976#37327'('#20108')'
      'UNITPRICE2='#21934#20729'('#20108')'
      'TOTALAMOUNT2='#32317#37329#38989'('#20108')'
      'DESCRIPTION3='#21697#21517'('#19977')'
      'STARTDATE3='#26377#25928#26085#36215'('#19977')'
      'ENDDATE3='#26377#25928#26085#36804'('#19977')'
      'QUANTITY3='#25976#37327'('#19977')'
      'UNITPRICE3='#21934#20729'('#19977')'
      'TOTALAMOUNT3='#32317#37329#38989'('#19977')'
      'DESCRIPTION4='#21697#21517'('#22235')'
      'STARTDATE4='#26377#25928#26085#36215'('#22235')'
      'ENDDATE4='#26377#25928#26085#36804'('#22235')'
      'QUANTITY4='#25976#37327'('#22235')'
      'UNITPRICE4='#21934#20729'('#22235')'
      'TOTALAMOUNT4='#32317#37329#38989'('#22235')'
      'DESCRIPTION5='#21697#21517'('#20116')'
      'STARTDATE5='#26377#25928#26085#36215'('#20116')'
      'ENDDATE5='#26377#25928#26085#36804'('#20116')'
      'QUANTITY5='#25976#37327'('#20116')'
      'UNITPRICE5='#21934#20729'('#20116')'
      'TOTALAMOUNT5='#32317#37329#38989'('#20116')'
      'MEMO1='#20633#35387
      'MAININVID='#20027#30332#31080#34399#30908
      'MAINSALEAMOUNT='#20027#30332#31080#32317#37559#21806#38989
      'MAINTAXAMOUNT='#20027#30332#31080#32317#31237#38989
      'MAININVAMOUNT='#20027#30332#31080#32317#37329#38989)
    DataSet = cdsTempory
    Left = 152
    Top = 224
  end
  object frxA07_2: TfrxDBDataset
    UserName = 'frxA07_2'
    CloseDataSource = False
    FieldAliases.Strings = (
      'COMPID='#20844#21496#21029
      'COMPSNAME='#20844#21496#21517#31281
      'YEARMONTH='#30003#22577#24180#26376'('#19968')'
      'YEARMONTH2='#30003#22577#24180#26376'('#20108')'
      'PAPERNO='#35657#26126#21934#34399#30908
      'PAPERDATE='#35657#26126#21934#26085#26399
      'INVID='#30332#31080#34399#30908
      'INVDATE='#30332#31080#26085#26399
      'CUSTID='#23458#32232
      'CUSTNAME='#23458#25142#31777#31281
      'BUSINESSID='#32113#32232
      'INVFORMAT='#30332#31080#31278#39006
      'SALEAMOUNT='#37559#21806#38989
      'TAXAMOUNT='#31237#38989
      'INVAMOUNT='#30332#31080#37329#38989
      'TAXTYPE='#31237#21029
      'INVOICEKIND='#30332#31080#22411#24907
      'UPLOADFLAG='#19978#20659#35387#35352
      'ISOBSOLETE='#20316#24290
      'ITEMNAME='#30332#31080#21697#21517
      'COMPNAME='#20844#21496#20840#21517
      'BUSINESSID2='#20844#21496#32113#32232
      'COMPADDR='#20844#21496#29151#26989#22320#22336
      'INVWORD='#38283#31435#30332#31080#23383#36556
      'INVNUM='#38283#31435#30332#31080#34399#30908
      'PAPERYEAR='#25240#35731#21934#26085#26399'('#24180')'
      'PAPERMONTH='#25240#35731#21934#26085#26399'('#26376')'
      'PAPERDAY='#25240#35731#21934#26085#26399'('#26085')'
      'INVYEAR='#38283#31435#30332#31080'('#24180')'
      'INVMONTH='#38283#31435#30332#31080'('#26376')'
      'INVDAY='#38283#31435#30332#31080'('#26085')'
      'PRINTNAME='#21015#21360#20154#21729
      'INVTITLE='#30332#31080#25260#38957
      'MAILADDR='#23458#25142#37109#23492#22320#22336)
    Left = 284
    Top = 122
  end
end
