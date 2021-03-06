object dtmReport: TdtmReport
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 427
  Top = 257
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
      'MAININVAMOUNT='#20027#30332#31080#32317#37329#38989)
    DataSet = cdsTempory
    Left = 168
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
    ReportOptions.CreateDate = 40365.595511423600000000
    ReportOptions.LastChange = 40374.493460648150000000
    ScriptLanguage = 'PascalScript'
    ScriptText.Strings = (
      'begin'
      ''
      'end.')
    Left = 36
    Top = 24
    Datasets = <
      item
        DataSet = frxB06_1
        DataSetName = 'frxB06_1'
      end>
    Variables = <
      item
        Name = ' '#33258#35330
        Value = Null
      end
      item
        Name = #20013#29518#34399#30908#21934#25260#38957
        Value = ''
      end>
    Style = <>
    object Page1: TfrxReportPage
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -21
      Font.Name = 'Arial'
      Font.Style = []
      PaperWidth = 210.000000000000000000
      PaperHeight = 297.000000000000000000
      PaperSize = 9
      LeftMargin = 10.000000000000000000
      RightMargin = 10.000000000000000000
      TopMargin = 10.000000000000000000
      BottomMargin = 10.000000000000000000
      object ReportTitle1: TfrxReportTitle
        Height = 117.165430000000000000
        Top = 18.897650000000000000
        Width = 718.110700000000000000
        object Memo14: TfrxMemoView
          Left = 238.110390000000000000
          Top = 26.456710000000000000
          Width = 34.015770000000010000
          Height = 26.456710000000000000
          Memo.Strings = (
            #24180)
        end
        object Memo17: TfrxMemoView
          Left = 404.409710000000100000
          Top = 26.456710000000000000
          Width = 34.015770000000010000
          Height = 26.456710000000000000
          Memo.Strings = (
            #26376)
        end
        object Memo13: TfrxMemoView
          Left = 147.401670000000000000
          Top = 26.456710000000000000
          Width = 79.370130000000000000
          Height = 34.015770000000010000
          DataField = #23565#29518#30332#31080#20351#29992#24180
          DataSet = frxB06_1
          DataSetName = 'frxB06_1'
          HAlign = haRight
          Memo.Strings = (
            '[frxB06_1."'#23565#29518#30332#31080#20351#29992#24180'"]')
        end
        object Memo15: TfrxMemoView
          Left = 275.905690000000100000
          Top = 26.456710000000000000
          Width = 45.354360000000000000
          Height = 30.236240000000000000
          DataField = #36215#22987#26376#20221
          DataSet = frxB06_1
          DataSetName = 'frxB06_1'
          Memo.Strings = (
            '[frxB06_1."'#36215#22987#26376#20221'"]')
        end
        object Memo18: TfrxMemoView
          Left = 370.393940000000100000
          Top = 26.456710000000000000
          Width = 37.795300000000000000
          Height = 34.015770000000010000
          DataField = #32066#27490#26376#20221
          DataSet = frxB06_1
          DataSetName = 'frxB06_1'
          Memo.Strings = (
            '[frxB06_1."'#32066#27490#26376#20221'"]')
        end
        object Memo19: TfrxMemoView
          Left = 328.819110000000100000
          Top = 26.456710000000000000
          Width = 34.015770000000010000
          Height = 30.236240000000000000
          Memo.Strings = (
            #33267)
        end
        object Memo6: TfrxMemoView
          Left = 196.535560000000000000
          Top = 71.811070000000000000
          Width = 177.637910000000000000
          Height = 26.456710000000000000
          HAlign = haCenter
          Memo.Strings = (
            '['#20013#29518#34399#30908#21934#25260#38957']')
        end
      end
      object PageFooter1: TfrxPageFooter
        Height = 22.677180000000000000
        Top = 468.661720000000000000
        Width = 718.110700000000000000
        object Memo1: TfrxMemoView
          Left = 642.739312740000000000
          Width = 75.590600000000000000
          Height = 18.897650000000000000
          HAlign = haRight
          Memo.Strings = (
            '[Page#]')
        end
      end
      object MasterData1: TfrxMasterData
        Height = 113.385900000000000000
        Top = 260.787570000000000000
        Width = 718.110700000000000000
        DataSet = frxB06_1
        DataSetName = 'frxB06_1'
        RowCount = 0
        object Memo3: TfrxMemoView
          Left = 3.779530000000000000
          Top = 3.779530000000079000
          Width = 52.913420000000000000
          Height = 30.236240000000000000
          Memo.Strings = (
            #29305#29518)
        end
        object Memo2: TfrxMemoView
          Left = 68.031540000000000000
          Top = 3.779530000000079000
          Width = 120.944960000000000000
          Height = 30.236240000000000000
          DataField = #29305#29518#34399#30908'1'
          DataSet = frxB06_1
          DataSetName = 'frxB06_1'
          Memo.Strings = (
            '[frxB06_1."'#29305#29518#34399#30908'1"]')
        end
        object Memo4: TfrxMemoView
          Left = 230.551330000000000000
          Top = 3.779530000000079000
          Width = 117.165430000000000000
          Height = 30.236240000000000000
          DataField = #29305#29518#34399#30908'2'
          DataSet = frxB06_1
          DataSetName = 'frxB06_1'
          Memo.Strings = (
            '[frxB06_1."'#29305#29518#34399#30908'2"]')
        end
        object Memo5: TfrxMemoView
          Left = 385.512060000000000000
          Top = 3.779530000000079000
          Width = 124.724490000000000000
          Height = 26.456710000000000000
          DataField = #29305#29518#34399#30908'3'
          DataSet = frxB06_1
          DataSetName = 'frxB06_1'
          Memo.Strings = (
            '[frxB06_1."'#29305#29518#34399#30908'3"]')
        end
        object Memo7: TfrxMemoView
          Left = 68.031540000000010000
          Top = 45.354359999999990000
          Width = 120.944960000000000000
          Height = 26.456710000000000000
          DataField = #38957#29518#34399#30908'1'
          DataSet = frxB06_1
          DataSetName = 'frxB06_1'
          Memo.Strings = (
            '[frxB06_1."'#38957#29518#34399#30908'1"]')
        end
        object Memo8: TfrxMemoView
          Left = 3.779530000000000000
          Top = 41.574830000000020000
          Width = 52.913420000000000000
          Height = 30.236240000000000000
          Memo.Strings = (
            #38957#29518)
        end
        object Memo11: TfrxMemoView
          Left = 230.551330000000000000
          Top = 45.354359999999990000
          Width = 117.165430000000000000
          Height = 26.456710000000000000
          DataField = #38957#29518#34399#30908'2'
          DataSet = frxB06_1
          DataSetName = 'frxB06_1'
          Memo.Strings = (
            '[frxB06_1."'#38957#29518#34399#30908'2"]')
        end
        object Memo12: TfrxMemoView
          Left = 385.512060000000000000
          Top = 45.354359999999990000
          Width = 120.944960000000000000
          Height = 26.456710000000000000
          DataField = #38957#29518#34399#30908'3'
          DataSet = frxB06_1
          DataSetName = 'frxB06_1'
          Memo.Strings = (
            '[frxB06_1."'#38957#29518#34399#30908'3"]')
        end
        object Memo16: TfrxMemoView
          Left = 3.779530000000000000
          Top = 79.370130000000010000
          Width = 52.913420000000000000
          Height = 30.236240000000000000
          Memo.Strings = (
            #21152#38283)
        end
        object Memo20: TfrxMemoView
          Left = 68.031540000000010000
          Top = 79.370130000000010000
          Width = 120.944960000000000000
          Height = 30.236240000000000000
          DataField = #21152#38283#34399#30908'1'
          DataSet = frxB06_1
          DataSetName = 'frxB06_1'
          Memo.Strings = (
            '[frxB06_1."'#21152#38283#34399#30908'1"]')
        end
        object Memo21: TfrxMemoView
          Left = 230.551330000000000000
          Top = 83.149659999999990000
          Width = 117.165430000000000000
          Height = 26.456710000000000000
          DataField = #21152#38283#34399#30908'2'
          DataSet = frxB06_1
          DataSetName = 'frxB06_1'
          Memo.Strings = (
            '[frxB06_1."'#21152#38283#34399#30908'2"]')
        end
      end
      object GroupHeader1: TfrxGroupHeader
        Height = 41.574830000000000000
        Top = 196.535560000000000000
        Width = 718.110700000000000000
        Condition = 'frxB06_1."'#20844#21496#21029'"'
        object Memo9: TfrxMemoView
          Left = 105.826840000000000000
          Top = 3.779529999999994000
          Width = 257.008040000000000000
          Height = 26.456710000000000000
          DataField = #20844#21496#21517#31281
          DataSet = frxB06_1
          DataSetName = 'frxB06_1'
          Memo.Strings = (
            '[frxB06_1."'#20844#21496#21517#31281'"]')
        end
        object Memo10: TfrxMemoView
          Left = 3.779530000000000000
          Top = 3.779529999999994000
          Width = 94.488250000000000000
          Height = 26.456710000000000000
          Memo.Strings = (
            #20844#21496#21517#31281)
        end
      end
      object GroupFooter1: TfrxGroupFooter
        Height = 11.338590000000000000
        Top = 396.850650000000000000
        Width = 718.110700000000000000
        object Line1: TfrxLineView
          Align = baWidth
          Width = 718.110700000000000000
          Frame.Typ = [ftTop]
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
      'TAXTYPE='#31237#21029)
    Left = 276
    Top = 104
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
      'CHARGEDATE='#25910#36027#26085#26399)
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
      'SPECIALPRIZE1='#29305#29518#34399#30908'1'
      'SPECIALPRIZE2='#29305#29518#34399#30908'2'
      'SPECIALPRIZE3='#29305#29518#34399#30908'3'
      'FIRSTPRIZE1='#38957#29518#34399#30908'1'
      'FIRSTPRIZE2='#38957#29518#34399#30908'2'
      'FIRSTPRIZE3='#38957#29518#34399#30908'3'
      'EXTRAPRIZE1='#21152#38283#34399#30908'1'
      'EXTRAPRIZE2='#21152#38283#34399#30908'2'
      'COMPID='#20844#21496#21029
      'COMPNAME='#20844#21496#21517#31281
      'STARTMONTH='#36215#22987#26376#20221
      'ENDMONTH='#32066#27490#26376#20221)
    DataSet = cdsTempory
    Left = 120
    Top = 168
  end
end
