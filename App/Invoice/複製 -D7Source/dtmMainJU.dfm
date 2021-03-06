object dtmMainJ: TdtmMainJ
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 339
  Top = 131
  Height = 552
  Width = 795
  object adoQryCommon: TADOQuery
    Connection = dtmMain.InvConnection
    CursorType = ctStatic
    Parameters = <>
    Left = 24
    Top = 8
  end
  object dspCommon: TDataSetProvider
    DataSet = adoQryCommon
    Options = [poAllowCommandText]
    Left = 32
    Top = 24
  end
  object cdsCommon: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspCommon'
    Left = 40
    Top = 32
  end
  object cdsPrintInvoice: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 48
    Top = 200
    object cdsPrintInvoiceCHECKNO: TWideStringField
      FieldName = 'CHECKNO'
      Size = 2
    end
    object cdsPrintInvoiceINVID: TWideStringField
      FieldName = 'INVID'
      Size = 10
    end
    object cdsPrintInvoiceBUSINESSID: TWideStringField
      FieldName = 'BUSINESSID'
      Size = 8
    end
    object cdsPrintInvoiceINVDATE: TDateTimeField
      FieldName = 'INVDATE'
    end
    object cdsPrintInvoiceCUSTID: TWideStringField
      FieldName = 'CUSTID'
      Size = 8
    end
    object cdsPrintInvoiceCUSTSNAME: TWideStringField
      FieldName = 'CUSTSNAME'
      Size = 50
    end
    object cdsPrintInvoiceINVTITLE: TWideStringField
      FieldName = 'INVTITLE'
      Size = 50
    end
    object cdsPrintInvoiceZIPCODE: TWideStringField
      FieldName = 'ZIPCODE'
      Size = 5
    end
    object cdsPrintInvoiceMAILADDR: TWideStringField
      FieldName = 'MAILADDR'
      Size = 60
    end
    object cdsPrintInvoiceINVFORMAT: TWideStringField
      FieldName = 'INVFORMAT'
      FixedChar = True
      Size = 1
    end
    object cdsPrintInvoiceTAXTYPE: TWideStringField
      FieldName = 'TAXTYPE'
      Size = 1
    end
    object cdsPrintInvoiceSALEAMOUNT: TBCDField
      FieldName = 'SALEAMOUNT'
      Precision = 10
      Size = 0
    end
    object cdsPrintInvoiceTAXAMOUNT: TBCDField
      FieldName = 'TAXAMOUNT'
      Precision = 10
      Size = 0
    end
    object cdsPrintInvoiceINVAMOUNT: TBCDField
      FieldName = 'INVAMOUNT'
      Precision = 10
      Size = 0
    end
    object cdsPrintInvoiceDESCRIPTION: TWideStringField
      FieldName = 'DESCRIPTION1'
      Size = 40
    end
    object cdsPrintInvoiceSTARTDATE: TDateTimeField
      FieldName = 'STARTDATE1'
    end
    object cdsPrintInvoiceENDDATE: TDateTimeField
      FieldName = 'ENDDATE1'
    end
    object cdsPrintInvoiceQUANTITY: TIntegerField
      FieldName = 'QUANTITY1'
    end
    object cdsPrintInvoiceUNITPRICE1: TFloatField
      FieldName = 'UNITPRICE1'
    end
    object cdsPrintInvoiceTOTALAMOUNT: TBCDField
      FieldName = 'TOTALAMOUNT1'
      Precision = 10
      Size = 0
    end
    object cdsPrintInvoiceDESCRIPTION2: TWideStringField
      FieldName = 'DESCRIPTION2'
      Size = 40
    end
    object cdsPrintInvoiceSTARTDATE2: TDateTimeField
      FieldName = 'STARTDATE2'
    end
    object cdsPrintInvoiceENDDATE2: TDateTimeField
      FieldName = 'ENDDATE2'
    end
    object cdsPrintInvoiceQUANTITY2: TIntegerField
      FieldName = 'QUANTITY2'
    end
    object cdsPrintInvoiceUNITPRICE2: TFloatField
      FieldName = 'UNITPRICE2'
    end
    object cdsPrintInvoiceTOTALAMOUNT2: TBCDField
      FieldName = 'TOTALAMOUNT2'
      Precision = 10
      Size = 0
    end
    object cdsPrintInvoiceDESCRIPTION3: TWideStringField
      FieldName = 'DESCRIPTION3'
      Size = 40
    end
    object cdsPrintInvoiceSTARTDATE3: TDateTimeField
      FieldName = 'STARTDATE3'
    end
    object cdsPrintInvoiceENDDATE3: TDateTimeField
      FieldName = 'ENDDATE3'
    end
    object cdsPrintInvoiceQUANTITY3: TIntegerField
      FieldName = 'QUANTITY3'
    end
    object cdsPrintInvoiceUNITPRICE3: TFloatField
      FieldName = 'UNITPRICE3'
    end
    object cdsPrintInvoiceTOTALAMOUNT3: TBCDField
      FieldName = 'TOTALAMOUNT3'
      Precision = 10
      Size = 0
    end
    object cdsPrintInvoiceDESCRIPTION4: TWideStringField
      FieldName = 'DESCRIPTION4'
      Size = 40
    end
    object cdsPrintInvoiceSTARTDATE4: TDateTimeField
      FieldName = 'STARTDATE4'
    end
    object cdsPrintInvoiceENDDATE4: TDateTimeField
      FieldName = 'ENDDATE4'
    end
    object cdsPrintInvoiceQUANTITY4: TIntegerField
      FieldName = 'QUANTITY4'
    end
    object cdsPrintInvoiceUNITPRICE4: TFloatField
      FieldName = 'UNITPRICE4'
    end
    object cdsPrintInvoiceTOTALAMOUNT4: TBCDField
      FieldName = 'TOTALAMOUNT4'
      Precision = 10
      Size = 0
    end
    object cdsPrintInvoiceDESCRIPTION5: TWideStringField
      FieldName = 'DESCRIPTION5'
      Size = 40
    end
    object cdsPrintInvoiceSTARTDATE5: TDateTimeField
      FieldName = 'STARTDATE5'
    end
    object cdsPrintInvoiceENDDATE5: TDateTimeField
      FieldName = 'ENDDATE5'
    end
    object cdsPrintInvoiceQUANTITY5: TIntegerField
      FieldName = 'QUANTITY5'
    end
    object cdsPrintInvoiceUNITPRICE5: TFloatField
      FieldName = 'UNITPRICE5'
    end
    object cdsPrintInvoiceTOTALAMOUNT5: TBCDField
      FieldName = 'TOTALAMOUNT5'
      Precision = 10
      Size = 0
    end
    object cdsPrintInvoiceMEMO1: TStringField
      FieldName = 'MEMO1'
      Size = 60
    end
  end
  object adoQryTrans2Txt: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 432
    Top = 40
  end
  object ADOQuery1: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 486
    Top = 16
  end
  object adoQryChargeMasterData: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 136
    Top = 8
  end
  object dspChargeMasterData: TDataSetProvider
    DataSet = adoQryChargeMasterData
    Options = [poAllowCommandText]
    Left = 144
    Top = 16
  end
  object cdsChargeMasterData: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspChargeMasterData'
    Left = 155
    Top = 29
  end
  object adoQryChargeDetailData: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 32
    Top = 96
  end
  object dspChargeDetailData: TDataSetProvider
    DataSet = adoQryChargeDetailData
    Options = [poAllowCommandText]
    Left = 44
    Top = 108
  end
  object cdsChargeDetailData: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspChargeDetailData'
    Left = 51
    Top = 119
  end
  object cdsChargeMasterExcel: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'DESCRIPTION'
        DataType = ftString
        Size = 40
      end
      item
        Name = 'QUANTITY'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'SALEAMOUNT'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'TAXAMOUNT'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'TOTALAMOUNT'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 48
    Top = 248
    object cdsChargeMasterExcelDESCRIPTION: TStringField
      FieldName = 'DESCRIPTION'
      Size = 40
    end
    object cdsChargeMasterExcelQUANTITY: TStringField
      FieldName = 'QUANTITY'
    end
    object cdsChargeMasterExcelSALEAMOUNT: TStringField
      FieldName = 'SALEAMOUNT'
    end
    object cdsChargeMasterExcelTAXAMOUNT: TStringField
      FieldName = 'TAXAMOUNT'
    end
    object cdsChargeMasterExcelTOTALAMOUNT: TStringField
      FieldName = 'TOTALAMOUNT'
    end
  end
  object cdsChargeDetailExcel: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'INVDATE'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'INVID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'DESCRIPTION'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'SALEAMOUNT'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'TAXAMOUNT'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'TOTALAMOUNT'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CUSTID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CUSTSNAME'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'CHARGEDATE'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'Memo1'
        DataType = ftString
        Size = 60
      end
      item
        Name = 'Memo2'
        DataType = ftString
        Size = 60
      end
      item
        Name = 'BillID'
        DataType = ftString
        Size = 15
      end
      item
        Name = 'INVTITLE'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'BusinessId'
        DataType = ftString
        Size = 8
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 48
    Top = 296
    object cdsChargeDetailExcelINVDATE: TStringField
      FieldName = 'INVDATE'
    end
    object cdsChargeDetailExcelINVID: TStringField
      FieldName = 'INVID'
    end
    object cdsChargeDetailExcelDESCRIPTION: TStringField
      FieldName = 'DESCRIPTION'
      Size = 50
    end
    object cdsChargeDetailExcelBillID: TStringField
      FieldName = 'BillID'
      Size = 15
    end
    object cdsChargeDetailExcelSALEAMOUNT: TStringField
      FieldName = 'SALEAMOUNT'
    end
    object cdsChargeDetailExcelTAXAMOUNT: TStringField
      FieldName = 'TAXAMOUNT'
    end
    object cdsChargeDetailExcelTOTALAMOUNT: TStringField
      FieldName = 'TOTALAMOUNT'
    end
    object cdsChargeDetailExcelCUSTID: TStringField
      FieldName = 'CUSTID'
    end
    object cdsChargeDetailExcelCUSTSNAME: TStringField
      FieldName = 'CUSTSNAME'
      Size = 50
    end
    object cdsChargeDetailExcelINVTITLE: TStringField
      FieldName = 'INVTITLE'
      Size = 50
    end
    object cdsChargeDetailExcelBusinessId: TStringField
      FieldName = 'BusinessId'
      Size = 8
    end
    object cdsChargeDetailExcelCHARGEDATE: TStringField
      FieldName = 'CHARGEDATE'
    end
    object cdsChargeDetailExcelMemo1: TStringField
      FieldName = 'Memo1'
      Size = 60
    end
    object cdsChargeDetailExcelMemo2: TStringField
      FieldName = 'Memo2'
      Size = 60
    end
  end
  object adoQrySelectInv016: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM INV016 WHERE SEQ IS NULL')
    Left = 352
    Top = 24
    object adoQrySelectInv016SHOULDBEASSIGNED: TStringField
      DisplayLabel = #26159#21542#38283#31435
      FieldName = 'SHOULDBEASSIGNED'
      FixedChar = True
      Size = 1
    end
    object adoQrySelectInv016SEQ: TStringField
      DisplayLabel = #27969#27700#34399
      FieldName = 'SEQ'
      Size = 15
    end
    object adoQrySelectInv016COMPID: TStringField
      FieldName = 'COMPID'
      Visible = False
      Size = 6
    end
    object adoQrySelectInv016CUSTID: TStringField
      DisplayLabel = #23458#25142#20195#30908
      FieldName = 'CUSTID'
      Size = 8
    end
    object adoQrySelectInv016TEL: TStringField
      FieldName = 'TEL'
      Visible = False
    end
    object adoQrySelectInv016BUSINESSID: TStringField
      FieldName = 'BUSINESSID'
      Visible = False
      Size = 8
    end
    object adoQrySelectInv016TITLE: TStringField
      FieldName = 'TITLE'
      Visible = False
      Size = 50
    end
    object adoQrySelectInv016ZIPCODE: TStringField
      FieldName = 'ZIPCODE'
      Visible = False
      Size = 5
    end
    object adoQrySelectInv016INVADDR: TStringField
      FieldName = 'INVADDR'
      Visible = False
      Size = 60
    end
    object adoQrySelectInv016MAILADDR: TStringField
      FieldName = 'MAILADDR'
      Visible = False
      Size = 60
    end
    object adoQrySelectInv016BEASSIGNEDINVID: TStringField
      FieldName = 'BEASSIGNEDINVID'
      Visible = False
      FixedChar = True
      Size = 1
    end
    object adoQrySelectInv016CHARGEDATE: TDateTimeField
      DisplayLabel = #25910#36027#26085#26399
      FieldName = 'CHARGEDATE'
    end
    object adoQrySelectInv016TAXTYPE: TStringField
      DisplayLabel = #31237#21029
      FieldName = 'TAXTYPE'
      FixedChar = True
      Size = 1
    end
    object adoQrySelectInv016TAXRATE: TBCDField
      DisplayLabel = #29151#26989#31237#29575
      FieldName = 'TAXRATE'
      Precision = 5
      Size = 2
    end
    object adoQrySelectInv016SALEAMOUNT: TBCDField
      DisplayLabel = #37559#21806#38989
      FieldName = 'SALEAMOUNT'
      Precision = 10
      Size = 0
    end
    object adoQrySelectInv016TAXAMOUNT: TBCDField
      DisplayLabel = #31237#38989
      FieldName = 'TAXAMOUNT'
      Precision = 10
      Size = 0
    end
    object adoQrySelectInv016INVAMOUNT: TBCDField
      DisplayLabel = #30332#31080#37329#38989
      FieldName = 'INVAMOUNT'
      Precision = 10
      Size = 0
    end
    object adoQrySelectInv016ISVALID: TStringField
      DisplayLabel = #26159#21542#26377#25928
      FieldName = 'ISVALID'
      FixedChar = True
      Size = 1
    end
    object adoQrySelectInv016HOWTOCREATE: TStringField
      FieldName = 'HOWTOCREATE'
      FixedChar = True
      Size = 1
    end
    object adoQrySelectInv016UPTTIME: TDateTimeField
      FieldName = 'UPTTIME'
      Visible = False
    end
    object adoQrySelectInv016UPTEN: TStringField
      FieldName = 'UPTEN'
      Visible = False
    end
  end
  object adoQrySelectInv017: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM INV017 WHERE SEQ IS NULL')
    Left = 272
    Top = 8
    object adoQrySelectInv017ShouldBeAssigned: TStringField
      DisplayLabel = #26159#21542#38283#31435
      FieldName = 'ShouldBeAssigned'
      Size = 8
    end
    object adoQrySelectInv017SEQ: TStringField
      DisplayLabel = #27969#27700#34399
      FieldName = 'SEQ'
      Size = 15
    end
    object adoQrySelectInv017BILLID: TStringField
      DisplayLabel = #25910#36027#21934#34399
      FieldName = 'BILLID'
      Size = 15
    end
    object adoQrySelectInv017BILLIDITEMNO: TIntegerField
      DisplayLabel = #25910#36027#21934#34399#38917#27425
      FieldName = 'BILLIDITEMNO'
    end
    object adoQrySelectInv017TAXTYPE: TStringField
      DisplayLabel = #31237#21029
      FieldName = 'TAXTYPE'
      FixedChar = True
      Size = 1
    end
    object adoQrySelectInv017CHARGEDATE: TDateTimeField
      DisplayLabel = #25910#36027#26085#26399
      FieldName = 'CHARGEDATE'
    end
    object adoQrySelectInv017ITEMID: TStringField
      DisplayLabel = #21697#21517#20195#30908
      FieldName = 'ITEMID'
      Size = 7
    end
    object adoQrySelectInv017DESCRIPTION: TStringField
      DisplayLabel = #21697#21517
      FieldName = 'DESCRIPTION'
      Size = 40
    end
    object adoQrySelectInv017QUANTITY: TIntegerField
      DisplayLabel = #25976#37327
      FieldName = 'QUANTITY'
    end
    object adoQrySelectInv017UNITPRICE: TBCDField
      DisplayLabel = #21934#20729
      FieldName = 'UNITPRICE'
      Precision = 12
      Size = 2
    end
    object adoQrySelectInv017TAXRATE: TBCDField
      DisplayLabel = #29151#26989#31237#29575
      FieldName = 'TAXRATE'
      Precision = 5
      Size = 2
    end
    object adoQrySelectInv017TAXAMOUNT: TBCDField
      DisplayLabel = #31237#38989
      FieldName = 'TAXAMOUNT'
      Precision = 10
      Size = 0
    end
    object adoQrySelectInv017TOTALAMOUNT: TBCDField
      DisplayLabel = #32317#37329#38989
      FieldName = 'TOTALAMOUNT'
      Precision = 10
      Size = 0
    end
    object adoQrySelectInv017STARTDATE: TDateTimeField
      DisplayLabel = #26377#25928#36215#22987#26085
      FieldName = 'STARTDATE'
    end
    object adoQrySelectInv017ENDDATE: TDateTimeField
      DisplayLabel = #26377#25928#25130#27490#26085
      FieldName = 'ENDDATE'
    end
    object adoQrySelectInv017CHARGEEN: TStringField
      FieldName = 'CHARGEEN'
    end
  end
  object InV099DataSetProvider: TDataSetProvider
    DataSet = adoQryCommon
    Left = 134
    Top = 83
  end
  object Inv099DataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'InV099DataSetProvider'
    Left = 142
    Top = 99
  end
  object DataSource1: TDataSource
    AutoEdit = False
    DataSet = Inv099DataSet
    Left = 150
    Top = 115
  end
  object adoSourcePreFix: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 272
    Top = 64
  end
  object cdsUsePrefix: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'Order'
        DataType = ftString
        Size = 2
      end
      item
        Name = 'Prefix'
        DataType = ftString
        Size = 2
      end
      item
        Name = 'StartNum'
        DataType = ftString
        Size = 8
      end
      item
        Name = 'EndNum'
        DataType = ftString
        Size = 8
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 48
    Top = 352
    object cdsUsePrefixOrder: TStringField
      FieldName = 'Order'
      Size = 2
    end
    object cdsUsePrefixPrefix: TStringField
      FieldName = 'Prefix'
      Size = 2
    end
    object cdsUsePrefixStartNum: TStringField
      FieldName = 'StartNum'
      Size = 8
    end
    object cdsUsePrefixEndNum: TStringField
      FieldName = 'EndNum'
      Size = 8
    end
  end
  object adoInv007ToInv024: TADOQuery
    CacheSize = 1000
    Connection = dtmMain.InvConnection
    CursorType = ctOpenForwardOnly
    Parameters = <>
    Left = 272
    Top = 128
  end
  object adoQryInv024: TADOQuery
    Connection = dtmMain.InvConnection
    CursorType = ctStatic
    Parameters = <>
    Left = 416
    Top = 120
  end
  object adoInv001Code: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    SQL.Strings = (
      'select * from inv001')
    Left = 356
    Top = 326
  end
  object adoInv002Code: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 356
    Top = 388
  end
  object adoInv019Code: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 560
    Top = 210
  end
  object adoInv006Code: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 272
    Top = 198
  end
  object adoInv022Code: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 272
    Top = 244
  end
  object adoInv025Code: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 272
    Top = 292
  end
  object cdsPrefix: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'No'
        DataType = ftInteger
      end
      item
        Name = 'Prefix'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'LastInvDate'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'IdentifyId1'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'IdentifyId2'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CompId'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'YearMonth'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'StartNum'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'Memo'
        DataType = ftString
        Size = 50
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 48
    Top = 400
    object cdsPrefixNo: TIntegerField
      DisplayLabel = #38918#24207
      DisplayWidth = 4
      FieldName = 'No'
    end
    object cdsPrefixPrefix: TStringField
      DisplayLabel = #23383#36556
      DisplayWidth = 6
      FieldName = 'Prefix'
    end
    object cdsPrefixLastInvDate: TStringField
      DisplayLabel = #26368#24460#30332#31080#38283#31435#26085
      FieldName = 'LastInvDate'
    end
    object cdsPrefixMemo: TStringField
      DisplayLabel = #20633#35387
      FieldName = 'Memo'
      Size = 50
    end
    object cdsPrefixIdentifyId1: TStringField
      FieldName = 'IdentifyId1'
      Visible = False
    end
    object cdsPrefixIdentifyId2: TStringField
      FieldName = 'IdentifyId2'
      Visible = False
    end
    object cdsPrefixCompId: TStringField
      FieldName = 'CompId'
      Visible = False
    end
    object cdsPrefixYearMonth: TStringField
      FieldName = 'YearMonth'
      Visible = False
    end
    object cdsPrefixStartNum: TStringField
      FieldName = 'StartNum'
      Visible = False
    end
  end
  object AdoQryInsertInv016: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM INV016'
      'WHERE 1 = 2')
    Left = 574
    Top = 8
  end
  object dspQryInsertInv016: TDataSetProvider
    DataSet = AdoQryInsertInv016
    ResolveToDataSet = True
    Left = 602
    Top = 54
  end
  object cdsQryInsertInv016: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspQryInsertInv016'
    Left = 560
    Top = 54
  end
  object AdoQryInsertInv017: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM INV017'
      'WHERE 1 = 2')
    Left = 688
    Top = 24
  end
  object dspQryInsertInv017: TDataSetProvider
    DataSet = AdoQryInsertInv017
    ResolveToDataSet = True
    Left = 696
    Top = 40
  end
  object cdsQryInsertInv017: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspQryInsertInv017'
    Left = 708
    Top = 60
  end
  object spAssignInvID: TADOStoredProc
    Connection = dtmMain.InvConnection
    ProcedureName = 'SF_ASSIGNINVID'
    Parameters = <>
    Left = 352
    Top = 164
  end
  object adoQryInv031: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM INV031 ORDER BY ROWID')
    Left = 696
    Top = 156
  end
  object dspQInv031: TDataSetProvider
    DataSet = adoQryInv031
    Options = [poAllowCommandText]
    Left = 704
    Top = 164
  end
  object cdsQInv031: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspQInv031'
    Left = 712
    Top = 172
  end
  object adoQryInv033: TADOQuery
    Connection = dtmMain.InvConnection
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM INV033 WHERE BILLID IS NULL')
    Left = 540
    Top = 117
  end
  object dspInv033: TDataSetProvider
    DataSet = adoQryInv033
    Options = [poAllowCommandText]
    Left = 548
    Top = 125
  end
  object cdsInv033: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspInv033'
    Left = 558
    Top = 141
  end
  object adoInv099Code: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 272
    Top = 340
  end
  object cdsCheckBusinessID: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'Memo'
        DataType = ftString
        Size = 100
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 144
    Top = 200
    object cdsCheckBusinessIDMemo: TStringField
      DisplayLabel = #32113#19968#32232#34399#27298#26597#32080#26524
      FieldName = 'Memo'
      Size = 100
    end
  end
  object cdsCheckJumpInvID: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'Memo'
        DataType = ftString
        Size = 100
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 144
    Top = 256
    object cdsCheckJumpInvIDMemo: TStringField
      DisplayLabel = #36339#34399#27298#26597#32080#26524
      FieldName = 'Memo'
      Size = 100
    end
  end
  object adoInv007AmtCheck: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 471
    Top = 168
  end
  object adoInv008AmtCheck: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 472
    Top = 216
  end
  object cdsCheckAmount: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'InvDate'
        DataType = ftDate
      end
      item
        Name = 'InvID'
        DataType = ftString
        Size = 12
      end
      item
        Name = 'MasterAmt'
        DataType = ftInteger
      end
      item
        Name = 'DetailAmt'
        DataType = ftInteger
      end
      item
        Name = 'IsObsolete'
        DataType = ftString
        Size = 10
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 472
    Top = 264
    object cdsCheckAmountInvDate: TDateField
      DisplayLabel = #30332#31080#26085#26399
      FieldName = 'InvDate'
    end
    object cdsCheckAmountInvID: TStringField
      DisplayLabel = #30332#31080#34399#30908
      DisplayWidth = 12
      FieldName = 'InvID'
      Size = 12
    end
    object cdsCheckAmountMasterAmt: TIntegerField
      DisplayLabel = #20027#37329#38989
      FieldName = 'MasterAmt'
    end
    object cdsCheckAmountDetailAmt: TIntegerField
      DisplayLabel = #21103#37329#38989
      FieldName = 'DetailAmt'
    end
    object cdsCheckAmountIsObsolete: TStringField
      DisplayLabel = #26159#21542#20316#24290
      FieldName = 'IsObsolete'
      Size = 10
    end
  end
  object cdsCheckMisData: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'Reason'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'MisCustID'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'MisCustName'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'BillNo'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'InvID'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'MisTitle'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'MisBusinessID'
        DataType = ftString
        Size = 15
      end
      item
        Name = 'InvTitle'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'InvBusinessID'
        DataType = ftString
        Size = 15
      end
      item
        Name = 'InvCustID'
        DataType = ftString
        Size = 10
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 144
    Top = 304
    object cdsCheckMisDataReason: TStringField
      DisplayLabel = #21407#22240
      FieldName = 'Reason'
      Size = 50
    end
    object cdsCheckMisDataMisCustID: TStringField
      DisplayLabel = #23458#32232
      FieldName = 'MisCustID'
      Size = 10
    end
    object cdsCheckMisDataMisCustName: TStringField
      DisplayLabel = #25910#20214#20154
      DisplayWidth = 20
      FieldName = 'MisCustName'
      Size = 50
    end
    object cdsCheckMisDataBillNo: TStringField
      DisplayLabel = #25910#36027#21934#34399
      DisplayWidth = 15
      FieldName = 'BillNo'
    end
    object cdsCheckMisDataInvID: TStringField
      DisplayLabel = #30332#31080#34399#30908
      FieldName = 'InvID'
      Size = 10
    end
    object cdsCheckMisDataMisTitle: TStringField
      DisplayLabel = #23458#26381#25260#38957
      DisplayWidth = 20
      FieldName = 'MisTitle'
      Size = 50
    end
    object cdsCheckMisDataMisBusinessID: TStringField
      DisplayLabel = #23458#26381#32113#32232
      FieldName = 'MisBusinessID'
      Size = 15
    end
    object cdsCheckMisDataInvTitle: TStringField
      DisplayLabel = #30332#31080#25260#38957
      DisplayWidth = 20
      FieldName = 'InvTitle'
      Size = 50
    end
    object cdsCheckMisDataInvBusinessID: TStringField
      DisplayLabel = #30332#31080#32113#32232
      FieldName = 'InvBusinessID'
      Size = 15
    end
    object cdsCheckMisDataInvCustID: TStringField
      DisplayLabel = #30332#31080#23458#32232
      FieldName = 'InvCustID'
      Size = 10
    end
  end
  object adoInv004Code: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 356
    Top = 268
  end
  object cdsCreateTypeInfo: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'InvDate'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CreateType1'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CreateType2'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CreateType3'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CreateType4'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'IsObsolete'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'RecordTotal'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 144
    Top = 400
    object cdsCreateTypeInfoInvDate: TStringField
      FieldName = 'InvDate'
    end
    object cdsCreateTypeInfoIsObsolete: TStringField
      FieldName = 'IsObsolete'
      Size = 10
    end
    object cdsCreateTypeInfoCreateType1: TStringField
      DisplayLabel = 'mis'#25291#27284', '#38928#38283
      FieldName = 'CreateType1'
    end
    object cdsCreateTypeInfoCreateType2: TStringField
      FieldName = 'CreateType2'
    end
    object cdsCreateTypeInfoCreateType3: TStringField
      FieldName = 'CreateType3'
    end
    object cdsCreateTypeInfoCreateType4: TStringField
      DisplayLabel = 'invoice create, '#24460#38283
      FieldName = 'CreateType4'
    end
    object cdsCreateTypeInfoRecordTotal: TStringField
      FieldName = 'RecordTotal'
    end
  end
  object cdsObsoleteInvoiceData: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'InvID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'InvDate'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'InvAmount'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ObsoleteReason'
        DataType = ftString
        Size = 25
      end
      item
        Name = 'CustID'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'HowToCreate'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 144
    Top = 352
    object cdsObsoleteInvoiceDataInvID: TStringField
      FieldName = 'InvID'
    end
    object cdsObsoleteInvoiceDataInvDate: TStringField
      FieldName = 'InvDate'
    end
    object cdsObsoleteInvoiceDataInvAmount: TStringField
      FieldName = 'InvAmount'
    end
    object cdsObsoleteInvoiceDataObsoleteReason: TStringField
      FieldName = 'ObsoleteReason'
      Size = 25
    end
    object cdsObsoleteInvoiceDataCustID: TStringField
      FieldName = 'CustID'
      Size = 10
    end
    object cdsObsoleteInvoiceDataHowToCreate: TStringField
      FieldName = 'HowToCreate'
    end
  end
  object adoQryInv018: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 616
    Top = 112
  end
  object adoInv003Code: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 272
    Top = 388
  end
  object cdsSelectInv016: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'ShouldBeAssigned'
        DataType = ftString
        Size = 8
      end
      item
        Name = 'Seq'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CustId'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'ChargeDate'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'TaxType'
        DataType = ftString
        Size = 5
      end
      item
        Name = 'TaxRate'
        DataType = ftString
        Size = 8
      end
      item
        Name = 'SaleAmount'
        DataType = ftString
        Size = 15
      end
      item
        Name = 'TaxAmount'
        DataType = ftString
        Size = 15
      end
      item
        Name = 'InvAmount'
        DataType = ftString
        Size = 15
      end
      item
        Name = 'IsValid'
        DataType = ftString
        Size = 8
      end
      item
        Name = 'HowToCreate'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 352
    Top = 80
    object cdsSelectInv016ShouldBeAssigned: TStringField
      DisplayLabel = #26159#21542#38283#31435
      FieldName = 'ShouldBeAssigned'
      Size = 8
    end
    object cdsSelectInv016Seq: TStringField
      DisplayLabel = #27969#27700#34399
      DisplayWidth = 17
      FieldName = 'Seq'
    end
    object cdsSelectInv016CustId: TStringField
      DisplayLabel = #23458#25142#32232#34399
      FieldName = 'CustId'
      Size = 10
    end
    object cdsSelectInv016ChargeDate: TStringField
      DisplayLabel = #25910#36027#26085#26399
      DisplayWidth = 12
      FieldName = 'ChargeDate'
    end
    object cdsSelectInv016TaxType: TStringField
      DisplayLabel = #31237#21029
      FieldName = 'TaxType'
      Size = 5
    end
    object cdsSelectInv016TaxRate: TStringField
      DisplayLabel = #29151#26989#31237#29575
      FieldName = 'TaxRate'
      Size = 8
    end
    object cdsSelectInv016SaleAmount: TStringField
      DisplayLabel = #37559#21806#38989
      DisplayWidth = 10
      FieldName = 'SaleAmount'
      Size = 15
    end
    object cdsSelectInv016TaxAmount: TStringField
      DisplayLabel = #31237#38989
      DisplayWidth = 8
      FieldName = 'TaxAmount'
      Size = 15
    end
    object cdsSelectInv016InvAmount: TStringField
      DisplayLabel = #30332#31080#37329#38989
      DisplayWidth = 10
      FieldName = 'InvAmount'
      Size = 15
    end
    object cdsSelectInv016IsValid: TStringField
      DisplayLabel = #26159#21542#26377#25928
      FieldName = 'IsValid'
      Size = 8
    end
    object cdsSelectInv016HowToCreate: TStringField
      DisplayLabel = #22914#20309#38283#31435
      DisplayWidth = 17
      FieldName = 'HowToCreate'
    end
  end
  object XMLDoc: TXMLDocument
    Left = 48
    Top = 456
    DOMVendorDesc = 'MSXML'
  end
  object cdsWitchOne: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'SEQ'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'TAXTYPE'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'TAXRATE'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 704
    Top = 104
  end
  object adoInv012Code: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 356
    Top = 220
  end
  object adoQryCommon2: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    Left = 144
    Top = 456
  end
  object cdsTemp: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 476
    Top = 324
  end
  object adoInv028Code: TADOQuery
    CacheSize = 1000
    Connection = dtmMain.InvConnection
    Parameters = <>
    SQL.Strings = (
      'select * from INV028 order by itemid')
    Left = 564
    Top = 316
  end
  object adoInv040Code: TADOQuery
    Connection = dtmMain.InvConnection
    Parameters = <>
    SQL.Strings = (
      'select * from inv001')
    Left = 416
    Top = 330
  end
  object adoInvD09: TADOConnection
    LoginPrompt = False
    Left = 417
    Top = 354
  end
end
