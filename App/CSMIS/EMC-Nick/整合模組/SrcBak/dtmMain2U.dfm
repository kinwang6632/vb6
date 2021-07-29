object dtmMain2: TdtmMain2
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 298
  Top = 157
  Height = 472
  HorizontalOffset = 775
  VerticalOffset = 136
  Width = 680
  object cdsSo110: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO110'
    RemoteServer = DCOMConnection1
    OnCalcFields = cdsSo110CalcFields
    Left = 91
    Top = 8
    object cdsSo110FORMULA_ID: TStringField
      DisplayLabel = #20195#30908
      FieldName = 'FORMULA_ID'
      Size = 4
    end
    object cdsSo110FORMULA_NAME: TStringField
      DisplayLabel = #21517#31281
      FieldName = 'FORMULA_NAME'
      Size = 50
    end
    object cdsSo110FORMULA_DESC: TStringField
      DisplayLabel = #35498#26126
      FieldName = 'FORMULA_DESC'
      Size = 100
    end
    object cdsSo110REF_NO: TStringField
      FieldName = 'REF_NO'
      FixedChar = True
      Size = 1
    end
    object cdsSo110REFNO_NAME: TStringField
      DisplayLabel = #20195#34399
      FieldKind = fkCalculated
      FieldName = 'REFNO_NAME'
      Size = 40
      Calculated = True
    end
    object cdsSo110STOPFLAG: TIntegerField
      FieldName = 'STOPFLAG'
    end
    object cdsSo110STOPFLAG_NAME: TStringField
      DisplayLabel = #26159#21542#20572#29992
      FieldKind = fkCalculated
      FieldName = 'STOPFLAG_NAME'
      Size = 2
      Calculated = True
    end
    object cdsSo110OPERATOR: TStringField
      DisplayLabel = #30064#21205#20154#21729
      FieldName = 'OPERATOR'
    end
    object cdsSo110UPDTIME: TStringField
      DisplayLabel = #30064#21205#26178#38291
      FieldName = 'UPDTIME'
    end
  end
  object DCOMConnection1: TDCOMConnection
    ServerGUID = '{775D7E70-4ED9-41BF-B4EF-C49216E38A77}'
    ServerName = 'Emc_Separate.Emc_Separate1'
    Left = 33
    Top = 9
  end
  object cdsSo112M: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO112M'
    RemoteServer = DCOMConnection1
    AfterScroll = cdsSo112MAfterScroll
    Left = 147
    Top = 8
    object cdsSo112MPRODUCT_ID: TStringField
      DisplayLabel = #29986#21697#20195#30908
      FieldName = 'PRODUCT_ID'
      Size = 4
    end
    object cdsSo112MPRODUCT_NAME: TStringField
      DisplayLabel = #29986#21697#21517#31281
      FieldName = 'PRODUCT_NAME'
      Size = 30
    end
    object cdsSo112MPROVIDER_ID: TStringField
      DisplayLabel = #38971#36947#21830#20195#30908
      FieldName = 'PROVIDER_ID'
      Size = 4
    end
    object cdsSo112MPROVIDER_NAME: TStringField
      DisplayLabel = #38971#36947#21830#21517#31281
      FieldKind = fkLookup
      FieldName = 'PROVIDER_NAME'
      LookupDataSet = cdsSo113CT
      LookupKeyFields = 'PROVIDER_ID'
      LookupResultField = 'PROVIDER_NAME'
      KeyFields = 'PROVIDER_ID'
      Size = 30
      Lookup = True
    end
    object cdsSo112MPRICE: TIntegerField
      DisplayLabel = #21407#23450#20729
      FieldName = 'PRICE'
    end
    object cdsSo112MOPERATOR: TStringField
      DisplayLabel = #30064#21205#20154#21729
      FieldName = 'OPERATOR'
    end
    object cdsSo112MTYPE: TStringField
      DisplayLabel = #39006#22411
      FieldName = 'TYPE'
      Size = 2
    end
    object cdsSo112MREF_NO: TIntegerField
      DisplayLabel = #21443#32771#34399
      FieldName = 'REF_NO'
    end
    object cdsSo112MUPDTIME: TStringField
      DisplayLabel = #30064#21205#26178#38291
      FieldName = 'UPDTIME'
    end
    object cdsSo112MUSEBASEFORMULA: TStringField
      DisplayLabel = #21855#29992#22522#26412#20844#24335
      FieldName = 'USEBASEFORMULA'
      FixedChar = True
      Size = 1
    end
    object cdsSo112MPROVIDER_PERCENT: TIntegerField
      DisplayLabel = #38971#36947#21830#25286#24115#30334#20998#27604
      FieldName = 'PROVIDER_PERCENT'
    end
    object cdsSo112MSO_PERCENT: TIntegerField
      FieldName = 'SO_PERCENT'
    end
  end
  object cdsCd024CT: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspCd024CT'
    RemoteServer = DCOMConnection1
    Left = 274
    Top = 8
    object cdsCd024CTCODENO: TStringField
      FieldName = 'CODENO'
      Size = 3
    end
    object cdsCd024CTDESCRIPTION: TStringField
      FieldName = 'DESCRIPTION'
    end
    object cdsCd024CTREFNO: TIntegerField
      FieldName = 'REFNO'
    end
    object cdsCd024CTPAYFLAG: TIntegerField
      FieldName = 'PAYFLAG'
    end
    object cdsCd024CTCITEMCODE: TIntegerField
      FieldName = 'CITEMCODE'
    end
    object cdsCd024CTCITEMNAME: TStringField
      FieldName = 'CITEMNAME'
      Size = 12
    end
    object cdsCd024CTUPDTIME: TStringField
      FieldName = 'UPDTIME'
      Size = 19
    end
    object cdsCd024CTUPDEN: TStringField
      FieldName = 'UPDEN'
    end
    object cdsCd024CTCOMPCODE: TIntegerField
      FieldName = 'COMPCODE'
    end
    object cdsCd024CTSTOPFLAG: TIntegerField
      FieldName = 'STOPFLAG'
    end
    object cdsCd024CTCHANCEDAYS: TIntegerField
      FieldName = 'CHANCEDAYS'
    end
    object cdsCd024CTCHANNELID: TStringField
      FieldName = 'CHANNELID'
      Size = 12
    end
  end
  object cdsSqlStatement: TClientDataSet
    Aggregates = <>
    Params = <
      item
        DataType = ftString
        Name = 'PRODUCT_ID'
        ParamType = ptInput
      end
      item
        DataType = ftString
        Name = 'FORMULA_ID'
        ParamType = ptInput
      end>
    ProviderName = 'dspSqlStatement'
    RemoteServer = DCOMConnection1
    Left = 413
    Top = 7
  end
  object cdsSo113CT: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSo113CT'
    RemoteServer = DCOMConnection1
    OnCalcFields = cdsSo113CTCalcFields
    Left = 339
    Top = 8
    object cdsSo113CTPROVIDER_ID: TStringField
      DisplayLabel = #20195#30908
      FieldName = 'PROVIDER_ID'
      Size = 4
    end
    object cdsSo113CTPROVIDER_NAME: TStringField
      DisplayLabel = #21517#31281
      FieldName = 'PROVIDER_NAME'
      Size = 50
    end
    object cdsSo113CTATTRIBUTE: TIntegerField
      FieldName = 'ATTRIBUTE'
    end
    object cdsSo113CTATTRIBUTE_NAME: TStringField
      DisplayLabel = #23660#24615
      FieldKind = fkCalculated
      FieldName = 'ATTRIBUTE_NAME'
      Size = 35
      Calculated = True
    end
    object cdsSo113CTCONTACTEE: TStringField
      DisplayLabel = #32879#32097#20154
      FieldName = 'CONTACTEE'
    end
    object cdsSo113CTTEL: TStringField
      DisplayLabel = #38651#35441
      FieldName = 'TEL'
    end
    object cdsSo113CTADDRESS: TStringField
      DisplayLabel = #22320#22336
      FieldName = 'ADDRESS'
      Size = 90
    end
    object cdsSo113CTNOTES: TStringField
      DisplayLabel = #20633#35387
      FieldName = 'NOTES'
      Size = 100
    end
    object cdsSo113CTSTOPFLAG: TIntegerField
      FieldName = 'STOPFLAG'
    end
    object cdsSo113CTSTOPFLAG_NAME: TStringField
      DisplayLabel = #26159#21542#20572#29992
      FieldKind = fkCalculated
      FieldName = 'STOPFLAG_NAME'
      Size = 2
      Calculated = True
    end
    object cdsSo113CTOPERATOR: TStringField
      DisplayLabel = #30064#21205#20154#21729
      FieldName = 'OPERATOR'
    end
    object cdsSo113CTUPDTIME: TStringField
      DisplayLabel = #30064#21205#26178#38291
      FieldName = 'UPDTIME'
    end
  end
  object cdsSo111D: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO111D'
    RemoteServer = DCOMConnection1
    OnCalcFields = cdsSo111DCalcFields
    Left = 211
    Top = 8
    object cdsSo111DCOMP_TYPE: TStringField
      DisplayLabel = #31995#32113#21488'('#20844#21496#21029')or'#20379#25033#21830#39006#22411
      FieldName = 'COMP_TYPE'
      Visible = False
      FixedChar = True
      Size = 1
    end
    object cdsSo111DSO_PROVIDER_ID: TStringField
      DisplayLabel = #20844#21496#21029'/'#20379#25033#21830#20195#30908
      FieldName = 'SO_PROVIDER_ID'
      Visible = False
      Size = 10
    end
    object cdsSo111DSO_PROVIDER_DESC: TStringField
      DisplayLabel = #20844#21496#21029'/'#38971#36947#21830#21517#31281
      FieldName = 'SO_PROVIDER_DESC'
      LookupDataSet = cdsSo113Cd039
      LookupKeyFields = 'CLASSIFY_ID'
      LookupResultField = 'DESC'
      KeyFields = 'COMP_TYPE'
      Size = 50
    end
    object cdsSo111DPRODUCT_ID: TStringField
      DisplayLabel = #29986#21697#20195#30908
      FieldName = 'PRODUCT_ID'
      Size = 4
    end
    object cdsSo111DPRODUCT_NAME: TStringField
      DisplayLabel = #29986#21697#21517#31281
      FieldName = 'PRODUCT_NAME'
      Size = 30
    end
    object cdsSo111DFORMULA_NAME_LU: TStringField
      DisplayLabel = #20844#24335#21517#31281
      DisplayWidth = 30
      FieldKind = fkLookup
      FieldName = 'FORMULA_NAME_LU'
      LookupDataSet = cdsSo110
      LookupKeyFields = 'FORMULA_ID'
      LookupResultField = 'FORMULA_NAME'
      KeyFields = 'FORMULA_ID'
      Size = 50
      Lookup = True
    end
    object cdsSo111DFORMULA_ID: TStringField
      DisplayLabel = #20844#24335#20195#30908
      FieldName = 'FORMULA_ID'
      Visible = False
      EditMask = 'AAAA;0;_'
      Size = 4
    end
    object cdsSo111DCOMPUTE_ISSUE: TIntegerField
      FieldName = 'COMPUTE_ISSUE'
    end
    object cdsSo111DCOMPUTE_ISSUE_NAME: TStringField
      DisplayLabel = #35336#31639#20381#25818
      FieldKind = fkLookup
      FieldName = 'COMPUTE_ISSUE_NAME'
      LookupDataSet = cdsCompute_issue
      LookupKeyFields = 'CODE'
      LookupResultField = 'COMPUTE_ISSUE'
      KeyFields = 'COMPUTE_ISSUE'
      Size = 16
      Lookup = True
    end
    object cdsSo111DRANGE_UNIT: TStringField
      FieldName = 'RANGE_UNIT'
      FixedChar = True
      Size = 1
    end
    object cdsSo111DRANGE_UNIT_NAME: TStringField
      DisplayLabel = #25142#25976#27604#36611#21934#20301
      FieldKind = fkLookup
      FieldName = 'RANGE_UNIT_NAME'
      LookupDataSet = cdsRange_Unit
      LookupKeyFields = 'RANGE_UNIT'
      LookupResultField = 'RANGE_UNIT_NAME'
      KeyFields = 'RANGE_UNIT'
      Size = 6
      Lookup = True
    end
    object cdsSo111DSUBSCRIBER_COUNT_PERCENT1: TBCDField
      FieldName = 'SUBSCRIBER_COUNT_PERCENT1'
      Precision = 8
      Size = 2
    end
    object cdsSo111DSUBSCRIBER_COUNT_PERCENT2: TBCDField
      FieldName = 'SUBSCRIBER_COUNT_PERCENT2'
      Precision = 8
      Size = 2
    end
    object cdsSo111DSUBSCRIBER_COUNT_PERCENT3: TBCDField
      FieldName = 'SUBSCRIBER_COUNT_PERCENT3'
      Precision = 8
      Size = 2
    end
    object cdsSo111DSUBSCRIBER_COUNT_PERCENT4: TBCDField
      FieldName = 'SUBSCRIBER_COUNT_PERCENT4'
      Precision = 8
      Size = 2
    end
    object cdsSo111DSUBSCRIBER_COUNT_PERCENT5: TBCDField
      FieldName = 'SUBSCRIBER_COUNT_PERCENT5'
      Precision = 8
      Size = 2
    end
    object cdsSo111DAMOUNT_PERCENT1: TBCDField
      FieldName = 'AMOUNT_PERCENT1'
      Precision = 10
      Size = 2
    end
    object cdsSo111DAMOUNT_PERCENT2: TBCDField
      FieldName = 'AMOUNT_PERCENT2'
      Precision = 10
      Size = 2
    end
    object cdsSo111DAMOUNT_PERCENT3: TBCDField
      FieldName = 'AMOUNT_PERCENT3'
      Precision = 10
      Size = 2
    end
    object cdsSo111DAMOUNT_PERCENT4: TBCDField
      FieldName = 'AMOUNT_PERCENT4'
      Precision = 10
      Size = 2
    end
    object cdsSo111DAMOUNT_PERCENT5: TBCDField
      FieldName = 'AMOUNT_PERCENT5'
      Precision = 10
      Size = 2
    end
    object cdsSo111DUPDTIME: TStringField
      DisplayLabel = #30064#21205#26178#38291
      FieldName = 'UPDTIME'
    end
    object cdsSo111DOPERATOR: TStringField
      DisplayLabel = #30064#21205#20154#21729
      FieldName = 'OPERATOR'
    end
  end
  object cdsSqlDelete: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSqlDelete'
    RemoteServer = DCOMConnection1
    Left = 494
    Top = 7
  end
  object cdsCompute_issue: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'CODE'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'COMPUTE_ISSUE'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 344
    Top = 56
  end
  object cdsCD039: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspCD039'
    RemoteServer = DCOMConnection1
    Left = 36
    Top = 61
  end
  object cdsSo113Cd039: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'CLASSIFY_ID'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CODE'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'DESC'
        DataType = ftString
        Size = 50
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 170
    Top = 58
    object cdsSo113Cd039CLASSIFY_ID: TStringField
      DisplayLabel = #39006#22411
      FieldName = 'CLASSIFY_ID'
      Size = 1
    end
    object cdsSo113Cd039CODE: TStringField
      DisplayLabel = #20195#30908
      FieldName = 'CODE'
      Size = 10
    end
    object cdsSo113Cd039DESC: TStringField
      DisplayLabel = #21517#31281
      FieldName = 'DESC'
      Size = 50
    end
  end
  object cdsSo113: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSo113'
    RemoteServer = DCOMConnection1
    Left = 96
    Top = 64
  end
  object cdsRange_Unit: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'RANGE_UNIT'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'RANGE_UNIT_NAME'
        DataType = ftString
        Size = 6
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 256
    Top = 56
  end
  object cdsSo114N: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 424
    Top = 56
  end
  object cdsCodeCd039: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspCodeCd039'
    RemoteServer = DCOMConnection1
    Left = 496
    Top = 56
  end
  object cdsParam: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspParam'
    RemoteServer = DCOMConnection1
    Left = 38
    Top = 115
  end
  object cdsProductInfoN: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'ProductID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'PackageID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'RealCharge'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'OriginalPrice'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'FormulaID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ProviderID'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 454
    Top = 111
  end
  object cdsRefDataAN: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'PRODUCT_ID'
        DataType = ftString
        Size = 4
      end
      item
        Name = 'CITEMCODE'
        DataType = ftString
        Size = 3
      end
      item
        Name = 'FORMULA_ID'
        DataType = ftString
        Size = 4
      end
      item
        Name = 'REF_NO'
        DataType = ftString
        Size = 1
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 168
    Top = 111
  end
  object cdsV_ChargeInfoForTally: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspChargeInfoForTally'
    RemoteServer = DCOMConnection1
    Left = 272
    Top = 109
  end
  object cdsChargeDataN: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspChargeInfoForTally'
    RemoteServer = DCOMConnection1
    Left = 376
    Top = 109
  end
  object cdsPackageId: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspPackageId'
    RemoteServer = DCOMConnection1
    Left = 40
    Top = 163
  end
  object cdsRefDataN: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'PRODUCT_ID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CITEMCODE'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'FORMULA_ID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'REF_NO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'PROVIDER_ID'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 93
    Top = 112
  end
  object cdsPackageInfoN: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'PackageID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'TotalOriginalPrice'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'TotalRealCharge'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 536
    Top = 109
  end
  object cdsSearSo112: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSearSo112'
    RemoteServer = DCOMConnection1
    Left = 120
    Top = 163
  end
  object cdsRefData2: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'PRODUCT_ID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CITEMCODE'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'FORMULA_ID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'REF_NO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'PROVIDER_ID'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    ProviderName = 'DataSetProvider1'
    RemoteServer = DCOMConnection1
    StoreDefs = True
    Left = 376
    Top = 163
  end
  object cdsCodeCD039A: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dsoCodeCD039A'
    RemoteServer = DCOMConnection1
    Left = 32
    Top = 259
  end
  object cdsCom: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspCom'
    RemoteServer = DCOMConnection1
    Left = 16
    Top = 315
  end
  object cdsRefData: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspRefData'
    RemoteServer = DCOMConnection1
    Left = 528
    Top = 251
  end
  object cdsRefDataA: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspRefDataA'
    RemoteServer = DCOMConnection1
    Left = 534
    Top = 304
  end
  object cdsChargeData: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspChargeData'
    RemoteServer = DCOMConnection1
    Left = 461
    Top = 305
  end
  object cdsCustData: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspCustData'
    RemoteServer = DCOMConnection1
    Left = 594
    Top = 251
  end
  object cdsProductInfo: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'ProductID'
        DataType = ftString
        Size = 4
      end
      item
        Name = 'PackageID'
        DataType = ftString
        Size = 3
      end
      item
        Name = 'ProviderID'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'RealCharge'
        DataType = ftFloat
      end
      item
        Name = 'OrigionalPrice'
        DataType = ftFloat
      end
      item
        Name = 'FormulaID'
        DataType = ftString
        Size = 4
      end
      item
        Name = 'UseBaseFormula'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'Provider_Percent'
        DataType = ftInteger
      end
      item
        Name = 'So_Percent'
        DataType = ftInteger
      end
      item
        Name = 'IsInCome'
        DataType = ftBoolean
      end
      item
        Name = 'FirstFlag'
        DataType = ftString
        Size = 1
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 112
    Top = 259
    object cdsProductInfoProductID: TStringField
      FieldName = 'ProductID'
      Size = 4
    end
    object cdsProductInfoPackageID: TStringField
      FieldName = 'PackageID'
      Size = 3
    end
    object cdsProductInfoProviderID: TStringField
      FieldName = 'ProviderID'
      Size = 10
    end
    object cdsProductInfoRealCharge: TFloatField
      FieldName = 'RealCharge'
    end
    object cdsProductInfoOrigionalPrice: TFloatField
      FieldName = 'OrigionalPrice'
    end
    object cdsProductInfoFormulaID: TStringField
      FieldName = 'FormulaID'
      Size = 4
    end
    object cdsProductInfoUseBaseFormula: TStringField
      FieldName = 'UseBaseFormula'
      Size = 1
    end
    object cdsProductInfoProvider_Percent: TIntegerField
      FieldName = 'Provider_Percent'
    end
    object cdsProductInfoSo_Percent: TIntegerField
      FieldName = 'So_Percent'
    end
    object cdsProductInfoIsInCome: TBooleanField
      FieldName = 'IsInCome'
    end
    object cdsProductInfoFirstFlag: TStringField
      FieldName = 'FirstFlag'
      Size = 1
    end
  end
  object cdsPackageInfo: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'PackageID'
        DataType = ftString
        Size = 3
      end
      item
        Name = 'TotalOrigionalPrice'
        DataType = ftFloat
      end
      item
        Name = 'TotalRealCharge'
        DataType = ftFloat
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 80
    Top = 315
    object cdsPackageInfoPackageID: TStringField
      FieldName = 'PackageID'
      Size = 3
    end
    object cdsPackageInfoTotalOrigionalPrice: TFloatField
      FieldName = 'TotalOrigionalPrice'
    end
    object cdsPackageInfoTotalRealCharge: TFloatField
      FieldName = 'TotalRealCharge'
    end
  end
  object cdsSO115: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO115'
    RemoteServer = DCOMConnection1
    Left = 459
    Top = 253
  end
  object cdsSO114: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO114'
    RemoteServer = DCOMConnection1
    Left = 328
    Top = 257
  end
  object cdsSO116: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO116'
    RemoteServer = DCOMConnection1
    Left = 328
    Top = 307
  end
  object cdsSO117: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO117'
    RemoteServer = DCOMConnection1
    Left = 392
    Top = 307
  end
  object cdsCd024Look: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspCd024Look'
    RemoteServer = DCOMConnection1
    Left = 616
    Top = 160
    object cdsCd024LookCODENO: TStringField
      FieldName = 'CODENO'
      Size = 3
    end
    object cdsCd024LookCODENO_DESCRIPTION: TStringField
      FieldName = 'CODENO_DESCRIPTION'
      Size = 24
    end
    object cdsCd024LookDESCRIPTION: TStringField
      FieldName = 'DESCRIPTION'
    end
  end
  object cdsAttribute: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'CODENO'
        DataType = ftString
        Size = 5
      end
      item
        Name = 'DESCRIPTION'
        DataType = ftString
        Size = 40
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 576
    Top = 8
  end
  object cdsSo113Look: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSo113Look'
    RemoteServer = DCOMConnection1
    Left = 456
    Top = 163
    object cdsSo113LookPROVIDER_ID: TStringField
      FieldName = 'PROVIDER_ID'
      Size = 10
    end
    object cdsSo113LookPROVIDER_ID_NAME: TStringField
      FieldName = 'PROVIDER_ID_NAME'
      Size = 61
    end
    object cdsSo113LookPROVIDER_NAME: TStringField
      FieldName = 'PROVIDER_NAME'
      Size = 50
    end
  end
  object cdsSO115B: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO115B'
    RemoteServer = DCOMConnection1
    Left = 272
    Top = 258
  end
  object cdsPriv: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'cdsPriv'
    RemoteServer = DCOMConnection1
    Left = 272
    Top = 306
  end
  object cdsSo110Look: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO110Look'
    RemoteServer = DCOMConnection1
    Left = 536
    Top = 163
    object cdsSo110LookFORMULA_ID: TStringField
      FieldName = 'FORMULA_ID'
      Size = 4
    end
    object cdsSo110LookFORMULA_ID_NAME: TStringField
      FieldName = 'FORMULA_ID_NAME'
      ReadOnly = True
      Size = 55
    end
    object cdsSo110LookFORMULA_NAME: TStringField
      FieldName = 'FORMULA_NAME'
      Size = 50
    end
  end
  object cdsChannelCounts: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'ProductID'
        DataType = ftString
        Size = 4
      end
      item
        Name = 'ChannelName'
        DataType = ftString
        Size = 30
      end
      item
        Name = 'CitemCode'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ChannelTotalViewCounts'
        DataType = ftInteger
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 168
    Top = 314
    object cdsChannelCountsProductID: TStringField
      FieldName = 'ProductID'
      Size = 4
    end
    object cdsChannelCountsChannelName: TStringField
      FieldName = 'ChannelName'
      Size = 30
    end
    object cdsChannelCountsCitemCode: TStringField
      FieldName = 'CitemCode'
    end
    object cdsChannelCountsChannelTotalViewCounts: TIntegerField
      FieldName = 'ChannelTotalViewCounts'
    end
  end
  object cdsSO114B: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO114B'
    RemoteServer = DCOMConnection1
    Left = 392
    Top = 255
  end
  object cdsTallyRefData1: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspTallyRefData1'
    RemoteServer = DCOMConnection1
    Left = 272
    Top = 360
  end
  object cdsTallyRefData2: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspTallyRefData2'
    RemoteServer = DCOMConnection1
    Left = 360
    Top = 360
  end
  object cdsSO119: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO119'
    RemoteServer = DCOMConnection1
    Left = 464
    Top = 360
  end
  object cdsCD019: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspCd019'
    RemoteServer = DCOMConnection1
    Left = 592
    Top = 312
  end
  object cdsTempSO119: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspTempSO119'
    RemoteServer = DCOMConnection1
    Left = 536
    Top = 360
  end
  object cdsRefNo: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'CODENO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'DESCRIPTION'
        DataType = ftString
        Size = 40
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 576
    Top = 64
    object cdsRefNoCODENO: TStringField
      FieldName = 'CODENO'
    end
    object cdsRefNoDESCRIPTION: TStringField
      FieldName = 'DESCRIPTION'
      Size = 40
    end
  end
  object cdsSO114C: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO114C'
    RemoteServer = DCOMConnection1
    Left = 608
    Top = 360
  end
  object ClientDataSet1: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'ProductID'
        DataType = ftString
        Size = 4
      end
      item
        Name = 'PackageID'
        DataType = ftInteger
      end
      item
        Name = 'ProviderID'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'RealCharge'
        DataType = ftFloat
      end
      item
        Name = 'OrigionalPrice'
        DataType = ftFloat
      end
      item
        Name = 'FormulaID'
        DataType = ftString
        Size = 4
      end
      item
        Name = 'UseBaseFormula'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'Provider_Percent'
        DataType = ftInteger
      end
      item
        Name = 'So_Percent'
        DataType = ftInteger
      end
      item
        Name = 'IsInCome'
        DataType = ftBoolean
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 168
    Top = 235
    object StringField1: TStringField
      FieldName = 'ProductID'
      Size = 4
    end
    object IntegerField1: TIntegerField
      FieldName = 'PackageID'
    end
    object StringField2: TStringField
      FieldName = 'ProviderID'
      Size = 10
    end
    object FloatField1: TFloatField
      FieldName = 'RealCharge'
    end
    object FloatField2: TFloatField
      FieldName = 'OrigionalPrice'
    end
    object StringField3: TStringField
      FieldName = 'FormulaID'
      Size = 4
    end
    object StringField4: TStringField
      FieldName = 'UseBaseFormula'
      Size = 1
    end
    object IntegerField2: TIntegerField
      FieldName = 'Provider_Percent'
    end
    object IntegerField3: TIntegerField
      FieldName = 'So_Percent'
    end
    object BooleanField1: TBooleanField
      FieldName = 'IsInCome'
    end
  end
  object ClientDataSet2: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'PackageID'
        DataType = ftInteger
      end
      item
        Name = 'TotalOrigionalPrice'
        DataType = ftFloat
      end
      item
        Name = 'TotalRealCharge'
        DataType = ftFloat
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 80
    Top = 371
    object IntegerField4: TIntegerField
      FieldName = 'PackageID'
    end
    object FloatField3: TFloatField
      FieldName = 'TotalOrigionalPrice'
    end
    object FloatField4: TFloatField
      FieldName = 'TotalRealCharge'
    end
  end
  object cdsSO116Rpt: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'CustID'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'CustName'
        DataType = ftString
        Size = 30
      end
      item
        Name = 'Channel_String'
        DataType = ftWideString
        Size = 4000
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 328
    Top = 216
    object cdsSO116RptCustID: TStringField
      FieldName = 'CustID'
      Size = 50
    end
    object cdsSO116RptCustName: TStringField
      FieldName = 'CustName'
      Size = 30
    end
    object cdsSO116RptChannel_String: TWideStringField
      FieldName = 'Channel_String'
      Size = 4000
    end
  end
  object cdsSO114Rpt: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'PROVIDERDESC'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'ProductName'
        DataType = ftString
        Size = 30
      end
      item
        Name = 'FirstCount'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'LastCount'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'AddCount'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ReduceCount'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'Income'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'Outcome'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'DiffCome'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'EMC_Benefit'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'So_Benefit'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'Provider_Benefit'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 400
    Top = 216
    object cdsSO114RptPROVIDERDESC: TStringField
      FieldName = 'PROVIDERDESC'
      Size = 50
    end
    object cdsSO114RptProductName: TStringField
      FieldName = 'ProductName'
      Size = 30
    end
    object cdsSO114RptFirstCount: TStringField
      FieldName = 'FirstCount'
    end
    object cdsSO114RptLastCount: TStringField
      FieldName = 'LastCount'
    end
    object cdsSO114RptAddCount: TStringField
      FieldName = 'AddCount'
    end
    object cdsSO114RptReduceCount: TStringField
      FieldName = 'ReduceCount'
    end
    object cdsSO114RptIncome: TStringField
      FieldName = 'Income'
    end
    object cdsSO114RptOutcome: TStringField
      FieldName = 'Outcome'
    end
    object cdsSO114RptDiffCome: TStringField
      FieldName = 'DiffCome'
    end
    object cdsSO114RptEMC_Benefit: TStringField
      FieldName = 'EMC_Benefit'
    end
    object cdsSO114RptSo_Benefit: TStringField
      FieldName = 'So_Benefit'
    end
    object cdsSO114RptProvider_Benefit: TStringField
      FieldName = 'Provider_Benefit'
    end
  end
  object cdsSO114SubTotal: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'FirstCount'
        DataType = ftInteger
      end
      item
        Name = 'LastCount'
        DataType = ftInteger
      end
      item
        Name = 'AddCount'
        DataType = ftInteger
      end
      item
        Name = 'ReduceCount'
        DataType = ftInteger
      end
      item
        Name = 'Income'
        DataType = ftFloat
      end
      item
        Name = 'Outcome'
        DataType = ftFloat
      end
      item
        Name = 'DiffCome'
        DataType = ftFloat
      end
      item
        Name = 'EMC_Benefit'
        DataType = ftFloat
      end
      item
        Name = 'So_Benefit'
        DataType = ftFloat
      end
      item
        Name = 'Provider_Benefit'
        DataType = ftFloat
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 488
    Top = 216
    object cdsSO114SubTotalFirstCount: TIntegerField
      FieldName = 'FirstCount'
    end
    object cdsSO114SubTotalLastCount: TIntegerField
      FieldName = 'LastCount'
    end
    object cdsSO114SubTotalAddCount: TIntegerField
      FieldName = 'AddCount'
    end
    object cdsSO114SubTotalReduceCount: TIntegerField
      FieldName = 'ReduceCount'
    end
    object cdsSO114SubTotalIncome: TFloatField
      FieldName = 'Income'
    end
    object cdsSO114SubTotalOutcome: TFloatField
      FieldName = 'Outcome'
    end
    object cdsSO114SubTotalDiffCome: TFloatField
      FieldName = 'DiffCome'
    end
    object cdsSO114SubTotalEMC_Benefit: TFloatField
      FieldName = 'EMC_Benefit'
    end
    object cdsSO114SubTotalSo_Benefit: TFloatField
      FieldName = 'So_Benefit'
    end
    object cdsSO114SubTotalProvider_Benefit: TFloatField
      FieldName = 'Provider_Benefit'
    end
  end
  object cdsSO114Total: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'FirstCount'
        DataType = ftInteger
      end
      item
        Name = 'LastCount'
        DataType = ftInteger
      end
      item
        Name = 'AddCount'
        DataType = ftInteger
      end
      item
        Name = 'ReduceCount'
        DataType = ftInteger
      end
      item
        Name = 'Income'
        DataType = ftFloat
      end
      item
        Name = 'Outcome'
        DataType = ftFloat
      end
      item
        Name = 'DiffCome'
        DataType = ftFloat
      end
      item
        Name = 'EMC_Benefit'
        DataType = ftFloat
      end
      item
        Name = 'So_Benefit'
        DataType = ftFloat
      end
      item
        Name = 'Provider_Benefit'
        DataType = ftFloat
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 584
    Top = 216
    object IntegerField5: TIntegerField
      FieldName = 'FirstCount'
    end
    object IntegerField6: TIntegerField
      FieldName = 'LastCount'
    end
    object IntegerField7: TIntegerField
      FieldName = 'AddCount'
    end
    object IntegerField8: TIntegerField
      FieldName = 'ReduceCount'
    end
    object FloatField5: TFloatField
      FieldName = 'Income'
    end
    object FloatField6: TFloatField
      FieldName = 'Outcome'
    end
    object FloatField7: TFloatField
      FieldName = 'DiffCome'
    end
    object FloatField8: TFloatField
      FieldName = 'EMC_Benefit'
    end
    object FloatField9: TFloatField
      FieldName = 'So_Benefit'
    end
    object FloatField10: TFloatField
      FieldName = 'Provider_Benefit'
    end
  end
  object cdsCodeCD019: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspCodeCD019'
    RemoteServer = DCOMConnection1
    Left = 728
    Top = 8
  end
  object cdsSO033: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO033'
    RemoteServer = DCOMConnection1
    Left = 728
    Top = 64
  end
  object cdsSO131: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO131'
    RemoteServer = DCOMConnection1
    Left = 728
    Top = 112
    object cdsSO131COMPUTEYM: TStringField
      DisplayLabel = #35336#31639#24180#26376
      DisplayWidth = 100
      FieldName = 'COMPUTEYM'
      Size = 100
    end
    object cdsSO131REALORSHOULDDATE: TIntegerField
      DisplayLabel = #20197#23526#25910#25110#25033#25910#26085#26399#35336#31639
      FieldName = 'REALORSHOULDDATE'
    end
    object cdsSO131SEQNO: TIntegerField
      DisplayLabel = #24207#34399
      FieldName = 'SEQNO'
    end
    object cdsSO131BILLNO: TStringField
      DisplayLabel = #21934#25818#32232#34399
      FieldName = 'BILLNO'
      Size = 15
    end
    object cdsSO131ITEM: TIntegerField
      DisplayLabel = #38917#27425
      FieldName = 'ITEM'
    end
    object cdsSO131CUSTID: TIntegerField
      DisplayLabel = #23458#25142#32232#34399
      FieldName = 'CUSTID'
    end
    object cdsSO131CITEMCODE: TIntegerField
      DisplayLabel = #25910#36027#38917#30446#20195#34399
      FieldName = 'CITEMCODE'
    end
    object cdsSO131CITEMNAME: TStringField
      DisplayLabel = #25910#36027#38917#30446#21517#31281
      FieldName = 'CITEMNAME'
    end
    object cdsSO131SHOULDDATE: TDateTimeField
      DisplayLabel = #25033#25910#26085#26399
      FieldName = 'SHOULDDATE'
    end
    object cdsSO131SHOULDAMT: TIntegerField
      DisplayLabel = #25033#25910#37329#38989
      FieldName = 'SHOULDAMT'
    end
    object cdsSO131REALDATE: TDateTimeField
      DisplayLabel = #23526#25910#26085#26399
      FieldName = 'REALDATE'
    end
    object cdsSO131REALAMT: TIntegerField
      DisplayLabel = #23526#25910#37329#38989
      FieldName = 'REALAMT'
    end
    object cdsSO131REALPERIOD: TIntegerField
      DisplayLabel = #25910#36027#26399#25976
      FieldName = 'REALPERIOD'
    end
    object cdsSO131REALSTARTDATE: TDateTimeField
      DisplayLabel = #26377#25928#36215#22987#26085#26399
      FieldName = 'REALSTARTDATE'
    end
    object cdsSO131REALSTOPDATE: TDateTimeField
      DisplayLabel = #26377#25928#25130#27490#26085#26399
      FieldName = 'REALSTOPDATE'
    end
    object cdsSO131COMPCODE: TIntegerField
      DisplayLabel = #20844#21496#21029
      FieldName = 'COMPCODE'
    end
    object cdsSO131ORDERNO: TStringField
      DisplayLabel = #35330#21934#21934#34399
      FieldName = 'ORDERNO'
      Size = 15
    end
    object cdsSO131SBILLNO: TStringField
      DisplayLabel = #21407#22987#21934#34399
      FieldName = 'SBILLNO'
      Size = 15
    end
    object cdsSO131SITEM: TIntegerField
      DisplayLabel = #21407#22987#38917#27425
      FieldName = 'SITEM'
    end
    object cdsSO131MEDIACODE: TIntegerField
      DisplayLabel = #20171#32057#23186#20171#20195#30908
      FieldName = 'MEDIACODE'
    end
    object cdsSO131MEDIANAME: TStringField
      DisplayLabel = #20171#32057#23186#20171#21517#31281
      FieldName = 'MEDIANAME'
    end
    object cdsSO131ACCEPTEN: TStringField
      DisplayLabel = #21463#29702#20154#20195#34399
      FieldName = 'ACCEPTEN'
      Size = 10
    end
    object cdsSO131ACCEPTNAME: TStringField
      DisplayLabel = #21463#29702#20154#22995#21517
      FieldName = 'ACCEPTNAME'
    end
    object cdsSO131INTROID: TStringField
      DisplayLabel = #20171#32057#20154#20195#34399
      FieldName = 'INTROID'
      Size = 10
    end
    object cdsSO131INTRONAME: TStringField
      DisplayLabel = #20171#32057#20154#22995#21517
      FieldName = 'INTRONAME'
      Size = 50
    end
    object cdsSO131NOTES: TStringField
      DisplayLabel = #20633#35387
      FieldName = 'NOTES'
      Size = 255
    end
  end
  object cdsSpSO131: TClientDataSet
    Aggregates = <>
    Params = <
      item
        DataType = ftBCD
        Name = 'RETURN_VALUE'
        ParamType = ptResult
      end
      item
        DataType = ftString
        Name = 'P_COMPCODE'
        ParamType = ptInput
      end
      item
        DataType = ftString
        Name = 'P_SERVICETYPE'
        ParamType = ptInput
      end
      item
        DataType = ftString
        Name = 'P_COMPUTEYM'
        ParamType = ptInput
      end
      item
        DataType = ftString
        Name = 'P_STARTDATE'
        ParamType = ptInput
      end
      item
        DataType = ftString
        Name = 'P_ENDDATE'
        ParamType = ptInput
      end
      item
        DataType = ftString
        Name = 'P_CHARGEITEMSQL'
        ParamType = ptInput
      end
      item
        DataType = ftBCD
        Name = 'P_REALORSHOULDDATE'
        ParamType = ptInput
      end
      item
        DataType = ftBCD
        Name = 'P_RETCODE'
        ParamType = ptOutput
      end
      item
        DataType = ftString
        Name = 'P_RETMSG'
        ParamType = ptOutput
      end>
    ProviderName = 'dspSpSO131'
    RemoteServer = DCOMConnection1
    Left = 728
    Top = 160
  end
  object cdsCd019A: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspCd019A'
    RemoteServer = DCOMConnection1
    Left = 728
    Top = 216
    object cdsCd019CODENO: TIntegerField
      FieldName = 'CODENO'
    end
    object cdsCd019DESCRIPTION: TStringField
      FieldName = 'DESCRIPTION'
    end
    object cdsCd019REFDATA: TStringField
      FieldName = 'REFDATA'
      ReadOnly = True
      Size = 61
    end
  end
  object cdsSO131Excel: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO131Excel'
    RemoteServer = DCOMConnection1
    OnCalcFields = cdsSO131ExcelCalcFields
    Left = 728
    Top = 320
    object cdsSO131ExcelCOMPUTEYM: TStringField
      DisplayLabel = #32113#35336#24180#26376
      FieldName = 'COMPUTEYM'
      Size = 6
    end
    object cdsSO131ExcelCOMPUTEYM2: TStringField
      DisplayLabel = #32113#35336#24180#26376
      FieldKind = fkInternalCalc
      FieldName = 'COMPUTEYM2'
      Size = 6
    end
    object cdsSO131ExcelREALORSHOULDDATE: TIntegerField
      DisplayLabel = #25033#25910#25110#23526#25910#26085#26399
      FieldName = 'REALORSHOULDDATE'
    end
    object cdsSO131ExcelSEQNO: TIntegerField
      DisplayLabel = #24207#34399
      FieldName = 'SEQNO'
    end
    object cdsSO131ExcelBILLNO: TStringField
      DisplayLabel = #21934#25818#32232#34399
      FieldName = 'BILLNO'
      Size = 15
    end
    object cdsSO131ExcelITEM: TIntegerField
      DisplayLabel = #38917#27425
      FieldName = 'ITEM'
    end
    object cdsSO131ExcelCUSTID: TIntegerField
      DisplayLabel = #23458#25142#32232#34399
      FieldName = 'CUSTID'
    end
    object cdsSO131ExcelCITEMCODE: TIntegerField
      DisplayLabel = #25910#36027#38917#30446#20195#34399
      FieldName = 'CITEMCODE'
    end
    object cdsSO131ExcelCITEMNAME: TStringField
      DisplayLabel = #25910#36027#38917#30446#21517#31281
      FieldName = 'CITEMNAME'
    end
    object cdsSO131ExcelSHOULDDATE: TDateTimeField
      DisplayLabel = #25033#25910#26085#26399
      FieldName = 'SHOULDDATE'
    end
    object cdsSO131ExcelSHOULDDATE2: TStringField
      DisplayLabel = #25033#25910#26085#26399
      FieldKind = fkInternalCalc
      FieldName = 'SHOULDDATE2'
    end
    object cdsSO131ExcelSHOULDAMT: TStringField
      DisplayLabel = #25033#25910#37329#38989
      FieldName = 'SHOULDAMT'
      ReadOnly = True
      Visible = False
      Size = 40
    end
    object cdsSO131ExcelCalcShouldAmt: TIntegerField
      DisplayLabel = #25033#25910#37329#38989
      FieldKind = fkInternalCalc
      FieldName = 'CalcShouldAmt'
    end
    object cdsSO131ExcelREALDATE: TDateTimeField
      DisplayLabel = #23526#25910#26085#26399
      FieldName = 'REALDATE'
    end
    object cdsSO131ExcelREALDATE2: TStringField
      DisplayLabel = #23526#25910#26085#26399
      FieldKind = fkInternalCalc
      FieldName = 'REALDATE2'
    end
    object cdsSO131ExcelREALAMT: TStringField
      FieldName = 'REALAMT'
      ReadOnly = True
      Visible = False
      Size = 40
    end
    object cdsSO131ExcelCalcRealAmt: TIntegerField
      DisplayLabel = #23526#25910#37329#38989
      FieldKind = fkInternalCalc
      FieldName = 'CalcRealAmt'
    end
    object cdsSO131ExcelREALPERIOD: TIntegerField
      DisplayLabel = #25910#36027#26399#25976
      FieldName = 'REALPERIOD'
    end
    object cdsSO131ExcelREALSTARTDATE: TDateTimeField
      DisplayLabel = #26377#25928#36215#22987#26085#26399
      FieldName = 'REALSTARTDATE'
    end
    object cdsSO131ExcelREALSTARTDATE2: TStringField
      DisplayLabel = #26377#25928#36215#22987#26085#26399
      FieldKind = fkInternalCalc
      FieldName = 'REALSTARTDATE2'
    end
    object cdsSO131ExcelREALSTOPDATE: TDateTimeField
      DisplayLabel = #26377#25928#25130#27490#26085#26399
      FieldName = 'REALSTOPDATE'
    end
    object cdsSO131ExcelREALSTOPDATE2: TStringField
      DisplayLabel = #26377#25928#25130#27490#26085#26399
      FieldKind = fkInternalCalc
      FieldName = 'REALSTOPDATE2'
    end
    object cdsSO131ExcelCOMPCODE: TIntegerField
      DisplayLabel = #20844#21496#21029
      FieldName = 'COMPCODE'
    end
    object cdsSO131ExcelORDERNO: TStringField
      DisplayLabel = #35330#21934#21934#34399
      FieldName = 'ORDERNO'
      Size = 15
    end
    object cdsSO131ExcelSBILLNO: TStringField
      DisplayLabel = #21407#22987#21934#34399
      FieldName = 'SBILLNO'
      Size = 15
    end
    object cdsSO131ExcelSITEM: TIntegerField
      DisplayLabel = #21407#22987#38917#27425
      FieldName = 'SITEM'
    end
    object cdsSO131ExcelMEDIACODE: TIntegerField
      DisplayLabel = #20171#32057#23186#20171#20195#30908
      FieldName = 'MEDIACODE'
    end
    object cdsSO131ExcelMEDIANAME: TStringField
      DisplayLabel = #20171#32057#23186#20171#21517#31281
      FieldName = 'MEDIANAME'
    end
    object cdsSO131ExcelACCEPTEN: TStringField
      DisplayLabel = #21463#29702#20154#20195#34399
      FieldName = 'ACCEPTEN'
      Size = 10
    end
    object cdsSO131ExcelACCEPTNAME: TStringField
      DisplayLabel = #21463#29702#20154#22995#21517
      FieldName = 'ACCEPTNAME'
    end
    object cdsSO131ExcelINTROID: TStringField
      DisplayLabel = #20171#32057#20154#20195#34399
      FieldName = 'INTROID'
      Size = 10
    end
    object cdsSO131ExcelINTRONAME: TStringField
      DisplayLabel = #20171#32057#20154#22995#21517
      FieldName = 'INTRONAME'
      Size = 50
    end
    object cdsSO131ExcelNOTES: TStringField
      DisplayLabel = #20633#35387
      FieldName = 'NOTES'
      Size = 255
    end
    object cdsSO131ExcelREFNO: TIntegerField
      FieldName = 'REFNO'
    end
    object cdsSO131ExcelCalcRefNo: TStringField
      DisplayLabel = #26159#21542#28858#21574#24115
      FieldKind = fkInternalCalc
      FieldName = 'CalcRefNo'
      Size = 30
    end
  end
  object cdsOtherSO131Excel: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspOtherSO131Excel'
    RemoteServer = DCOMConnection1
    OnCalcFields = cdsOtherSO131ExcelCalcFields
    Left = 728
    Top = 368
    object StringField5: TStringField
      DisplayLabel = #32113#35336#24180#26376
      FieldName = 'COMPUTEYM'
      Size = 6
    end
    object cdsOtherSO131ExcelCOMPUTEYM2: TStringField
      DisplayLabel = #32113#35336#24180#26376
      FieldKind = fkInternalCalc
      FieldName = 'COMPUTEYM2'
      Size = 6
    end
    object IntegerField9: TIntegerField
      DisplayLabel = #25033#25910#25110#23526#25910#26085#26399
      FieldName = 'REALORSHOULDDATE'
    end
    object IntegerField10: TIntegerField
      DisplayLabel = #24207#34399
      FieldName = 'SEQNO'
    end
    object StringField6: TStringField
      DisplayLabel = #21934#25818#32232#34399
      FieldName = 'BILLNO'
      Size = 15
    end
    object IntegerField11: TIntegerField
      DisplayLabel = #38917#27425
      FieldName = 'ITEM'
    end
    object IntegerField12: TIntegerField
      DisplayLabel = #23458#25142#32232#34399
      FieldName = 'CUSTID'
    end
    object IntegerField13: TIntegerField
      DisplayLabel = #25910#36027#38917#30446#20195#34399
      FieldName = 'CITEMCODE'
    end
    object StringField7: TStringField
      DisplayLabel = #25910#36027#38917#30446#21517#31281
      FieldName = 'CITEMNAME'
    end
    object DateTimeField1: TDateTimeField
      DisplayLabel = #25033#25910#26085#26399
      FieldName = 'SHOULDDATE'
    end
    object cdsOtherSO131ExcelSHOULDDATE2: TStringField
      DisplayLabel = #25033#25910#26085#26399
      FieldKind = fkInternalCalc
      FieldName = 'SHOULDDATE2'
    end
    object StringField8: TStringField
      FieldName = 'SHOULDAMT'
      ReadOnly = True
      Visible = False
      Size = 40
    end
    object IntegerField14: TIntegerField
      DisplayLabel = #25033#25910#37329#38989
      FieldKind = fkCalculated
      FieldName = 'CalcShouldAmt'
      Calculated = True
    end
    object DateTimeField2: TDateTimeField
      DisplayLabel = #23526#25910#26085#26399
      FieldName = 'REALDATE'
    end
    object cdsOtherSO131ExcelREALDATE2: TStringField
      DisplayLabel = #23526#25910#26085#26399
      FieldKind = fkInternalCalc
      FieldName = 'REALDATE2'
    end
    object StringField9: TStringField
      FieldName = 'REALAMT'
      ReadOnly = True
      Visible = False
      Size = 40
    end
    object IntegerField15: TIntegerField
      DisplayLabel = #23526#25910#37329#38989
      FieldKind = fkCalculated
      FieldName = 'CalcRealAmt'
      Calculated = True
    end
    object IntegerField16: TIntegerField
      DisplayLabel = #25910#36027#26399#25976
      FieldName = 'REALPERIOD'
    end
    object DateTimeField3: TDateTimeField
      DisplayLabel = #26377#25928#36215#22987#26085#26399
      FieldName = 'REALSTARTDATE'
    end
    object cdsOtherSO131ExcelREALSTARTDATE2: TStringField
      DisplayLabel = #26377#25928#36215#22987#26085#26399
      FieldKind = fkInternalCalc
      FieldName = 'REALSTARTDATE2'
    end
    object DateTimeField4: TDateTimeField
      DisplayLabel = #26377#25928#25130#27490#26085#26399
      FieldName = 'REALSTOPDATE'
    end
    object cdsOtherSO131ExcelREALSTOPDATE2: TStringField
      DisplayLabel = #26377#25928#25130#27490#26085#26399
      FieldKind = fkInternalCalc
      FieldName = 'REALSTOPDATE2'
    end
    object IntegerField17: TIntegerField
      DisplayLabel = #20844#21496#21029
      FieldName = 'COMPCODE'
    end
    object StringField10: TStringField
      DisplayLabel = #35330#21934#21934#34399
      FieldName = 'ORDERNO'
      Size = 15
    end
    object StringField11: TStringField
      DisplayLabel = #21407#22987#21934#34399
      FieldName = 'SBILLNO'
      Size = 15
    end
    object IntegerField18: TIntegerField
      DisplayLabel = #21407#22987#38917#27425
      FieldName = 'SITEM'
    end
    object IntegerField19: TIntegerField
      DisplayLabel = #20171#32057#23186#20171#20195#30908
      FieldName = 'MEDIACODE'
    end
    object StringField12: TStringField
      DisplayLabel = #20171#32057#23186#20171#21517#31281
      FieldName = 'MEDIANAME'
    end
    object StringField13: TStringField
      DisplayLabel = #21463#29702#20154#20195#34399
      FieldName = 'ACCEPTEN'
      Size = 10
    end
    object StringField14: TStringField
      DisplayLabel = #21463#29702#20154#22995#21517
      FieldName = 'ACCEPTNAME'
    end
    object StringField15: TStringField
      DisplayLabel = #20171#32057#20154#20195#34399
      FieldName = 'INTROID'
      Size = 10
    end
    object StringField16: TStringField
      DisplayLabel = #20171#32057#20154#22995#21517
      FieldName = 'INTRONAME'
      Size = 50
    end
    object StringField17: TStringField
      DisplayLabel = #20633#35387
      FieldName = 'NOTES'
      Size = 255
    end
    object cdsOtherSO131ExcelREFNO: TIntegerField
      FieldName = 'REFNO'
    end
    object cdsOtherSO131ExcelCalcRefNo: TStringField
      DisplayLabel = #26159#21542#28858#21574#24115
      FieldKind = fkCalculated
      FieldName = 'CalcRefNo'
      Size = 30
      Calculated = True
    end
  end
  object cdsSO113ForUSeful: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO113ForUSeful'
    RemoteServer = DCOMConnection1
    Left = 352
    Top = 480
  end
  object cdsPercentXmlData: TClientDataSet
    Active = True
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'sProviderID'
        DataType = ftString
        Size = 10
      end
      item
        Name = 'sProviderName'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'nPercent'
        DataType = ftInteger
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    AfterPost = cdsPercentXmlDataAfterPost
    BeforeScroll = cdsPercentXmlDataBeforeScroll
    Left = 552
    Top = 480
    Data = {
      6D0000009619E0BD0100000018000000030000000000030000006D000B735072
      6F766964657249440100490000000100055749445448020002000A000D735072
      6F76696465724E616D650100490000000100055749445448020002003200086E
      50657263656E7404000100000000000000}
    object cdsPercentXmlDatasProviderID: TStringField
      FieldName = 'sProviderID'
      Size = 10
    end
    object cdsPercentXmlDatasProviderName: TStringField
      FieldName = 'sProviderName'
      Size = 50
    end
    object cdsPercentXmlDatanPercent: TIntegerField
      FieldName = 'nPercent'
    end
  end
  object cdsSo112A: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSo112A'
    RemoteServer = DCOMConnection1
    AfterScroll = cdsSo112AAfterScroll
    Left = 452
    Top = 480
  end
end
