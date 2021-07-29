object Emc_Separate1: TEmc_Separate1
  OldCreateOrder = False
  OnCreate = RemoteDataModuleCreate
  OnDestroy = RemoteDataModuleDestroy
  Left = 6
  Top = 31
  Height = 502
  Width = 785
  object dspSO110: TDataSetProvider
    DataSet = qrySO110
    Constraints = True
    Options = [poAllowCommandText]
    Left = 66
    Top = 25
  end
  object dspSO111D: TDataSetProvider
    DataSet = qrySO111D
    Constraints = True
    Options = [poAllowCommandText]
    Left = 136
    Top = 90
  end
  object dspSO112M: TDataSetProvider
    DataSet = qrySO112M
    Constraints = True
    Options = [poAllowCommandText]
    Left = 56
    Top = 91
  end
  object dspCd024CT: TDataSetProvider
    DataSet = qryCd024CT
    Constraints = True
    Options = [poAllowCommandText]
    Left = 229
    Top = 97
  end
  object dspSqlStatement: TDataSetProvider
    DataSet = qrySqlStatement
    Constraints = True
    Options = [poAllowCommandText]
    Left = 357
    Top = 23
  end
  object dspSo113CT: TDataSetProvider
    DataSet = qrySo113CT
    Constraints = True
    Options = [poAllowCommandText]
    Left = 257
    Top = 24
  end
  object dspSqlDelete: TDataSetProvider
    DataSet = qrySqlDelete
    Constraints = True
    Options = [poAllowCommandText]
    Left = 371
    Top = 80
  end
  object dspSo113: TDataSetProvider
    DataSet = qrySO113
    Constraints = True
    Options = [poAllowCommandText]
    Left = 50
    Top = 160
  end
  object dspCD039: TDataSetProvider
    DataSet = qryCD039
    Constraints = True
    Options = [poAllowCommandText]
    Left = 133
    Top = 159
  end
  object dspSo114N: TDataSetProvider
    DataSet = qrySo114N
    Constraints = True
    Left = 216
    Top = 159
  end
  object dspCodeCd039: TDataSetProvider
    DataSet = qryCodeCd039
    Constraints = True
    Left = 318
    Top = 159
  end
  object dspParam: TDataSetProvider
    DataSet = qryParam
    Constraints = True
    Options = [poAllowCommandText]
    Left = 686
    Top = 23
  end
  object dspCd019: TDataSetProvider
    DataSet = qryCd019
    Constraints = True
    Left = 318
    Top = 93
  end
  object dspChargeInfoForTally: TDataSetProvider
    DataSet = qryChargeInfoForTally
    Constraints = True
    Options = [poAllowMultiRecordUpdates, poAllowCommandText]
    Left = 478
    Top = 24
  end
  object dspPackageId: TDataSetProvider
    DataSet = qryPackageId
    Constraints = True
    Options = [poAllowCommandText]
    Left = 497
    Top = 95
  end
  object dspSearSo112: TDataSetProvider
    DataSet = qrySearSo112
    Constraints = True
    Options = [poAllowCommandText]
    Left = 582
    Top = 24
  end
  object qryCodeCD039A: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM CD039')
    Left = 271
    Top = 217
  end
  object dsoCodeCD039A: TDataSetProvider
    DataSet = qryCodeCD039A
    Constraints = True
    Options = [poAllowCommandText]
    Left = 303
    Top = 233
  end
  object ADOConnection1: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 31
    Top = 224
  end
  object qryCom: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 92
    Top = 219
  end
  object dspCom: TDataSetProvider
    DataSet = qryCom
    Constraints = True
    Options = [poAllowCommandText]
    Left = 127
    Top = 227
  end
  object qryRefData: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      
        'SELECT A.PRODUCT_ID,A.FORMULA_ID,A.REF_NO,B.CITEMCODE FROM SO112' +
        ' A,CD019A B WHERE A.PRODUCT_ID IS NULL')
    Left = 366
    Top = 217
  end
  object dspRefData: TDataSetProvider
    DataSet = qryRefData
    Constraints = True
    Options = [poAllowCommandText]
    Left = 391
    Top = 233
  end
  object qryRefDataA: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      
        'SELECT A.PRODUCT_ID,A.FORMULA_ID,A.REF_NO,B.CITEMCODE FROM SO112' +
        ' A,CD019A B WHERE A.PRODUCT_ID IS NULL')
    Left = 176
    Top = 217
  end
  object dspRefDataA: TDataSetProvider
    DataSet = qryRefDataA
    Constraints = True
    Options = [poAllowCommandText]
    Left = 208
    Top = 233
  end
  object qryChargeData: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO033 WHERE CITEMCODE IS NULL')
    Left = 100
    Top = 281
  end
  object dspChargeData: TDataSetProvider
    DataSet = qryChargeData
    Constraints = True
    Options = [poAllowCommandText]
    Left = 132
    Top = 297
  end
  object qryCustData: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO033 WHERE CITEMCODE IS NULL')
    Left = 15
    Top = 276
  end
  object dspCustData: TDataSetProvider
    DataSet = qryCustData
    Constraints = True
    Options = [poAllowCommandText]
    Left = 47
    Top = 300
  end
  object qrySO115: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO115 WHERE CUSTID IS NULL')
    Left = 280
    Top = 288
  end
  object dspSO115: TDataSetProvider
    DataSet = qrySO115
    Constraints = True
    Options = [poAllowCommandText]
    Left = 312
    Top = 296
  end
  object qrySO114: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO114 WHERE COMP_ID IS NULL')
    Left = 193
    Top = 288
  end
  object dspSO114: TDataSetProvider
    DataSet = qrySO114
    Constraints = True
    Options = [poAllowCommandText]
    Left = 225
    Top = 298
  end
  object qrySO116: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO116 WHERE COMP_ID IS NULL')
    Left = 361
    Top = 288
  end
  object dspSO116: TDataSetProvider
    DataSet = qrySO116
    Constraints = True
    Options = [poAllowCommandText]
    Left = 393
    Top = 296
  end
  object qrySO117: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO117 WHERE COMP_ID IS NULL')
    Left = 446
    Top = 288
  end
  object dspSO117: TDataSetProvider
    DataSet = qrySO117
    Constraints = True
    Options = [poAllowCommandText]
    Left = 478
    Top = 296
  end
  object dspSO110Look: TDataSetProvider
    DataSet = qrySO110Look
    Constraints = True
    Left = 164
    Top = 24
  end
  object dspSo113Look: TDataSetProvider
    DataSet = qrySo113Look
    Constraints = True
    Left = 590
    Top = 96
  end
  object dspCd024Look: TDataSetProvider
    DataSet = qryCd024Look
    Constraints = True
    Left = 686
    Top = 96
  end
  object qrySO110: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT FORMULA_ID,'
      'FORMULA_NAME,FORMULA_DESC,REF_NO,STOPFLAG,OPERATOR,UPDTIME'
      'FROM SO110'
      'ORDER BY FORMULA_ID')
    Left = 26
    Top = 8
  end
  object qrySO112M: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM  SO112'
      'ORDER BY  PRODUCT_ID')
    Left = 24
    Top = 72
  end
  object qryCd024CT: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT CODENO,'
      'DESCRIPTION,'
      'REFNO,PAYFLAG, CITEMCODE,'
      'CITEMNAME,UPDTIME,UPDEN,COMPCODE, STOPFLAG,'
      'CHANCEDAYS,CHANNELID FROM CD024'
      'ORDER BY CODENO')
    Left = 197
    Top = 80
  end
  object qrySo113CT: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT PROVIDER_ID,'
      'PROVIDER_NAME,ATTRIBUTE, CONTACTEE,TEL, ADDRESS,NOTES, '
      'STOPFLAG,OPERATOR,UPDTIME'
      'FROM SO113'
      'ORDER BY PROVIDER_ID')
    Left = 221
    Top = 8
  end
  object qrySqlStatement: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from So111')
    Left = 326
    Top = 8
  end
  object qrySqlDelete: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      '')
    Left = 406
    Top = 91
  end
  object qryChargeInfoForTally: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 438
    Top = 8
  end
  object qrySearSo112: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO112 WHERE (FORMULA_ID='#39'2'#39' '
      'OR FORMULA_ID='#39'4'#39' ) AND REF_NO<>'#39'1'#39)
    Left = 550
    Top = 8
  end
  object qrySO113: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO113')
    Left = 16
    Top = 144
  end
  object qryCD039: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM CD039')
    Left = 96
    Top = 144
  end
  object qrySo114N: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select * from SO114')
    Left = 184
    Top = 144
  end
  object qryCodeCd039: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      
        'SELECT CODENO,CONCAT(CONCAT(CODENO,'#39'-'#39'),DESCRIPTION) CODENO_DESC' +
        'RIPTION FROM CD039')
    Left = 280
    Top = 144
  end
  object qryParam: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select count(*) from so001 where CustStatusCode=1')
    Left = 648
    Top = 8
  end
  object qryCd019: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select * from cd019')
    Left = 283
    Top = 78
  end
  object qryPackageId: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 465
    Top = 80
  end
  object qrySo113Look: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT PROVIDER_ID,'
      
        'CONCAT(CONCAT(PROVIDER_ID,'#39'- -'#39'),PROVIDER_NAME) PROVIDER_ID_NAME' +
        ','
      'PROVIDER_NAME FROM SO113 WHERE STOPFLAG=0')
    Left = 558
    Top = 80
  end
  object qryCd024Look: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      
        'SELECT CODENO,CONCAT(CONCAT(CODENO,'#39'-'#39'), DESCRIPTION) CODENO_DES' +
        'CRIPTION,'
      'DESCRIPTION FROM CD024 WHERE STOPFLAG=0'
      '')
    Left = 646
    Top = 80
  end
  object qrySO115B: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO115 WHERE CUSTID IS NULL')
    Left = 533
    Top = 288
  end
  object dspSO115B: TDataSetProvider
    DataSet = qrySO115B
    Constraints = True
    Options = [poAllowCommandText]
    Left = 565
    Top = 296
  end
  object qryPriv: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT MID FROM SO029 WHERE COMPCODE IS NULL')
    Left = 614
    Top = 288
  end
  object cdsPriv: TDataSetProvider
    DataSet = qryPriv
    Constraints = True
    Options = [poAllowCommandText]
    Left = 646
    Top = 304
  end
  object qrySO111D: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select * from SO111')
    Left = 104
    Top = 72
  end
  object qrySO110Look: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      
        'SELECT FORMULA_ID,CONCAT(CONCAT(FORMULA_ID, '#39'-'#39'), FORMULA_NAME) ' +
        ' FORMULA_ID_NAME,'
      'FORMULA_NAME FROM SO110')
    Left = 130
    Top = 8
  end
  object qrySO114B: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO114 WHERE COMP_ID IS NULL')
    Left = 690
    Top = 288
  end
  object dspSO114B: TDataSetProvider
    DataSet = qrySO114B
    Constraints = True
    Options = [poAllowCommandText]
    Left = 722
    Top = 296
  end
  object qryTallyRefData1: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO111 WHERE PRODUCT_ID IS NULL')
    Left = 32
    Top = 360
  end
  object qryTallyRefData2: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO111 WHERE PRODUCT_ID IS NULL')
    Left = 126
    Top = 360
  end
  object dspTallyRefData1: TDataSetProvider
    DataSet = qryTallyRefData1
    Constraints = True
    Options = [poAllowCommandText]
    Left = 64
    Top = 376
  end
  object dspTallyRefData2: TDataSetProvider
    DataSet = qryTallyRefData2
    Constraints = True
    Options = [poAllowCommandText]
    Left = 158
    Top = 375
  end
  object qrySO119: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO119 WHERE COMP_ID IS NULL'
      '')
    Left = 280
    Top = 360
  end
  object dspSO119: TDataSetProvider
    DataSet = qrySO119
    Constraints = True
    Options = [poAllowCommandText]
    Left = 312
    Top = 368
  end
  object qryTempSO119: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO119')
    Left = 392
    Top = 360
  end
  object dspTempSO119: TDataSetProvider
    DataSet = qryTempSO119
    Constraints = True
    Options = [poAllowCommandText]
    Left = 424
    Top = 368
  end
  object qrySO114C: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO114')
    Left = 484
    Top = 353
  end
  object dspSO114C: TDataSetProvider
    DataSet = qrySO114C
    Constraints = True
    Options = [poAllowCommandText]
    Left = 516
    Top = 361
  end
  object qryCodeCD019: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT CODENO,DESCRIPTION FROM CD019')
    Left = 408
    Top = 161
    object qryCodeCD019CODENO: TIntegerField
      DisplayLabel = #20195#30908
      FieldName = 'CODENO'
    end
    object qryCodeCD019DESCRIPTION: TStringField
      DisplayLabel = #21517#31281
      FieldName = 'DESCRIPTION'
    end
  end
  object dspCodeCD019: TDataSetProvider
    DataSet = qryCodeCD019
    Constraints = True
    Options = [poAllowCommandText]
    Left = 448
    Top = 176
  end
  object qrySO033: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO033 WHERE BillNo IS NULL')
    Left = 496
    Top = 160
  end
  object dspSO033: TDataSetProvider
    DataSet = qrySO033
    Constraints = True
    Options = [poAllowCommandText]
    Left = 528
    Top = 176
  end
  object qrySO131: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO131 WHERE COMPUTEYM IS NULL')
    Left = 568
    Top = 160
  end
  object dspSO131: TDataSetProvider
    DataSet = qrySO131
    Constraints = True
    Options = [poAllowCommandText]
    Left = 604
    Top = 168
  end
  object spSO131: TADOStoredProc
    Connection = ADOConnection1
    ProcedureName = 'SF_CALCULATESO131'
    Parameters = <
      item
        Name = 'RETURN_VALUE'
        DataType = ftBCD
        Direction = pdReturnValue
        Value = Null
      end
      item
        Name = 'P_COMPCODE'
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'P_SERVICETYPE'
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'P_COMPUTEYM'
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'P_STARTDATE'
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'P_ENDDATE'
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'P_CHARGEITEMSQL'
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'P_REALORSHOULDDATE'
        DataType = ftBCD
        Value = Null
      end
      item
        Name = 'P_RETCODE'
        DataType = ftBCD
        Direction = pdOutput
        Value = Null
      end
      item
        Name = 'P_RETMSG'
        DataType = ftString
        Direction = pdOutput
        Size = 4000
        Value = Null
      end>
    Left = 544
    Top = 224
  end
  object dspSpSO131: TDataSetProvider
    DataSet = spSO131
    Constraints = True
    Options = [poAllowCommandText]
    Left = 576
    Top = 232
  end
  object qryCd019A: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        'select CODENO,DESCRIPTION, concat(concat(CODENO, '#39'-'#39'),DESCRIPTIO' +
        'N) REFDATA from CD019 where REFNO=2')
    Left = 656
    Top = 160
  end
  object dspCd019A: TDataSetProvider
    DataSet = qryCd019A
    Constraints = True
    Left = 688
    Top = 176
  end
  object qrySo112A: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from SO112A order by CODENO')
    Left = 664
    Top = 224
  end
  object dspSo112A: TDataSetProvider
    DataSet = qrySo112A
    Constraints = True
    Left = 696
    Top = 232
  end
  object qrySO131Excel: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      
        'SELECT ComputeYM,RealOrShouldDate,SeqNo,BillNo,Item,CustId,Citem' +
        'Code,CitemName,ShouldDate,TO_CHAR(ShouldAmt) ShouldAmt,RealDate,' +
        'TO_CHAR(RealAmt) RealAmt,RealPeriod,RealStartDate,RealStopDate,C' +
        'ompCode,OrderNo,SBillNo,SItem,MediaCode,MediaName,AcceptEn,Accep' +
        'tName,IntroId,IntroName,Notes,RefNo FROM SO131 WHERE ComputeYM I' +
        'S NULL')
    Left = 688
    Top = 344
  end
  object dspSO131Excel: TDataSetProvider
    DataSet = qrySO131Excel
    Constraints = True
    Options = [poAllowCommandText]
    Left = 720
    Top = 360
  end
  object qryOtherSO131Excel: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      
        'SELECT ComputeYM,RealOrShouldDate,SeqNo,BillNo,Item,CustId,Citem' +
        'Code,CitemName,ShouldDate,TO_CHAR(ShouldAmt) ShouldAmt,RealDate,' +
        'TO_CHAR(RealAmt) RealAmt,RealPeriod,RealStartDate,RealStopDate,C' +
        'ompCode,OrderNo,SBillNo,SItem,MediaCode,MediaName,AcceptEn,Accep' +
        'tName,IntroId,IntroName,Notes,RefNo FROM SO131 WHERE ComputeYM I' +
        'S NULL')
    Left = 632
    Top = 392
  end
  object dspOtherSO131Excel: TDataSetProvider
    DataSet = qryOtherSO131Excel
    Constraints = True
    Options = [poAllowCommandText]
    Left = 664
    Top = 408
  end
end
