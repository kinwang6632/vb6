object dtmSO: TdtmSO
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 347
  Top = 215
  Height = 421
  Width = 493
  object SOConnection: TADOConnection
    LoginPrompt = False
    Left = 72
    Top = 20
  end
  object adoComm: TADOQuery
    CacheSize = 1000
    Connection = SOConnection
    Parameters = <>
    Left = 72
    Top = 80
  end
  object adoQrySynSO041: TADOQuery
    Connection = SOConnection
    Parameters = <>
    Left = 172
    Top = 20
  end
  object adoQrySynCD019: TADOQuery
    Connection = SOConnection
    Parameters = <>
    Left = 172
    Top = 76
  end
  object adoQrySynCD032: TADOQuery
    Connection = SOConnection
    Parameters = <>
    Left = 172
    Top = 133
  end
  object adoQrySynCD046: TADOQuery
    Connection = SOConnection
    Parameters = <>
    Left = 172
    Top = 187
  end
  object adoQrySynSO001: TADOQuery
    Connection = SOConnection
    Parameters = <>
    Left = 172
    Top = 244
  end
  object adoQrySynSO002: TADOQuery
    Connection = SOConnection
    Parameters = <>
    Left = 172
    Top = 299
  end
  object adoMdCustomer: TADOQuery
    CacheSize = 100
    Connection = SOConnection
    Parameters = <>
    Prepared = True
    Left = 72
    Top = 148
  end
  object adoChargeInfo34: TADOQuery
    CacheSize = 100
    Connection = SOConnection
    CursorType = ctStatic
    Parameters = <>
    Prepared = True
    SQL.Strings = (
      ''
      'SELECT 34 SOURCE, '
      '       A.COMPCODE, '
      '       A.CUSTID, '
      '       A.BILLNO, '
      '       A.ITEM,'
      '       A.CITEMCODE, '
      '       A.CITEMNAME, '
      '       A.SHOULDDATE,'
      '       A.SHOULDAMT,'
      '       A.REALDATE,'
      '       A.REALAMT,'
      '       A.REALPERIOD,'
      '       A.REALSTARTDATE,'
      '       A.REALSTOPDATE,'
      '       A.CLCTNAME,'
      '       A.STNAME,'
      '       A.ACCOUNTNO,'
      '       A.SERVICETYPE,'
      '       C.RATE1, '
      '       B.TAXCODE'
      '  FROM SO034 A, CD019 B, CD033 C'
      ' WHERE A.CITEMCODE = B.CODENO '
      '   AND B.TAXCODE = C.CODENO'
      '   AND C.TAXFLAG = 1'
      '   AND A.CANCELFLAG = 0'
      '   AND ( A.SHOULDAMT + A.REALAMT <> 0 )'
      '   AND A.GUINO IS NULL'
      '   AND A.INVOICETIME IS NULL')
    Left = 74
    Top = 212
  end
  object adoChargeInfo33: TADOQuery
    CacheSize = 100
    Connection = SOConnection
    CursorType = ctStatic
    Parameters = <>
    Prepared = True
    SQL.Strings = (
      ''
      'SELECT 33 SOURCE, '
      '       A.COMPCODE, '
      '       A.CUSTID, '
      '       A.BILLNO, '
      '       A.ITEM,'
      '       A.CITEMCODE, '
      '       A.CITEMNAME, '
      '       A.SHOULDDATE,'
      '       A.SHOULDAMT,'
      '       A.REALDATE,'
      '       A.REALAMT,'
      '       A.REALPERIOD,'
      '       A.REALSTARTDATE,'
      '       A.REALSTOPDATE,'
      '       A.CLCTNAME,'
      '       A.STNAME,'
      '       A.ACCOUNTNO,'
      '       A.SERVICETYPE,'
      '       C.RATE1, '
      '       B.TAXCODE'
      '  FROM SO033 A, CD019 B, CD033 C'
      ' WHERE A.CITEMCODE = B.CODENO '
      '   AND B.TAXCODE = C.CODENO'
      '   AND C.TAXFLAG = 1'
      '   AND A.CANCELFLAG = 0'
      '   AND ( A.SHOULDAMT + A.REALAMT <> 0 )'
      '   AND A.GUINO IS NULL'
      '   AND A.INVOICETIME IS NULL')
    Left = 76
    Top = 268
  end
  object adoA08: TADOQuery
    Connection = SOConnection
    Parameters = <>
    Left = 76
    Top = 328
  end
  object adoQrySynCD095: TADOQuery
    Connection = SOConnection
    Parameters = <>
    Left = 268
    Top = 19
  end
end
