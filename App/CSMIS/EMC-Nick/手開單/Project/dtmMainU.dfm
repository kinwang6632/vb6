object dtmMain: TdtmMain
  OldCreateOrder = False
  Left = 746
  Top = 169
  Height = 507
  Width = 215
  object adoConnection: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 48
    Top = 16
  end
  object qrySo126: TADOQuery
    Connection = adoConnection
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from SO126 where SEQ=-99')
    Left = 48
    Top = 80
    object qrySo126SEQ: TBCDField
      FieldName = 'SEQ'
      Precision = 10
      Size = 0
    end
    object qrySo126COMPCODE: TIntegerField
      FieldName = 'COMPCODE'
    end
    object qrySo126EMPNO: TStringField
      FieldName = 'EMPNO'
    end
    object qrySo126EMPNAME: TStringField
      DisplayLabel = #38936#21934#20154#21729
      FieldName = 'EMPNAME'
      Size = 30
    end
    object qrySo126GETPAPERDATE: TDateTimeField
      DisplayLabel = #38936#21934#26085#26399
      FieldName = 'GETPAPERDATE'
      OnGetText = qrySo126GETPAPERDATEGetText
    end
    object qrySo126PREFIX: TStringField
      FieldName = 'PREFIX'
      Size = 5
    end
    object qrySo126BEGINNUM: TStringField
      DisplayLabel = #36215#22987#21934#34399
      FieldName = 'BEGINNUM'
      Size = 10
    end
    object qrySo126ENDNUM: TStringField
      DisplayLabel = #25130#27490#21934#34399
      FieldName = 'ENDNUM'
      Size = 10
    end
    object qrySo126TOTALPAPERCOUNT: TIntegerField
      DisplayLabel = #21512#35336#24373#25976
      FieldName = 'TOTALPAPERCOUNT'
    end
    object qrySo126OPERATOR: TStringField
      DisplayLabel = #30064#21205#20154#21729
      FieldName = 'OPERATOR'
    end
    object qrySo126UPDTIME: TStringField
      DisplayLabel = #30064#21205#26178#38291
      FieldName = 'UPDTIME'
    end
    object qrySo126RETURNDATE: TDateTimeField
      DisplayLabel = #36864#22238#26085#26399
      FieldName = 'RETURNDATE'
    end
    object qrySo126CLEARDATE: TDateTimeField
      DisplayLabel = #28165#26597#26085#26399
      FieldName = 'CLEARDATE'
    end
    object qrySo126NOTE: TStringField
      DisplayLabel = #20633#35387
      DisplayWidth = 200
      FieldName = 'NOTE'
      Size = 2000
    end
  end
  object qryCm003: TADOQuery
    Connection = adoConnection
    Parameters = <>
    SQL.Strings = (
      'SELECT EMPNO, EMPNAME FROM CM003')
    Left = 120
    Top = 80
    object qryCm003EMPNO: TStringField
      FieldName = 'EMPNO'
      Size = 10
    end
    object qryCm003EMPNAME: TStringField
      FieldName = 'EMPNAME'
    end
  end
  object qryCommon: TADOQuery
    Connection = adoConnection
    Parameters = <>
    Left = 120
    Top = 16
  end
  object cdsSo126: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSo126'
    BeforePost = cdsSo126BeforePost
    Left = 52
    Top = 192
    object cdsSo126SEQ: TBCDField
      FieldName = 'SEQ'
      Precision = 10
      Size = 0
    end
    object cdsSo126COMPCODE: TIntegerField
      FieldName = 'COMPCODE'
    end
    object cdsSo126EMPNO: TStringField
      FieldName = 'EMPNO'
    end
    object cdsSo126EMPNAME: TStringField
      DisplayLabel = #38936#21934#20154#21729
      FieldName = 'EMPNAME'
      Size = 30
    end
    object cdsSo126GETPAPERDATE: TDateTimeField
      DisplayLabel = #38936#21934#26085#26399
      FieldName = 'GETPAPERDATE'
      OnGetText = cdsSo126GETPAPERDATEGetText
    end
    object cdsSo126PREFIX: TStringField
      FieldName = 'PREFIX'
      Size = 5
    end
    object cdsSo126BEGINNUM: TStringField
      DisplayLabel = #36215#22987#21934#34399
      FieldName = 'BEGINNUM'
      Size = 10
    end
    object cdsSo126ENDNUM: TStringField
      DisplayLabel = #25130#27490#21934#34399
      FieldName = 'ENDNUM'
      Size = 10
    end
    object cdsSo126TOTALPAPERCOUNT: TIntegerField
      DisplayLabel = #21512#35336#24373#25976
      FieldName = 'TOTALPAPERCOUNT'
    end
    object cdsSo126OPERATOR: TStringField
      DisplayLabel = #30064#21205#20154#21729
      FieldName = 'OPERATOR'
    end
    object cdsSo126RETURNDATE: TDateTimeField
      DisplayLabel = #36864#22238#26085#26399
      FieldName = 'RETURNDATE'
      OnGetText = cdsSo126RETURNDATEGetText
    end
    object cdsSo126CLEARDATE: TDateTimeField
      DisplayLabel = #28165#26597#26085#26399
      FieldName = 'CLEARDATE'
      OnGetText = cdsSo126CLEARDATEGetText
    end
    object cdsSo126NOTE: TStringField
      DisplayLabel = #20633#35387
      FieldName = 'NOTE'
      Size = 2000
    end
    object cdsSo126UPDTIME: TStringField
      DisplayLabel = #30064#21205#26178#38291
      FieldName = 'UPDTIME'
    end
  end
  object dspSo126: TDataSetProvider
    DataSet = qrySo126
    Options = [poAllowCommandText, poRetainServerOrder]
    Left = 48
    Top = 136
  end
  object qryReport1: TADOQuery
    Connection = adoConnection
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        'select EMPNO,EMPNAME,GETPAPERDATE,COUNT(SEQ) COUNT, MIN(PAPERNUM' +
        ') MIN_PAPER_NUM, MAX(PAPERNUM) MAX_PAPER_NUM from SO127  group b' +
        'y EMPNO, EMPNAME,GETPAPERDATE')
    Left = 120
    Top = 136
  end
  object qryReport2: TADOQuery
    Connection = adoConnection
    Parameters = <>
    SQL.Strings = (
      'select EMPNO,EMPNAME,GETPAPERDATE,PAPERNUM from SO127'
      'where status=1')
    Left = 120
    Top = 192
    object qryReport2EMPNO: TStringField
      FieldName = 'EMPNO'
    end
    object qryReport2EMPNAME: TStringField
      FieldName = 'EMPNAME'
      Size = 30
    end
    object qryReport2GETPAPERDATE: TDateTimeField
      FieldName = 'GETPAPERDATE'
    end
    object qryReport2PAPERNUM: TStringField
      FieldName = 'PAPERNUM'
      Size = 15
    end
  end
  object qryReport3: TADOQuery
    Connection = adoConnection
    Parameters = <>
    SQL.Strings = (
      'select EMPNO,EMPNAME,GETPAPERDATE,PAPERNUM, BILLNO,  '
      
        'STATUS,CUSTID,CUSTNAME,CUSTTEL,REALDATE  from SO127  where (STAT' +
        'US=0 or (BILLNO<>'#39#39' or BILLNO IS NOT NULL))  and CompCode=1 orde' +
        'r by EMPNO, GETPAPERDATE, PAPERNUM')
    Left = 120
    Top = 256
    object qryReport3EMPNO: TStringField
      FieldName = 'EMPNO'
    end
    object qryReport3EMPNAME: TStringField
      FieldName = 'EMPNAME'
      Size = 30
    end
    object qryReport3GETPAPERDATE: TDateTimeField
      FieldName = 'GETPAPERDATE'
    end
    object qryReport3PAPERNUM: TStringField
      FieldName = 'PAPERNUM'
      Size = 15
    end
    object qryReport3BILLNO: TStringField
      FieldName = 'BILLNO'
      Size = 15
    end
    object qryReport3CUSTID: TIntegerField
      FieldName = 'CUSTID'
    end
    object qryReport3CUSTNAME: TStringField
      FieldName = 'CUSTNAME'
    end
    object qryReport3CUSTTEL: TStringField
      FieldName = 'CUSTTEL'
    end
    object qryReport3REALDATE: TDateTimeField
      FieldName = 'REALDATE'
    end
    object qryReport3STATUS: TIntegerField
      FieldName = 'STATUS'
    end
  end
  object cdsReport2: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 120
    Top = 320
    object cdsReport2EmpNo: TStringField
      FieldName = 'EmpNo'
    end
    object cdsReport2EmpName: TStringField
      FieldName = 'EmpName'
    end
    object cdsReport2GetPaperDate: TStringField
      FieldName = 'GetPaperDate'
    end
    object cdsReport2PaperNumbers: TStringField
      FieldName = 'PaperNumbers'
      Size = 100
    end
    object cdsReport2GroupID: TIntegerField
      FieldName = 'GroupID'
    end
  end
  object qryReport2Common: TADOQuery
    Connection = adoConnection
    Parameters = <>
    SQL.Strings = (
      
        'select EMPNO, TO_CHAR(GETPAPERDATE,'#39'YYYYMMDD'#39') GETPAPERDATE, COU' +
        'NT(PAPERNUM) COUNTS from SO127  '
      'where status=1 group by EMPNO, GETPAPERDATE ')
    Left = 56
    Top = 248
    object qryReport2CommonEMPNO: TStringField
      FieldName = 'EMPNO'
    end
    object qryReport2CommonGETPAPERDATE: TStringField
      FieldName = 'GETPAPERDATE'
      ReadOnly = True
      Size = 8
    end
    object qryReport2CommonCOUNTS: TBCDField
      FieldName = 'COUNTS'
      ReadOnly = True
      Precision = 32
      Size = 0
    end
  end
  object qryPriv: TADOQuery
    Connection = adoConnection
    Parameters = <>
    Left = 48
    Top = 312
  end
  object qrySo18C2: TADOQuery
    Connection = adoConnection
    Parameters = <>
    SQL.Strings = (
      
        'select PAPERNUM,EMPNAME,GETPAPERDATE,BILLNO from SO127  where PA' +
        'PERNUM between to_number(0000000001) and  to_number(0000000010) ' +
        'and STATUS<>0')
    Left = 120
    Top = 376
    object qrySo18C2PAPERNUM: TStringField
      FieldName = 'PAPERNUM'
      Size = 15
    end
    object qrySo18C2EMPNAME: TStringField
      FieldName = 'EMPNAME'
      Size = 30
    end
    object qrySo18C2GETPAPERDATE: TDateTimeField
      FieldName = 'GETPAPERDATE'
      OnGetText = qrySo18C2GETPAPERDATEGetText
    end
    object qrySo18C2BILLNO: TStringField
      DisplayWidth = 20
      FieldName = 'BILLNO'
      Size = 15
    end
  end
  object qrySo034: TADOQuery
    Connection = adoConnection
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        'select a.CustId,b.CustName, b.Tel1, a.BillNo,a.ManualNo,a.Item,a' +
        '.CitemName,a.OldAmt,a.ShouldAmt,a.RealDate, a.RealAmt, a.ClctNam' +
        'e,'
      
        'a.CMName,a.UCName,a.STName from SO034 a, SO001 b where a.CompCod' +
        'e=-99'
      'and a.CustID=b.CustID and a.CompCode=b.CompCode')
    Left = 48
    Top = 376
    object qrySo034CUSTID: TIntegerField
      DisplayLabel = #23458#25142#32232#34399
      FieldName = 'CUSTID'
    end
    object qrySo034CUSTNAME: TStringField
      DisplayLabel = #23458#25142#22995#21517
      DisplayWidth = 15
      FieldName = 'CUSTNAME'
      Size = 50
    end
    object qrySo034TEL1: TStringField
      DisplayLabel = #38651#35441
      DisplayWidth = 15
      FieldName = 'TEL1'
    end
    object qrySo034BILLNO: TStringField
      DisplayLabel = #21934#25818#32232#34399
      FieldName = 'BILLNO'
      Size = 15
    end
    object qrySo034MANUALNO: TStringField
      DisplayLabel = #25163#38283#21934#34399
      FieldName = 'MANUALNO'
      Size = 10
    end
    object qrySo034ITEM: TIntegerField
      DisplayLabel = #38917#27425
      FieldName = 'ITEM'
    end
    object qrySo034CITEMNAME: TStringField
      DisplayLabel = #25910#36027#38917#30446#21517#31281
      FieldName = 'CITEMNAME'
    end
    object qrySo034OLDAMT: TIntegerField
      DisplayLabel = #21407#25033#25910#37329#38989
      FieldName = 'OLDAMT'
    end
    object qrySo034SHOULDAMT: TIntegerField
      DisplayLabel = #20986#21934#37329#38989
      FieldName = 'SHOULDAMT'
    end
    object qrySo034REALDATE: TDateTimeField
      DisplayLabel = #23526#25910#26085#26399
      FieldName = 'REALDATE'
    end
    object qrySo034REALAMT: TIntegerField
      DisplayLabel = #23526#25910#37329#38989
      FieldName = 'REALAMT'
    end
    object qrySo034CLCTNAME: TStringField
      DisplayLabel = #25910#36027#20154#21729#22995#21517
      FieldName = 'CLCTNAME'
    end
    object qrySo034CMNAME: TStringField
      DisplayLabel = #25910#36027#26041#24335
      FieldName = 'CMNAME'
    end
    object qrySo034UCNAME: TStringField
      DisplayLabel = #26410#25910#21407#22240
      FieldName = 'UCNAME'
    end
    object qrySo034STNAME: TStringField
      DisplayLabel = #30701#25910#21407#22240
      FieldName = 'STNAME'
    end
  end
  object qryReport4: TADOQuery
    Connection = adoConnection
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from SO128 where COMPCODE=-99')
    Left = 50
    Top = 424
    object qryReport4COMPCODE: TIntegerField
      FieldName = 'COMPCODE'
    end
    object qryReport4PAPERNUM: TStringField
      FieldName = 'PAPERNUM'
      Size = 15
    end
    object qryReport4UPDATESTATUS: TStringField
      FieldName = 'UPDATESTATUS'
      Size = 4000
    end
    object qryReport4OPERATOR: TStringField
      FieldName = 'OPERATOR'
    end
    object qryReport4UPDTIME: TDateTimeField
      FieldName = 'UPDTIME'
    end
  end
  object cdsReport4: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 120
    Top = 424
    object cdsReport4PaperNum: TStringField
      FieldName = 'PaperNum'
      Size = 15
    end
    object cdsReport4FieldName: TStringField
      FieldName = 'FieldName'
    end
    object cdsReport4OldValue: TStringField
      FieldName = 'OldValue'
    end
    object cdsReport4NewValue: TStringField
      FieldName = 'NewValue'
    end
    object cdsReport4Operator: TStringField
      FieldName = 'Operator'
    end
    object cdsReport4UpdTime: TStringField
      FieldName = 'UpdTime'
    end
    object cdsReport4Group: TIntegerField
      FieldName = 'Group'
    end
  end
end
