object dtmMain1: TdtmMain1
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 376
  Top = 157
  Height = 580
  Width = 783
  object qryCodeCD039: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT CODENO,DESCRIPTION FROM CD039')
    Left = 201
    Top = 121
  end
  object qryCodeCD019: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT CODENO,DESCRIPTION FROM CD019')
    Left = 200
    Top = 177
    object qryCodeCD019CODENO: TIntegerField
      DisplayLabel = #20195#30908
      FieldName = 'CODENO'
    end
    object qryCodeCD019DESCRIPTION: TStringField
      DisplayLabel = #21517#31281
      FieldName = 'DESCRIPTION'
    end
  end
  object qryCodeCD032: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT CODENO,DESCRIPTION FROM CD032')
    Left = 120
    Top = 169
  end
  object qryCom: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    Left = 129
    Top = 17
  end
  object qrySO122: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO122'
      'where COMPCODE is NULL')
    Left = 336
    Top = 9
  end
  object qrySo033So034: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      
        'SELECT A.RefNo,B.CustId,B.OrderNo,B.FirstFlag,B.DiscountCode,B.D' +
        'iscountName,B.OrderNo,TO_DATE(TO_CHAR(B.RealStartDate,'#39'YYYY/MM/D' +
        'D'#39'),'#39'YYYY/MM/DD'#39') RealStartDate ,TO_DATE(TO_CHAR(B.RealStopDate,' +
        #39'YYYY/MM/DD'#39'),'#39'YYYY/MM/DD'#39') RealStopDate ,B.CMName,B.CitemCode,B' +
        '.CitemName,B.PTCode,B.CompCode,B.ClctEn,B.ClctName,B.BillNo,B.It' +
        'em, B.RealPeriod, B.SBILLNO, B.SITEM ,TO_DATE(TO_CHAR(B.RealDate' +
        ','#39'YYYY/MM/DD'#39'),'#39'YYYY/MM/DD'#39') RealDate,TO_DATE(TO_CHAR(B.ShouldDa' +
        'te,'#39'YYYY/MM/DD'#39'),'#39'YYYY/MM/DD'#39') ShouldDate,B.RealAmt,TO_CHAR(B.Re' +
        'alAmt) StrRealAmt FROM CD019 A, So033 B Where BillNO is NULL')
    Left = 26
    Top = 235
  end
  object qrySo120Formula: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    Left = 105
    Top = 65
  end
  object qrySO120: TADOQuery
    AutoCalcFields = False
    Connection = frmMainMenu.ADOConnection1
    CursorType = ctStatic
    OnCalcFields = qrySO120CalcFields
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO120')
    Left = 185
    Top = 225
    object qrySO120COMPCODE: TIntegerField
      FieldName = 'COMPCODE'
      Visible = False
    end
    object qrySO120CompName: TStringField
      DisplayLabel = #20844#21496#21029
      FieldKind = fkLookup
      FieldName = 'CompName'
      LookupDataSet = qryCodeCD039
      LookupKeyFields = 'CODENO'
      LookupResultField = 'DESCRIPTION'
      KeyFields = 'COMPCODE'
      Lookup = True
    end
    object qrySO120CODENO: TIntegerField
      FieldName = 'CODENO'
      Visible = False
    end
    object qrySO120DESCRIPTION: TStringField
      DisplayLabel = #25910#36027#38917#30446#21517#31281
      FieldName = 'DESCRIPTION'
    end
    object qrySO120PROMOTECODE: TBCDField
      FieldName = 'PROMOTECODE'
      Visible = False
      Precision = 10
      Size = 0
    end
    object qrySO120PROMOTENAME: TStringField
      DisplayLabel = #26989#21209#27963#21205#21517#31281
      FieldName = 'PROMOTENAME'
    end
    object qrySO120DISCOUNTCODE: TBCDField
      FieldName = 'DISCOUNTCODE'
      Visible = False
      Precision = 10
      Size = 0
    end
    object qrySO120DISCOUNTNAME: TStringField
      DisplayLabel = #20778#24800#36774#27861#21517#31281
      FieldName = 'DISCOUNTNAME'
    end
    object qrySO120PAYUNIT: TStringField
      FieldName = 'PAYUNIT'
      Visible = False
      FixedChar = True
      Size = 1
    end
    object qrySO120PAYUNITNAME: TStringField
      DisplayLabel = #20323#37329#21934#20301
      DisplayWidth = 8
      FieldKind = fkCalculated
      FieldName = 'PAYUNITNAME'
      Calculated = True
    end
    object qrySO120FIRSTCREDITCARD1: TBCDField
      DisplayLabel = #39318#27425#20449#29992#21345#30334#20998#27604'/'#37329#38989'-'#25512#24291#20154
      FieldName = 'FIRSTCREDITCARD1'
      Precision = 10
      Size = 2
    end
    object qrySO120FIRSTNOTCREDITCARD1: TBCDField
      DisplayLabel = #39318#27425#38750#20449#29992#21345#30334#20998#27604'/'#37329#38989'-'#25512#24291#20154
      FieldName = 'FIRSTNOTCREDITCARD1'
      Precision = 10
      Size = 2
    end
    object qrySO120FIRSTCREDITCARD2: TBCDField
      DisplayLabel = #39318#27425#20449#29992#21345#30334#20998#27604'/'#37329#38989'-'#25910#36027#21729
      FieldName = 'FIRSTCREDITCARD2'
      Precision = 10
      Size = 2
    end
    object qrySO120FIRSTNOTCREDITCARD2: TBCDField
      DisplayLabel = #39318#27425#38750#20449#29992#21345#30334#20998#27604'/'#37329#38989'-'#25910#36027#21729
      FieldName = 'FIRSTNOTCREDITCARD2'
      Precision = 10
      Size = 2
    end
    object qrySO120OTHERCREDITCARD2: TBCDField
      DisplayLabel = #32396#25910#20449#29992#21345#30334#20998#27604'/'#37329#38989'-'#25910#36027#21729
      FieldName = 'OTHERCREDITCARD2'
      Precision = 10
      Size = 2
    end
    object qrySO120OTHERNOTCREDITCARD2: TBCDField
      DisplayLabel = #32396#25910#38750#20449#29992#21345#30334#20998#27604'/'#37329#38989'-'#25910#36027#21729
      FieldName = 'OTHERNOTCREDITCARD2'
      Precision = 10
      Size = 2
    end
    object qrySO120REF_NO: TIntegerField
      DisplayLabel = #21443#32771#34399
      FieldName = 'REF_NO'
    end
    object qrySO120CHANNELVIEWDAYS: TIntegerField
      DisplayLabel = #38971#36947#38656#25910#35222#28415#22810#23569#22825#25165#32102#20323#37329
      FieldName = 'CHANNELVIEWDAYS'
    end
    object qrySO120OPERATOR: TStringField
      DisplayLabel = #30064#21205#20154#21729
      FieldName = 'OPERATOR'
    end
    object qrySO120UPDTIME: TStringField
      DisplayLabel = #30064#21205#26178#38291
      FieldName = 'UPDTIME'
    end
  end
  object qrySO121Z: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select * from So121'
      'where COMPCODE IS NULL'
      'AND  COMPUTEYM IS NULL'
      'AND  BELONGYM IS NULL'
      'AND  EMPID IS NULL'
      '')
    Left = 473
    Top = 9
  end
  object cdsSo121Z: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'COMPCODE'
        DataType = ftInteger
      end
      item
        Name = 'COMPUTEYM'
        DataType = ftString
        Size = 5
      end
      item
        Name = 'BELONGYM'
        DataType = ftString
        Size = 5
      end
      item
        Name = 'EMPID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'EMPNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'COMMISSION'
        DataType = ftBCD
        Precision = 10
        Size = 2
      end
      item
        Name = 'MEDIACODE'
        DataType = ftInteger
      end
      item
        Name = 'MEDIANAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'COMPTYPE'
        Attributes = [faFixed]
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CLANCOMPCODE'
        DataType = ftInteger
      end
      item
        Name = 'CLANCOMPNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'OPERATOR'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'UPDTIME'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    ProviderName = 'dspSo121'
    StoreDefs = True
    OnReconcileError = cdsSo121ZReconcileError
    Left = 489
    Top = 33
    object cdsSo121ZCOMPCODE: TIntegerField
      FieldName = 'COMPCODE'
      ProviderFlags = [pfInUpdate, pfInWhere, pfInKey]
    end
    object cdsSo121ZCOMPUTEYM: TStringField
      FieldName = 'COMPUTEYM'
      ProviderFlags = [pfInUpdate, pfInWhere, pfInKey]
      Size = 5
    end
    object cdsSo121ZBELONGYM: TStringField
      FieldName = 'BELONGYM'
      ProviderFlags = [pfInUpdate, pfInWhere, pfInKey]
      Size = 5
    end
    object cdsSo121ZEMPID: TStringField
      FieldName = 'EMPID'
      ProviderFlags = [pfInUpdate, pfInWhere, pfInKey]
    end
    object cdsSo121ZEMPNAME: TStringField
      FieldName = 'EMPNAME'
      ProviderFlags = []
    end
    object cdsSo121ZCOMMISSION: TBCDField
      FieldName = 'COMMISSION'
      ProviderFlags = []
      Precision = 10
      Size = 2
    end
    object cdsSo121ZMEDIACODE: TIntegerField
      FieldName = 'MEDIACODE'
      ProviderFlags = []
    end
    object cdsSo121ZMEDIANAME: TStringField
      FieldName = 'MEDIANAME'
      ProviderFlags = []
    end
    object cdsSo121ZCOMPTYPE: TStringField
      FieldName = 'COMPTYPE'
      ProviderFlags = []
      FixedChar = True
      Size = 1
    end
    object cdsSo121ZCLANCOMPCODE: TIntegerField
      FieldName = 'CLANCOMPCODE'
      ProviderFlags = []
    end
    object cdsSo121ZCLANCOMPNAME: TStringField
      FieldName = 'CLANCOMPNAME'
      ProviderFlags = []
    end
    object cdsSo121ZOPERATOR: TStringField
      FieldName = 'OPERATOR'
      ProviderFlags = []
    end
    object cdsSo121ZUPDTIME: TStringField
      FieldName = 'UPDTIME'
      ProviderFlags = []
    end
  end
  object cdsSo122: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'COMPCODE'
        DataType = ftInteger
      end
      item
        Name = 'COMPUTEYM'
        DataType = ftString
        Size = 5
      end
      item
        Name = 'BELONGYM'
        DataType = ftString
        Size = 5
      end
      item
        Name = 'STBNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'SNO'
        DataType = ftString
        Size = 15
      end
      item
        Name = 'BILLNO'
        DataType = ftString
        Size = 15
      end
      item
        Name = 'ITEM'
        DataType = ftInteger
      end
      item
        Name = 'PERIODID'
        DataType = ftInteger
      end
      item
        Name = 'WORKEREN1'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'WORKEREN1NAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'WORKEREN1COMM'
        DataType = ftBCD
        Precision = 10
        Size = 2
      end
      item
        Name = 'CLCTEN'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CLCTENNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CLCTENCOMM'
        DataType = ftBCD
        Precision = 10
        Size = 2
      end
      item
        Name = 'CLCTENORIPERCENT'
        DataType = ftBCD
        Precision = 8
        Size = 2
      end
      item
        Name = 'INTACCID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'INTACCNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'INTACCCOMM'
        DataType = ftBCD
        Precision = 10
        Size = 2
      end
      item
        Name = 'INTACCORIPERCENT'
        DataType = ftBCD
        Precision = 8
        Size = 2
      end
      item
        Name = 'STAFFCODE'
        Attributes = [faFixed]
        DataType = ftString
        Size = 1
      end
      item
        Name = 'REALSTARTDATE'
        DataType = ftDateTime
      end
      item
        Name = 'REALSTOPDATE'
        DataType = ftDateTime
      end
      item
        Name = 'OPERATOR'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'UPDTIME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'BUYORRENT'
        DataType = ftInteger
      end
      item
        Name = 'CITEMCODE'
        DataType = ftInteger
      end
      item
        Name = 'CITEMNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'MEDIACODE'
        DataType = ftInteger
      end
      item
        Name = 'MEDIANAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'REALAMT'
        DataType = ftInteger
      end
      item
        Name = 'PROMCODE'
        DataType = ftInteger
      end
      item
        Name = 'PROMNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'REALDATE'
        DataType = ftDateTime
      end
      item
        Name = 'SHOULDDATE'
        DataType = ftDateTime
      end
      item
        Name = 'CUSTID'
        DataType = ftInteger
      end
      item
        Name = 'CUSTNAME'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'FIRSTFLAG'
        Attributes = [faFixed]
        DataType = ftString
        Size = 1
      end
      item
        Name = 'ORDERNO'
        DataType = ftString
        Size = 15
      end
      item
        Name = 'CALLSELF'
        DataType = ftInteger
      end>
    IndexDefs = <>
    Params = <>
    ProviderName = 'dspSo122'
    StoreDefs = True
    Left = 352
    Top = 33
    object cdsSo122COMPCODE: TIntegerField
      FieldName = 'COMPCODE'
    end
    object cdsSo122COMPUTEYM: TStringField
      FieldName = 'COMPUTEYM'
      Size = 5
    end
    object cdsSo122BELONGYM: TStringField
      FieldName = 'BELONGYM'
      Size = 5
    end
    object cdsSo122STBNO: TStringField
      FieldName = 'STBNO'
    end
    object cdsSo122SNO: TStringField
      FieldName = 'SNO'
      Size = 15
    end
    object cdsSo122BILLNO: TStringField
      FieldName = 'BILLNO'
      Size = 15
    end
    object cdsSo122ITEM: TIntegerField
      FieldName = 'ITEM'
    end
    object cdsSo122PERIODID: TIntegerField
      FieldName = 'PERIODID'
    end
    object cdsSo122WORKEREN1: TStringField
      FieldName = 'WORKEREN1'
    end
    object cdsSo122WORKEREN1NAME: TStringField
      FieldName = 'WORKEREN1NAME'
    end
    object cdsSo122WORKEREN1COMM: TBCDField
      FieldName = 'WORKEREN1COMM'
      Precision = 10
      Size = 2
    end
    object cdsSo122CLCTEN: TStringField
      FieldName = 'CLCTEN'
    end
    object cdsSo122CLCTENNAME: TStringField
      FieldName = 'CLCTENNAME'
    end
    object cdsSo122CLCTENCOMM: TBCDField
      FieldName = 'CLCTENCOMM'
      Precision = 10
      Size = 2
    end
    object cdsSo122CLCTENORIPERCENT: TBCDField
      FieldName = 'CLCTENORIPERCENT'
      Precision = 8
      Size = 2
    end
    object cdsSo122INTACCID: TStringField
      FieldName = 'INTACCID'
    end
    object cdsSo122INTACCNAME: TStringField
      FieldName = 'INTACCNAME'
    end
    object cdsSo122INTACCCOMM: TBCDField
      FieldName = 'INTACCCOMM'
      Precision = 10
      Size = 2
    end
    object cdsSo122INTACCORIPERCENT: TBCDField
      FieldName = 'INTACCORIPERCENT'
      Precision = 8
      Size = 2
    end
    object cdsSo122STAFFCODE: TStringField
      FieldName = 'STAFFCODE'
      Size = 1
    end
    object cdsSo122REALSTARTDATE: TDateTimeField
      FieldName = 'REALSTARTDATE'
    end
    object cdsSo122REALSTOPDATE: TDateTimeField
      FieldName = 'REALSTOPDATE'
    end
    object cdsSo122OPERATOR: TStringField
      FieldName = 'OPERATOR'
    end
    object cdsSo122UPDTIME: TStringField
      FieldName = 'UPDTIME'
    end
    object cdsSo122BUYORRENT: TIntegerField
      FieldName = 'BUYORRENT'
    end
    object cdsSo122CITEMCODE: TIntegerField
      FieldName = 'CITEMCODE'
    end
    object cdsSo122CITEMNAME: TStringField
      FieldName = 'CITEMNAME'
    end
    object cdsSo122MEDIACODE: TIntegerField
      FieldName = 'MEDIACODE'
    end
    object cdsSo122MEDIANAME: TStringField
      FieldName = 'MEDIANAME'
    end
    object cdsSo122REALAMT: TIntegerField
      FieldName = 'REALAMT'
    end
    object cdsSo122PROMCODE: TBCDField
      FieldName = 'PROMCODE'
      Precision = 10
      Size = 0
    end
    object cdsSo122PROMNAME: TStringField
      FieldName = 'PROMNAME'
    end
    object cdsSo122REALDATE: TDateTimeField
      FieldName = 'REALDATE'
    end
    object cdsSo122SHOULDDATE: TDateTimeField
      FieldName = 'SHOULDDATE'
    end
    object cdsSo122CUSTID: TIntegerField
      FieldName = 'CUSTID'
    end
    object cdsSo122CUSTNAME: TStringField
      FieldName = 'CUSTNAME'
      Size = 50
    end
    object cdsSo122FIRSTFLAG: TStringField
      FieldName = 'FIRSTFLAG'
      FixedChar = True
      Size = 1
    end
    object cdsSo122ORDERNO: TStringField
      FieldName = 'ORDERNO'
      Size = 15
    end
    object cdsSo122CALLSELF: TIntegerField
      FieldName = 'CALLSELF'
    end
  end
  object qryCommon: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    Left = 201
    Top = 65
  end
  object dspSo121Z: TDataSetProvider
    DataSet = qrySO121Z
    Options = [poAllowCommandText, poRetainServerOrder]
    UpdateMode = upWhereKeyOnly
    Left = 481
    Top = 17
  end
  object dspSo122: TDataSetProvider
    DataSet = qrySO122
    Options = [poAllowCommandText]
    Left = 344
    Top = 17
  end
  object qryCodeCD009: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select CODENO,DESCRIPTION ,REFNO'
      'from CD009'
      'where STOPFLAG<>1')
    Left = 121
    Top = 121
  end
  object qrySo123: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select * from SO123'
      'where COMPCODE IS NULL')
    Left = 546
    Top = 9
  end
  object cdsSo123: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSo123'
    Left = 562
    Top = 33
  end
  object dspSo123: TDataSetProvider
    DataSet = qrySo123
    Left = 554
    Top = 17
  end
  object cdsSo122ReturnMoney: TClientDataSet
    Aggregates = <>
    FieldDefs = <>
    IndexDefs = <>
    Params = <>
    ProviderName = 'dspSo122ReturnMoney'
    StoreDefs = True
    Left = 168
    Top = 312
  end
  object qrySo122ReturnMoney: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select * from SO122 where COMPCODE IS NULL')
    Left = 152
    Top = 280
  end
  object dspSo122ReturnMoney: TDataSetProvider
    DataSet = qrySo122ReturnMoney
    Left = 160
    Top = 296
  end
  object cdsSo121ReturnMoney: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'COMPCODE'
        DataType = ftInteger
      end
      item
        Name = 'COMPUTEYM'
        DataType = ftString
        Size = 5
      end
      item
        Name = 'BELONGYM'
        DataType = ftString
        Size = 5
      end
      item
        Name = 'EMPID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'EMPNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'COMMISSION'
        DataType = ftBCD
        Precision = 8
        Size = 2
      end
      item
        Name = 'OPERATOR'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'UPDTIME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'MEDIACODE'
        DataType = ftInteger
      end
      item
        Name = 'MEDIANAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'COMPTYPE'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'CLANCOMPCODE'
        DataType = ftInteger
      end
      item
        Name = 'CLANCOMPNAME'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 288
    Top = 384
    object cdsSo121ReturnMoneyCOMPCODE: TIntegerField
      FieldName = 'COMPCODE'
    end
    object cdsSo121ReturnMoneyCOMPUTEYM: TStringField
      FieldName = 'COMPUTEYM'
      Size = 5
    end
    object cdsSo121ReturnMoneyBELONGYM: TStringField
      FieldName = 'BELONGYM'
      Size = 5
    end
    object cdsSo121ReturnMoneyEMPID: TStringField
      FieldName = 'EMPID'
    end
    object cdsSo121ReturnMoneyEMPNAME: TStringField
      FieldName = 'EMPNAME'
    end
    object cdsSo121ReturnMoneyCOMMISSION: TBCDField
      FieldName = 'COMMISSION'
      Precision = 8
      Size = 2
    end
    object cdsSo121ReturnMoneyOPERATOR: TStringField
      FieldName = 'OPERATOR'
    end
    object cdsSo121ReturnMoneyUPDTIME: TStringField
      FieldName = 'UPDTIME'
    end
    object cdsSo121ReturnMoneyMEDIACODE: TIntegerField
      FieldName = 'MEDIACODE'
    end
    object cdsSo121ReturnMoneyMEDIANAME: TStringField
      FieldName = 'MEDIANAME'
    end
    object cdsSo121ReturnMoneyCOMPTYPE: TStringField
      FieldName = 'COMPTYPE'
      Size = 1
    end
    object cdsSo121ReturnMoneyCLANCOMPCODE: TIntegerField
      FieldName = 'CLANCOMPCODE'
    end
    object cdsSo121ReturnMoneyCLANCOMPNAME: TStringField
      FieldName = 'CLANCOMPNAME'
    end
  end
  object qryCodeCD042: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM CD042')
    Left = 32
    Top = 80
  end
  object qryCD042: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM CD042 WHERE CODENO IS NULL')
    Left = 24
    Top = 128
  end
  object qryBox: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    Left = 24
    Top = 184
  end
  object cdsSo125: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSo125'
    Left = 424
    Top = 32
  end
  object qrySO125: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select * from so125')
    Left = 408
    Top = 8
  end
  object dspSO125: TDataSetProvider
    DataSet = qrySO125
    Options = [poAllowCommandText]
    Left = 416
    Top = 16
  end
  object qrySO124: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO124')
    Left = 352
    Top = 96
  end
  object qryCM003: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM CM003')
    Left = 272
    Top = 176
  end
  object qrySO013: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO013')
    Left = 280
    Top = 232
  end
  object qryOtherCompCode: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    Left = 360
    Top = 232
  end
  object cdsSo121ReturnMoneyA: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'COMPCODE'
        DataType = ftInteger
      end
      item
        Name = 'COMPUTEYM'
        DataType = ftString
        Size = 5
      end
      item
        Name = 'BELONGYM'
        DataType = ftString
        Size = 5
      end
      item
        Name = 'EMPID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'EMPNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'COMMISSION'
        DataType = ftBCD
        Precision = 8
        Size = 2
      end
      item
        Name = 'OPERATOR'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'UPDTIME'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    ProviderName = 'dspSo121ReturnMoney'
    StoreDefs = True
    Left = 288
    Top = 328
    object IntegerField1: TIntegerField
      FieldName = 'COMPCODE'
    end
    object StringField1: TStringField
      FieldName = 'COMPUTEYM'
      Size = 5
    end
    object StringField2: TStringField
      FieldName = 'BELONGYM'
      Size = 5
    end
    object StringField3: TStringField
      FieldName = 'EMPID'
    end
    object StringField4: TStringField
      FieldName = 'EMPNAME'
    end
    object BCDField1: TBCDField
      FieldName = 'COMMISSION'
      Precision = 8
      Size = 2
    end
    object StringField5: TStringField
      FieldName = 'OPERATOR'
    end
    object StringField6: TStringField
      FieldName = 'UPDTIME'
    end
  end
  object qrySO121: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select * from So121'
      'where COMPCODE IS NULL'
      'AND  COMPUTEYM IS NULL'
      'AND  BELONGYM IS NULL'
      'AND  EMPID IS NULL')
    Left = 248
    Top = 8
  end
  object dspSo121: TDataSetProvider
    DataSet = qrySO121
    Options = [poAllowCommandText]
    Left = 256
    Top = 16
  end
  object cdsSo121: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSo121'
    Left = 265
    Top = 29
    object cdsSo121COMPCODE: TIntegerField
      FieldName = 'COMPCODE'
    end
    object cdsSo121COMPUTEYM: TStringField
      FieldName = 'COMPUTEYM'
      Size = 5
    end
    object cdsSo121BELONGYM: TStringField
      FieldName = 'BELONGYM'
      Size = 5
    end
    object cdsSo121EMPID: TStringField
      FieldName = 'EMPID'
    end
    object cdsSo121EMPNAME: TStringField
      FieldName = 'EMPNAME'
    end
    object cdsSo121COMMISSION: TBCDField
      FieldName = 'COMMISSION'
      Precision = 10
      Size = 2
    end
    object cdsSo121MEDIACODE: TIntegerField
      FieldName = 'MEDIACODE'
    end
    object cdsSo121MEDIANAME: TStringField
      FieldName = 'MEDIANAME'
    end
    object cdsSo121COMPTYPE: TStringField
      FieldName = 'COMPTYPE'
      FixedChar = True
      Size = 1
    end
    object cdsSo121CLANCOMPCODE: TIntegerField
      FieldName = 'CLANCOMPCODE'
    end
    object cdsSo121CLANCOMPNAME: TStringField
      FieldName = 'CLANCOMPNAME'
    end
    object cdsSo121OPERATOR: TStringField
      FieldName = 'OPERATOR'
    end
    object cdsSo121UPDTIME: TStringField
      FieldName = 'UPDTIME'
    end
  end
  object ADOQuery1: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select * from So121'
      'where COMPCODE IS NULL'
      'AND  COMPUTEYM IS NULL'
      'AND  BELONGYM IS NULL'
      'AND  EMPID IS NULL')
    Left = 392
    Top = 96
    object ADOQuery1COMPCODE: TIntegerField
      FieldName = 'COMPCODE'
    end
    object ADOQuery1COMPUTEYM: TStringField
      FieldName = 'COMPUTEYM'
      Size = 5
    end
    object ADOQuery1BELONGYM: TStringField
      FieldName = 'BELONGYM'
      Size = 5
    end
    object ADOQuery1EMPID: TStringField
      FieldName = 'EMPID'
    end
    object ADOQuery1EMPNAME: TStringField
      FieldName = 'EMPNAME'
    end
    object ADOQuery1COMMISSION: TBCDField
      FieldName = 'COMMISSION'
      Precision = 10
      Size = 2
    end
    object ADOQuery1MEDIACODE: TIntegerField
      FieldName = 'MEDIACODE'
    end
    object ADOQuery1MEDIANAME: TStringField
      FieldName = 'MEDIANAME'
    end
    object ADOQuery1COMPTYPE: TStringField
      FieldName = 'COMPTYPE'
      FixedChar = True
      Size = 1
    end
    object ADOQuery1CLANCOMPCODE: TIntegerField
      FieldName = 'CLANCOMPCODE'
    end
    object ADOQuery1CLANCOMPNAME: TStringField
      FieldName = 'CLANCOMPNAME'
    end
    object ADOQuery1OPERATOR: TStringField
      FieldName = 'OPERATOR'
    end
    object ADOQuery1UPDTIME: TStringField
      FieldName = 'UPDTIME'
    end
  end
  object qryPriv: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    Left = 200
    Top = 16
  end
  object qryCD019: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM CD019 WHERE CODENO IS NULL')
    Left = 400
    Top = 152
  end
  object qrySO122Data: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select COMPCODE,COMPUTEYM,BELONGYM,STBNO,SNO,BILLNO,ITEM,'
      
        '              PERIODID,WORKEREN1,WORKEREN1NAME,WORKEREN1COMM,TO_' +
        'CHAR(WORKEREN1COMM) StrWORKEREN1COMM,CLCTEN,CLCTENNAME,'
      
        '              CLCTENCOMM,TO_CHAR(CLCTENCOMM) StrCLCTENCOMM,CLCTE' +
        'NORIPERCENT,INTACCID,INTACCNAME,INTACCCOMM,TO_CHAR(INTACCCOMM) S' +
        'trINTACCCOMM,INTACCORIPERCENT,'
      
        '              STAFFCODE,REALSTARTDATE,REALSTOPDATE,OPERATOR,UPDT' +
        'IME,BUYORRENT,CITEMCODE,'
      
        '              CITEMNAME,MEDIACODE,MEDIANAME,REALAMT,TO_CHAR(REAL' +
        'AMT) StrREALAMT,PROMCODE,PROMNAME,REALDATE,SHOULDDATE,'
      '              CUSTID,FirstFlag'
      '            from so122 WHERE BELONGYM IS NULL'
      '')
    Left = 616
    Top = 272
  end
  object dspSO122Data: TDataSetProvider
    DataSet = qrySO122Data
    Options = [poAllowCommandText]
    Left = 625
    Top = 286
  end
  object cdsSO122BOX: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'COMPCODE'
        DataType = ftInteger
      end
      item
        Name = 'COMPUTEYM'
        DataType = ftString
        Size = 6
      end
      item
        Name = 'BELONGYM'
        DataType = ftString
        Size = 6
      end
      item
        Name = 'STBNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'SNO'
        DataType = ftString
        Size = 15
      end
      item
        Name = 'WORKEREN1'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'WORKEREN1NAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'WORKEREN1COMM'
        DataType = ftBCD
        Precision = 10
        Size = 2
      end
      item
        Name = 'INTACCID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'INTACCNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'INTACCCOMM'
        DataType = ftInteger
      end
      item
        Name = 'STAFFCODE'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'REALSTARTDATE'
        DataType = ftDateTime
      end
      item
        Name = 'REALSTOPDATE'
        DataType = ftDateTime
      end
      item
        Name = 'REALSTARTDATE2'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'REALSTOPDATE2'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'OPERATOR'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'UPDTIME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'BUYORRENT'
        DataType = ftInteger
      end
      item
        Name = 'MEDIACODE'
        DataType = ftInteger
      end
      item
        Name = 'MEDIANAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'REALAMT'
        DataType = ftInteger
      end
      item
        Name = 'REALDATE'
        DataType = ftDate
      end
      item
        Name = 'SHOULDDATE'
        DataType = ftDate
      end
      item
        Name = 'REALDATE2'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'SHOULDDATE2'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'PROMCODE'
        DataType = ftBCD
        Precision = 10
        Size = 4
      end
      item
        Name = 'PROMNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CustID'
        DataType = ftInteger
      end
      item
        Name = 'CustName'
        DataType = ftString
        Size = 50
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    OnCalcFields = cdsSO122BOXCalcFields
    Left = 616
    Top = 352
    object cdsSO122BOXCOMPCODE: TIntegerField
      DisplayLabel = #20844#21496#21029
      DisplayWidth = 12
      FieldName = 'COMPCODE'
    end
    object cdsSO122BOXCOMPUTEYM: TStringField
      DisplayLabel = #32113#35336#24180#26376
      DisplayWidth = 11
      FieldName = 'COMPUTEYM'
      Size = 6
    end
    object cdsSO122BOXBELONGYM: TStringField
      DisplayLabel = #27512#23660#24180#26376
      DisplayWidth = 11
      FieldName = 'BELONGYM'
      Size = 6
    end
    object cdsSO122BOXSTBNO: TStringField
      DisplayLabel = 'STB '#24207#34399
      DisplayWidth = 6
      FieldName = 'STBNO'
    end
    object cdsSO122BOXSNO: TStringField
      DisplayLabel = #27966#24037#21934#34399
      DisplayWidth = 11
      FieldName = 'SNO'
      Size = 15
    end
    object cdsSO122BOXWORKEREN1: TStringField
      DisplayLabel = #24037#31243#20154#21729#20195#34399
      DisplayWidth = 24
      FieldName = 'WORKEREN1'
    end
    object cdsSO122BOXWORKEREN1NAME: TStringField
      DisplayLabel = #24037#31243#20154#21729#22995#21517
      DisplayWidth = 24
      FieldName = 'WORKEREN1NAME'
    end
    object cdsSO122BOXWORKEREN1COMM: TBCDField
      DisplayLabel = #24037#31243#20154#21729#20323#37329
      FieldName = 'WORKEREN1COMM'
      Precision = 10
      Size = 2
    end
    object cdsSO122BOXINTACCID: TStringField
      DisplayLabel = #20171#32057#20154'/'#21463#29702#20154'/'#25512#24291#20154#20195#34399
      DisplayWidth = 33
      FieldName = 'INTACCID'
    end
    object cdsSO122BOXINTACCNAME: TStringField
      DisplayLabel = #20171#32057#20154'/'#21463#29702#20154'/'#25512#24291#20154#22995#21517
      DisplayWidth = 33
      FieldName = 'INTACCNAME'
    end
    object cdsSO122BOXINTACCCOMM: TIntegerField
      DisplayLabel = #20171#32057#20154'/'#21463#29702#20154#20323#37329
      FieldName = 'INTACCCOMM'
    end
    object cdsSO122BOXSTAFFCODE: TStringField
      DisplayLabel = #25512#24291#20154#21729#24037#39006#21029
      DisplayWidth = 20
      FieldName = 'STAFFCODE'
      Size = 1
    end
    object cdsSO122BOXREALSTARTDATE: TDateTimeField
      DisplayLabel = #35037#27231#26085#26399
      DisplayWidth = 22
      FieldName = 'REALSTARTDATE'
      Visible = False
    end
    object cdsSO122BOXREALSTOPDATE: TDateTimeField
      DisplayLabel = #25286#27231#26085#26399
      DisplayWidth = 22
      FieldName = 'REALSTOPDATE'
      Visible = False
    end
    object cdsSO122BOXREALSTARTDATE2: TStringField
      DisplayLabel = #35037#27231#26085#26399
      FieldName = 'REALSTARTDATE2'
    end
    object cdsSO122BOXREALSTOPDATE2: TStringField
      DisplayLabel = #25286#27231#26085#26399
      FieldName = 'REALSTOPDATE2'
    end
    object cdsSO122BOXOPERATOR: TStringField
      DisplayLabel = #35336#31639#20154#21729
      DisplayWidth = 24
      FieldName = 'OPERATOR'
    end
    object cdsSO122BOXUPDTIME: TStringField
      DisplayLabel = #35336#31639#26178#38291
      DisplayWidth = 24
      FieldName = 'UPDTIME'
    end
    object cdsSO122BOXBUYORRENT: TIntegerField
      DisplayWidth = 17
      FieldName = 'BUYORRENT'
    end
    object cdsSO122BOXBuyName: TStringField
      DisplayLabel = #36023'/'#31199
      FieldKind = fkCalculated
      FieldName = 'BuyName'
      Calculated = True
    end
    object cdsSO122BOXMEDIACODE: TIntegerField
      DisplayLabel = #20171#32057#23186#20171#20195#30908
      DisplayWidth = 17
      FieldName = 'MEDIACODE'
    end
    object cdsSO122BOXMEDIANAME: TStringField
      DisplayLabel = #20171#32057#23186#20171#21517#31281
      DisplayWidth = 24
      FieldName = 'MEDIANAME'
    end
    object cdsSO122BOXREALDATE2: TDateField
      DisplayLabel = #23526#25910#26085#26399
      DisplayWidth = 12
      FieldName = 'REALDATE'
      Visible = False
    end
    object cdsSO122BOXSHOULDDATE2: TDateField
      DisplayLabel = #25033#25910#26085#26399
      DisplayWidth = 12
      FieldName = 'SHOULDDATE'
      Visible = False
    end
    object cdsSO122BOXREALDATE22: TStringField
      DisplayLabel = #23526#25910#26085#26399
      FieldName = 'REALDATE2'
    end
    object cdsSO122BOXSHOULDDATE22: TStringField
      DisplayLabel = #25033#25910#26085#26399
      FieldName = 'SHOULDDATE2'
    end
    object cdsSO122BOXPROMCODE: TBCDField
      DisplayLabel = #26989#21209#27963#21205#20195#30908
      FieldName = 'PROMCODE'
      Precision = 10
    end
    object cdsSO122BOXPROMNAME: TStringField
      DisplayLabel = #26989#21209#27963#21205#21517#31281
      FieldName = 'PROMNAME'
    end
    object cdsSO122BOXCustID: TIntegerField
      DisplayLabel = #23458#32232
      FieldName = 'CustID'
    end
  end
  object qryCD042Data: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    Left = 624
    Top = 216
  end
  object qrySO105Prom: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    Left = 288
    Top = 96
  end
  object qryDateData: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    Left = 240
    Top = 288
  end
  object cdsSO122Channel: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'COMPCODE'
        DataType = ftInteger
      end
      item
        Name = 'COMPUTEYM'
        DataType = ftString
        Size = 6
      end
      item
        Name = 'BELONGYM'
        DataType = ftString
        Size = 6
      end
      item
        Name = 'BILLNO'
        DataType = ftString
        Size = 15
      end
      item
        Name = 'ITEM'
        DataType = ftInteger
      end
      item
        Name = 'PERIODID'
        DataType = ftInteger
      end
      item
        Name = 'CLCTEN'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CLCTENNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CLCTENCOMM'
        DataType = ftFloat
      end
      item
        Name = 'CLCTENORIPERCENT'
        DataType = ftFloat
      end
      item
        Name = 'INTACCID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'INTACCNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'INTACCCOMM'
        DataType = ftFloat
      end
      item
        Name = 'INTACCORIPERCENT'
        DataType = ftFloat
      end
      item
        Name = 'STAFFCODE'
        DataType = ftString
        Size = 1
      end
      item
        Name = 'REALSTARTDATE'
        DataType = ftDateTime
      end
      item
        Name = 'REALSTOPDATE'
        DataType = ftDateTime
      end
      item
        Name = 'REALSTARTDATE2'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'REALSTOPDATE2'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'OPERATOR'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'UPDTIME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CITEMCODE'
        DataType = ftInteger
      end
      item
        Name = 'CITEMNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'MEDIACODE'
        DataType = ftInteger
      end
      item
        Name = 'MEDIANAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'REALAMT'
        DataType = ftFloat
      end
      item
        Name = 'REALDATE'
        DataType = ftDate
      end
      item
        Name = 'SHOULDDATE'
        DataType = ftDate
      end
      item
        Name = 'REALDATE2'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'SHOULDDATE2'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'PROMCODE'
        DataType = ftBCD
        Precision = 10
        Size = 4
      end
      item
        Name = 'PROMNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CustID'
        DataType = ftInteger
      end
      item
        Name = 'CustName'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'FirstFlagName'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 616
    Top = 404
    object cdsSO122ChannelCOMPCODE: TIntegerField
      DisplayLabel = #20844#21496#21029
      FieldName = 'COMPCODE'
    end
    object cdsSO122ChannelCOMPUTEYM: TStringField
      DisplayLabel = #32113#35336#24180#26376
      FieldName = 'COMPUTEYM'
      Size = 6
    end
    object cdsSO122ChannelBELONGYM: TStringField
      DisplayLabel = #27512#23660#24180#26376
      FieldName = 'BELONGYM'
      Size = 6
    end
    object cdsSO122ChannelBILLNO: TStringField
      DisplayLabel = #21934#25818#32232#34399
      FieldName = 'BILLNO'
      Size = 15
    end
    object cdsSO122ChannelITEM: TIntegerField
      DisplayLabel = #38917#27425
      FieldName = 'ITEM'
    end
    object cdsSO122ChannelPERIODID: TIntegerField
      DisplayLabel = #26399#25976#21029
      FieldName = 'PERIODID'
    end
    object cdsSO122ChannelCLCTEN: TStringField
      DisplayLabel = #25910#36027#21729#20195#34399
      FieldName = 'CLCTEN'
    end
    object cdsSO122ChannelCLCTENNAME: TStringField
      DisplayLabel = #25910#36027#21729#22995#21517
      FieldName = 'CLCTENNAME'
    end
    object cdsSO122ChannelCLCTENCOMM: TFloatField
      DisplayLabel = #25910#36027#21729#20323#37329
      FieldName = 'CLCTENCOMM'
    end
    object cdsSO122ChannelCLCTENORIPERCENT: TFloatField
      DisplayLabel = #25910#36027#21729#21407#20323#37329'%'
      FieldName = 'CLCTENORIPERCENT'
    end
    object cdsSO122ChannelINTACCID: TStringField
      DisplayLabel = #20171#32057#20154'/'#21463#29702#20154'/'#25512#24291#20154#20195#34399
      FieldName = 'INTACCID'
    end
    object cdsSO122ChannelINTACCNAME: TStringField
      DisplayLabel = #20171#32057#20154'/'#21463#29702#20154'/'#25512#24291#20154#22995#21517
      FieldName = 'INTACCNAME'
    end
    object cdsSO122ChannelINTACCCOMM: TFloatField
      DisplayLabel = #20171#32057#20154'/'#21463#29702#20154#20323#37329
      FieldName = 'INTACCCOMM'
    end
    object cdsSO122ChannelINTACCORIPERCENT: TFloatField
      DisplayLabel = #20171#32057#20154'/'#21463#29702#20154#21407#20323#37329'%'
      FieldName = 'INTACCORIPERCENT'
    end
    object cdsSO122ChannelSTAFFCODE: TStringField
      DisplayLabel = #25512#24291#20154#21729#24037#39006#21029
      FieldName = 'STAFFCODE'
      Size = 1
    end
    object cdsSO122ChannelREALSTARTDATE: TDateTimeField
      DisplayLabel = #26377#25928#36215#22987#26085#26399
      FieldName = 'REALSTARTDATE'
      Visible = False
    end
    object cdsSO122ChannelREALSTOPDATE: TDateTimeField
      DisplayLabel = #26377#25928#25130#27490#26085#26399
      FieldName = 'REALSTOPDATE'
      Visible = False
    end
    object cdsSO122ChannelREALSTARTDATE2: TStringField
      DisplayLabel = #26377#25928#36215#22987#26085#26399
      FieldName = 'REALSTARTDATE2'
    end
    object cdsSO122ChannelREALSTOPDATE2: TStringField
      DisplayLabel = #26377#25928#25130#27490#26085#26399
      FieldName = 'REALSTOPDATE2'
    end
    object cdsSO122ChannelOPERATOR: TStringField
      DisplayLabel = #35336#31639#20154#21729
      FieldName = 'OPERATOR'
    end
    object cdsSO122ChannelUPDTIME: TStringField
      DisplayLabel = #35336#31639#26178#38291
      FieldName = 'UPDTIME'
    end
    object cdsSO122ChannelCITEMCODE: TIntegerField
      DisplayLabel = #38971#36947#25910#36027#20195#30908
      FieldName = 'CITEMCODE'
    end
    object cdsSO122ChannelCITEMNAME: TStringField
      DisplayLabel = #38971#36947#25910#36027#21517#31281
      FieldName = 'CITEMNAME'
    end
    object cdsSO122ChannelMEDIACODE: TIntegerField
      DisplayLabel = #20171#32057#23186#20171#20195#30908
      FieldName = 'MEDIACODE'
    end
    object cdsSO122ChannelMEDIANAME: TStringField
      DisplayLabel = #20171#32057#23186#20171#21517#31281
      FieldName = 'MEDIANAME'
    end
    object cdsSO122ChannelREALAMT: TFloatField
      DisplayLabel = #38971#36947#23526#25910#37329#38989
      FieldName = 'REALAMT'
    end
    object cdsSO122ChannelREALDATE: TDateField
      DisplayLabel = #38971#36947#23526#25910#26085#26399
      FieldName = 'REALDATE'
      Visible = False
    end
    object cdsSO122ChannelSHOULDDATE: TDateField
      DisplayLabel = #25033#25910#26085#26399
      FieldName = 'SHOULDDATE'
      Visible = False
    end
    object cdsSO122ChannelREALDATE2: TStringField
      DisplayLabel = #38971#36947#23526#25910#26085#26399
      FieldName = 'REALDATE2'
    end
    object cdsSO122ChannelSHOULDDATE2: TStringField
      DisplayLabel = #25033#25910#26085#26399
      FieldName = 'SHOULDDATE2'
    end
    object cdsSO122ChannelPROMCODE: TBCDField
      DisplayLabel = #26989#21209#27963#21205#20195#30908
      FieldName = 'PROMCODE'
      Precision = 10
    end
    object cdsSO122ChannelPROMNAME: TStringField
      DisplayLabel = #26989#21209#27963#21205#21517#31281
      FieldName = 'PROMNAME'
    end
    object cdsSO122ChannelCustID: TIntegerField
      DisplayLabel = #23458#32232
      FieldName = 'CustID'
    end
    object cdsSO122ChannelFirstFlagName: TStringField
      DisplayLabel = #26159#21542#39318#26399
      FieldName = 'FirstFlagName'
    end
  end
  object qrySO132: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    BeforePost = qrySO132BeforePost
    BeforeDelete = qrySO132BeforeDelete
    Parameters = <>
    SQL.Strings = (
      'select * from SO132 where compcode=-1')
    Left = 552
    Top = 272
    object qrySO132COMPCODE: TIntegerField
      FieldName = 'COMPCODE'
      Visible = False
    end
    object qrySO132CUSTID: TIntegerField
      DisplayLabel = #23458#32232
      FieldName = 'CUSTID'
    end
    object qrySO132CITEMCODE: TIntegerField
      DisplayLabel = #25910#36027#38917#30446#20195#30908
      FieldName = 'CITEMCODE'
    end
    object qrySO132CITEMNAME: TStringField
      DisplayLabel = #25910#36027#38917#30446
      FieldName = 'CITEMNAME'
    end
    object qrySO132EMPNO: TStringField
      DisplayLabel = #27512#23660#32773#20195#30908
      FieldName = 'EMPNO'
      Size = 10
    end
    object qrySO132EMPNAME: TStringField
      DisplayLabel = #27512#23660#32773
      FieldName = 'EMPNAME'
    end
    object qrySO132DATACREATOR: TStringField
      DisplayLabel = #36039#26009#24314#31435#32773
      FieldName = 'DATACREATOR'
    end
    object qrySO132UPDEN: TStringField
      DisplayLabel = #30064#21205#20154#21729
      FieldName = 'UPDEN'
    end
    object qrySO132UPDTIME: TDateTimeField
      DisplayLabel = #30064#21205#26178#38291
      FieldName = 'UPDTIME'
    end
  end
  object qrySo133: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select * from SO133 where seq <0')
    Left = 552
    Top = 224
    object qrySo133SEQ: TBCDField
      FieldName = 'SEQ'
      Precision = 10
      Size = 0
    end
    object qrySo133COMPCODE: TIntegerField
      FieldName = 'COMPCODE'
    end
    object qrySo133OLDCITEMCODE: TIntegerField
      FieldName = 'OLDCITEMCODE'
    end
    object qrySo133OLDCITEMNAME: TStringField
      FieldName = 'OLDCITEMNAME'
    end
    object qrySo133NEWCITEMCODE: TIntegerField
      FieldName = 'NEWCITEMCODE'
    end
    object qrySo133NEWCITEMNAME: TStringField
      FieldName = 'NEWCITEMNAME'
    end
    object qrySo133OLDEMPNO: TStringField
      FieldName = 'OLDEMPNO'
      Size = 10
    end
    object qrySo133OLDEMPNAME: TStringField
      FieldName = 'OLDEMPNAME'
    end
    object qrySo133NEWEMPNO: TStringField
      FieldName = 'NEWEMPNO'
      Size = 10
    end
    object qrySo133NEWEMPNAME: TStringField
      FieldName = 'NEWEMPNAME'
    end
    object qrySo133OPERATIONMODE: TStringField
      FieldName = 'OPERATIONMODE'
      FixedChar = True
      Size = 1
    end
    object qrySo133UPDEN: TStringField
      FieldName = 'UPDEN'
    end
    object qrySo133UPDTIME: TDateTimeField
      FieldName = 'UPDTIME'
    end
  end
  object ADOQuery2: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    Left = 200
    Top = 416
  end
  object qryCodeCM003: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select  EMPNO,EMPNAME from CM003'
      'where STOPFLAG=0 and WORKCLASS='#39'C'#39
      '')
    Left = 552
    Top = 320
    object qryCm003EMPNO: TStringField
      DisplayLabel = #27512#23660#32773#20195#30908
      FieldName = 'EMPNO'
      Size = 10
    end
    object qryCm003EMPNAME: TStringField
      DisplayLabel = #27512#23660#32773
      FieldName = 'EMPNAME'
    end
  end
  object qryCodeCD019_2: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select CODENO,DESCRIPTION from CD019'
      'where StopFlag=0 and RefNo=2')
    Left = 552
    Top = 376
    object qryCD019CODENO: TIntegerField
      DisplayLabel = #25910#36027#38917#30446#20195#30908
      FieldName = 'CODENO'
    end
    object qryCD019DESCRIPTION: TStringField
      DisplayLabel = #25910#36027#38917#30446
      FieldName = 'DESCRIPTION'
    end
  end
  object ADOCommand1: TADOCommand
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    Left = 544
    Top = 424
  end
  object cdsSo033Log: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'CUSTID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'BILLNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ITEM'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CITEMCODE'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'CITEMNAME'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'SHOULDDATE'
        DataType = ftDate
      end
      item
        Name = 'SHOULDAMT'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'REALDATE'
        DataType = ftDate
      end
      item
        Name = 'REALAMT'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'REALPERIOD'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'REALSTARTDATE'
        DataType = ftDate
      end
      item
        Name = 'SBILLNO'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'SITEM'
        DataType = ftInteger
      end
      item
        Name = 'Notes'
        DataType = ftString
        Size = 50
      end
      item
        Name = 'SNo'
        DataType = ftString
        Size = 15
      end
      item
        Name = 'SEQNo'
        DataType = ftString
        Size = 15
      end
      item
        Name = 'PRSNO'
        DataType = ftString
        Size = 15
      end
      item
        Name = 'InstDate'
        DataType = ftDate
      end
      item
        Name = 'PRDate'
        DataType = ftDate
      end
      item
        Name = 'PromCode'
        DataType = ftBCD
        Precision = 10
        Size = 4
      end
      item
        Name = 'PromName'
        DataType = ftString
        Size = 30
      end
      item
        Name = 'DiscountCode'
        DataType = ftInteger
      end
      item
        Name = 'DiscountName'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 328
    Top = 160
    object cdsSo033LogCUSTID: TStringField
      DisplayLabel = #23458#25142#32232#34399
      FieldName = 'CUSTID'
    end
    object cdsSo033LogBILLNO: TStringField
      DisplayLabel = #21934#25818#32232#34399
      FieldName = 'BILLNO'
    end
    object cdsSo033LogITEM: TStringField
      DisplayLabel = #38917#27425
      FieldName = 'ITEM'
    end
    object cdsSo033LogCITEMCODE: TStringField
      DisplayLabel = #25910#36027#38917#30446#20195#34399
      FieldName = 'CITEMCODE'
    end
    object cdsSo033LogCITEMNAME: TStringField
      DisplayLabel = #25910#36027#38917#30446#21517#31281
      FieldName = 'CITEMNAME'
    end
    object cdsSo033LogSHOULDDATE: TDateField
      DisplayLabel = #25033#25910#26085#26399
      FieldName = 'SHOULDDATE'
    end
    object cdsSo033LogSHOULDAMT: TStringField
      DisplayLabel = #25033#25910#37329#38989
      FieldName = 'SHOULDAMT'
    end
    object cdsSo033LogREALDATE: TDateField
      DisplayLabel = #23526#25910#26085#26399
      FieldName = 'REALDATE'
    end
    object cdsSo033LogREALAMT: TStringField
      DisplayLabel = #23526#25910#37329#38989
      FieldName = 'REALAMT'
    end
    object cdsSo033LogREALPERIOD: TStringField
      DisplayLabel = #25910#36027#26399#25976
      FieldName = 'REALPERIOD'
    end
    object cdsSo033LogREALSTARTDATE: TDateField
      DisplayLabel = #26377#25928#36215#22987#26085#26399
      FieldName = 'REALSTARTDATE'
    end
    object cdsSo033LogSBILLNO: TStringField
      DisplayLabel = #21407#22987#21934#34399
      FieldName = 'SBILLNO'
    end
    object cdsSo033LogSITEM: TIntegerField
      DisplayLabel = #21407#22987#38917#27425
      FieldName = 'SITEM'
    end
    object cdsSo033LogPromCode: TBCDField
      DisplayLabel = #26989#21209#27963#21205#20195#30908
      FieldName = 'PromCode'
      Precision = 10
    end
    object cdsSo033LogPromName: TStringField
      DisplayLabel = #26989#21209#27963#21205#21517#31281
      FieldName = 'PromName'
      Size = 30
    end
    object cdsSo033LogDiscountCode: TIntegerField
      DisplayLabel = #20778#24800#36774#27861#20195#30908
      FieldName = 'DiscountCode'
    end
    object cdsSo033LogDiscountName: TStringField
      DisplayLabel = #20778#24800#36774#27861#21517#31281
      FieldName = 'DiscountName'
    end
    object cdsSo033LogNotes: TStringField
      DisplayLabel = #20633#35387
      FieldName = 'Notes'
      Size = 50
    end
    object cdsSo033LogSNo: TStringField
      DisplayLabel = #20358#28304#21934#32232#34399
      FieldName = 'SNo'
      Size = 15
    end
    object cdsSo033LogSEQNo: TStringField
      DisplayLabel = #35373#20633#27969#27700#34399
      FieldName = 'SEQNo'
      Size = 15
    end
    object cdsSo033LogPRSNO: TStringField
      DisplayLabel = #20572#25286#31227#27231#21934#34399
      FieldName = 'PRSNO'
      Size = 15
    end
    object cdsSo033LogInstDate: TDateField
      DisplayLabel = #23433#35037#26085#26399
      FieldName = 'InstDate'
    end
    object cdsSo033LogPRDate: TDateField
      DisplayLabel = #25286#38500#26085#26399
      FieldName = 'PRDate'
    end
  end
  object cdsSO122Data: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO122Data'
    Left = 632
    Top = 299
  end
  object dspSo033So034: TDataSetProvider
    DataSet = qrySo033So034
    Options = [poAllowCommandText]
    Left = 40
    Top = 248
  end
  object cdsSo033So034: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSo033So034'
    Left = 48
    Top = 268
  end
  object qryTempBox: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    Left = 40
    Top = 336
  end
  object dspTempBox: TDataSetProvider
    DataSet = qryTempBox
    Options = [poAllowCommandText]
    Left = 48
    Top = 344
  end
  object cdsTempBox: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspTempBox'
    Left = 56
    Top = 352
  end
  object qrySO004: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO004 WHERE SEQNo IS NOT NULL')
    Left = 128
    Top = 376
  end
  object qrySO134: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM SO134'
      'where COMPCODE is NULL')
    Left = 632
    Top = 152
  end
  object dspSO134: TDataSetProvider
    DataSet = qrySO134
    Options = [poAllowCommandText]
    Left = 640
    Top = 160
  end
  object cdsSO134: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO134'
    Left = 648
    Top = 168
  end
  object spGetChargeData: TADOStoredProc
    Connection = frmMainMenu.ADOConnection1
    ProcedureName = 'SF_GETCHARGEDATA'
    Parameters = <
      item
        Name = 'RETURN_VALUE'
        DataType = ftBCD
        Direction = pdReturnValue
        Value = Null
      end
      item
        Name = 'P_TABLE'
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'P_COMPCODE'
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'P_BELONGYM'
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'P_BILLNO'
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'P_ITEM'
        DataType = ftString
        Size = 4000
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
    Left = 552
    Top = 136
  end
  object dspGetChargeData: TDataSetProvider
    DataSet = spGetChargeData
    Options = [poAllowCommandText]
    Left = 560
    Top = 152
  end
  object cdsGetChargeData: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspGetChargeData'
    Left = 568
    Top = 168
  end
  object qrySO135: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select * from SO135'
      'where COMPCODE IS NULL')
    Left = 632
    Top = 88
  end
  object dspSO135: TDataSetProvider
    DataSet = qrySO135
    Options = [poAllowCommandText]
    Left = 640
    Top = 96
  end
  object cdsSO135: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO135'
    Left = 648
    Top = 104
  end
  object ClientDataSet1: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO122Data'
    Left = 488
    Top = 235
    object IntegerField2: TIntegerField
      FieldName = 'COMPCODE'
    end
    object StringField7: TStringField
      FieldName = 'COMPUTEYM'
      Size = 5
    end
    object StringField8: TStringField
      FieldName = 'BELONGYM'
      Size = 5
    end
    object StringField9: TStringField
      FieldName = 'STBNO'
    end
    object StringField10: TStringField
      FieldName = 'SNO'
      Size = 15
    end
    object StringField11: TStringField
      FieldName = 'BILLNO'
      Size = 15
    end
    object IntegerField3: TIntegerField
      FieldName = 'ITEM'
    end
    object IntegerField4: TIntegerField
      FieldName = 'PERIODID'
    end
    object StringField12: TStringField
      FieldName = 'WORKEREN1'
    end
    object StringField13: TStringField
      FieldName = 'WORKEREN1NAME'
    end
    object BCDField2: TBCDField
      FieldName = 'WORKEREN1COMM'
      Precision = 10
      Size = 2
    end
    object StringField14: TStringField
      FieldName = 'STRWORKEREN1COMM'
      ReadOnly = True
      Size = 40
    end
    object StringField15: TStringField
      FieldName = 'CLCTEN'
    end
    object StringField16: TStringField
      FieldName = 'CLCTENNAME'
    end
    object BCDField3: TBCDField
      FieldName = 'CLCTENCOMM'
      Precision = 10
      Size = 2
    end
    object StringField17: TStringField
      FieldName = 'STRCLCTENCOMM'
      ReadOnly = True
      Size = 40
    end
    object BCDField4: TBCDField
      FieldName = 'CLCTENORIPERCENT'
      Precision = 8
      Size = 2
    end
    object StringField18: TStringField
      FieldName = 'INTACCID'
    end
    object StringField19: TStringField
      FieldName = 'INTACCNAME'
    end
    object BCDField5: TBCDField
      FieldName = 'INTACCCOMM'
      Precision = 10
      Size = 2
    end
    object StringField20: TStringField
      FieldName = 'STRINTACCCOMM'
      ReadOnly = True
      Size = 40
    end
    object BCDField6: TBCDField
      FieldName = 'INTACCORIPERCENT'
      Precision = 8
      Size = 2
    end
    object StringField21: TStringField
      FieldName = 'STAFFCODE'
      Size = 1
    end
    object DateTimeField1: TDateTimeField
      FieldName = 'REALSTARTDATE'
    end
    object DateTimeField2: TDateTimeField
      FieldName = 'REALSTOPDATE'
    end
    object StringField22: TStringField
      FieldName = 'OPERATOR'
    end
    object StringField23: TStringField
      FieldName = 'UPDTIME'
    end
    object IntegerField5: TIntegerField
      FieldName = 'BUYORRENT'
    end
    object IntegerField6: TIntegerField
      FieldName = 'CITEMCODE'
    end
    object StringField24: TStringField
      FieldName = 'CITEMNAME'
    end
    object IntegerField7: TIntegerField
      FieldName = 'MEDIACODE'
    end
    object StringField25: TStringField
      FieldName = 'MEDIANAME'
    end
    object IntegerField8: TIntegerField
      FieldName = 'REALAMT'
    end
    object StringField26: TStringField
      FieldName = 'STRREALAMT'
      ReadOnly = True
      Size = 40
    end
    object IntegerField9: TBCDField
      FieldName = 'PROMCODE'
      Precision = 10
      Size = 0
    end
    object StringField27: TStringField
      FieldName = 'PROMNAME'
    end
    object DateTimeField3: TDateTimeField
      FieldName = 'REALDATE'
    end
    object DateTimeField4: TDateTimeField
      FieldName = 'SHOULDDATE'
    end
    object IntegerField10: TIntegerField
      FieldName = 'CUSTID'
    end
    object StringField28: TStringField
      FieldName = 'FIRSTFLAG'
      FixedChar = True
      Size = 1
    end
  end
  object qrySO134Data: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select COMPCODE,COMPUTEYM,BELONGYM,BILLNO,ITEM,'
      
        '              PERIODID,CLCTEN,CLCTENNAME,CLCTENCOMM,TO_CHAR(CLCT' +
        'ENCOMM) StrCLCTENCOMM,'
      
        '              CLCTENORIPERCENT,INTACCID,INTACCNAME,INTACCCOMM,TO' +
        '_CHAR(INTACCCOMM) StrINTACCCOMM,INTACCORIPERCENT,'
      
        '              STAFFCODE,REALSTARTDATE,REALSTOPDATE,OPERATOR,UPDT' +
        'IME,CITEMCODE,'
      
        '              CITEMNAME,MEDIACODE,MEDIANAME,REALAMT,TO_CHAR(REAL' +
        'AMT) StrREALAMT,PROMCODE,PROMNAME,'
      '              REALDATE,SHOULDDATE,CUSTID,FirstFlag'
      '            from so134 WHERE BELONGYM IS NULL')
    Left = 480
    Top = 288
  end
  object dspSO134Data: TDataSetProvider
    DataSet = qrySO134Data
    Options = [poAllowCommandText]
    Left = 488
    Top = 296
  end
  object cdsSO134Data: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspSO134Data'
    Left = 496
    Top = 304
  end
  object cdsStatisticExcel: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'CompCode'
        DataType = ftString
        Size = 3
      end
      item
        Name = 'EmpID'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'EmpName'
        DataType = ftString
        Size = 30
      end
      item
        Name = 'WorkerCounts'
        DataType = ftInteger
      end
      item
        Name = 'WorkerComm'
        DataType = ftFloat
      end
      item
        Name = 'StrWComm'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ClctCounts'
        DataType = ftInteger
      end
      item
        Name = 'ClctComm'
        DataType = ftFloat
      end
      item
        Name = 'StrCComm'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'BoxIntAccCounts'
        DataType = ftInteger
      end
      item
        Name = 'BoxIntAccComm'
        DataType = ftFloat
      end
      item
        Name = 'BoxStrIComm'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'ChIntAccCounts'
        DataType = ftInteger
      end
      item
        Name = 'ChIntAccComm'
        DataType = ftFloat
      end
      item
        Name = 'ChStrIComm'
        DataType = ftString
        Size = 20
      end
      item
        Name = 'TotalComm'
        DataType = ftFloat
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 552
    Top = 88
    object cdsStatisticExcelCompCode: TStringField
      FieldName = 'CompCode'
      Size = 3
    end
    object cdsStatisticExcelEmpID: TStringField
      DisplayLabel = #21729#24037#32232#34399
      FieldName = 'EmpID'
    end
    object cdsStatisticExcelEmpName: TStringField
      DisplayLabel = #21729#24037#21517#31281
      FieldName = 'EmpName'
      Size = 30
    end
    object cdsStatisticExcelWorkerCounts: TIntegerField
      DisplayLabel = #35037#27231#25976#37327
      FieldName = 'WorkerCounts'
    end
    object cdsStatisticExcelWorkerComm: TFloatField
      DisplayLabel = #35037#27231#29518#37329
      FieldName = 'WorkerComm'
    end
    object cdsStatisticExcelStrWComm: TStringField
      FieldName = 'StrWComm'
    end
    object cdsStatisticExcelClctCounts: TIntegerField
      DisplayLabel = #25910#36027#25976#37327
      FieldName = 'ClctCounts'
    end
    object cdsStatisticExcelClctComm: TFloatField
      DisplayLabel = #25910#36027#20323#37329
      FieldName = 'ClctComm'
    end
    object cdsStatisticExcelStrCComm: TStringField
      FieldName = 'StrCComm'
    end
    object cdsStatisticExcelBoxIntAccCounts: TIntegerField
      DisplayLabel = 'Box'#20171#32057#20154#25976#37327
      FieldName = 'BoxIntAccCounts'
    end
    object cdsStatisticExcelBoxIntAccComm: TFloatField
      DisplayLabel = 'Box'#20171#32057#20154#29518#37329
      FieldName = 'BoxIntAccComm'
    end
    object cdsStatisticExcelBoxStrIComm: TStringField
      FieldName = 'BoxStrIComm'
    end
    object cdsStatisticExcelChIntAccCounts: TIntegerField
      DisplayLabel = #38971#36947#20171#32057#20154#25976#37327
      FieldName = 'ChIntAccCounts'
    end
    object cdsStatisticExcelChIntAccComm: TFloatField
      DisplayLabel = #38971#36947#20171#32057#20154#20323#37329
      FieldName = 'ChIntAccComm'
    end
    object cdsStatisticExcelChStrIComm: TStringField
      FieldName = 'ChStrIComm'
    end
    object cdsStatisticExcelTotalComm: TFloatField
      DisplayLabel = #37329#38989#23567#35336
      FieldName = 'TotalComm'
    end
  end
  object cdsTotalCommData: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'Mode'
        DataType = ftString
        Size = 200
      end
      item
        Name = 'ModeName'
        DataType = ftString
        Size = 200
      end
      item
        Name = 'LastPay'
        DataType = ftFloat
      end
      item
        Name = 'NextPay'
        DataType = ftFloat
      end
      item
        Name = 'Month1'
        DataType = ftFloat
      end
      item
        Name = 'Month2'
        DataType = ftFloat
      end
      item
        Name = 'Month3'
        DataType = ftFloat
      end
      item
        Name = 'Month4'
        DataType = ftFloat
      end
      item
        Name = 'Month5'
        DataType = ftFloat
      end
      item
        Name = 'Month6'
        DataType = ftFloat
      end
      item
        Name = 'Month7'
        DataType = ftFloat
      end
      item
        Name = 'Month8'
        DataType = ftFloat
      end
      item
        Name = 'Month9'
        DataType = ftFloat
      end
      item
        Name = 'Month10'
        DataType = ftFloat
      end
      item
        Name = 'Month11'
        DataType = ftFloat
      end
      item
        Name = 'Month12'
        DataType = ftFloat
      end
      item
        Name = 'OtherMonth'
        DataType = ftFloat
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 480
    Top = 184
    object cdsTotalCommDataMode: TStringField
      FieldName = 'Mode'
      Size = 200
    end
    object cdsTotalCommDataModeName: TStringField
      FieldName = 'ModeName'
      Size = 200
    end
    object cdsTotalCommDataLastPay: TFloatField
      FieldName = 'LastPay'
    end
    object cdsTotalCommDataNextPay: TFloatField
      FieldName = 'NextPay'
    end
    object cdsTotalCommDataMonth1: TFloatField
      FieldName = 'Month1'
    end
    object cdsTotalCommDataMonth2: TFloatField
      FieldName = 'Month2'
    end
    object cdsTotalCommDataMonth3: TFloatField
      FieldName = 'Month3'
    end
    object cdsTotalCommDataMonth4: TFloatField
      FieldName = 'Month4'
    end
    object cdsTotalCommDataMonth5: TFloatField
      FieldName = 'Month5'
    end
    object cdsTotalCommDataMonth6: TFloatField
      FieldName = 'Month6'
    end
    object cdsTotalCommDataMonth7: TFloatField
      FieldName = 'Month7'
    end
    object cdsTotalCommDataMonth8: TFloatField
      FieldName = 'Month8'
    end
    object cdsTotalCommDataMonth9: TFloatField
      FieldName = 'Month9'
    end
    object cdsTotalCommDataMonth10: TFloatField
      FieldName = 'Month10'
    end
    object cdsTotalCommDataMonth11: TFloatField
      FieldName = 'Month11'
    end
    object cdsTotalCommDataMonth12: TFloatField
      FieldName = 'Month12'
    end
    object cdsTotalCommDataOtherMonth: TFloatField
      FieldName = 'OtherMonth'
    end
  end
  object qryTotalComm: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    SQL.Strings = (
      
        'select BELONGYM, sum(nvl(WORKEREN1COMM,0)) A1, sum(nvl(CLCTENCOM' +
        'M,0)) A2, sum(nvl(INTACCCOMM,0)) A3 '
      
        ',sum(nvl(WORKEREN1COMM,0)) + sum(nvl(CLCTENCOMM,0)) + sum(nvl(IN' +
        'TACCCOMM,0)) A4'
      'from so122 group by BELONGYM')
    Left = 472
    Top = 104
  end
  object dspTotalComm: TDataSetProvider
    DataSet = qryTotalComm
    Options = [poAllowCommandText]
    Left = 480
    Top = 120
  end
  object cdsTotalComm: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspTotalComm'
    Left = 488
    Top = 128
  end
  object cdsTotalMonthData: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'ComputeYM'
        DataType = ftString
        Size = 200
      end
      item
        Name = 'Month1'
        DataType = ftFloat
      end
      item
        Name = 'Month2'
        DataType = ftFloat
      end
      item
        Name = 'Month3'
        DataType = ftFloat
      end
      item
        Name = 'Month4'
        DataType = ftFloat
      end
      item
        Name = 'Month5'
        DataType = ftFloat
      end
      item
        Name = 'Month6'
        DataType = ftFloat
      end
      item
        Name = 'Month7'
        DataType = ftFloat
      end
      item
        Name = 'Month8'
        DataType = ftFloat
      end
      item
        Name = 'Month9'
        DataType = ftFloat
      end
      item
        Name = 'Month10'
        DataType = ftFloat
      end
      item
        Name = 'Month11'
        DataType = ftFloat
      end
      item
        Name = 'Month12'
        DataType = ftFloat
      end
      item
        Name = 'OtherMonth'
        DataType = ftFloat
      end
      item
        Name = 'TotalComm'
        DataType = ftFloat
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 464
    Top = 352
    object cdsTotalMonthDataComputeYM: TStringField
      DisplayWidth = 20
      FieldName = 'ComputeYM'
      Size = 200
    end
    object cdsTotalMonthDataMonth1: TFloatField
      FieldName = 'Month1'
    end
    object cdsTotalMonthDataMonth2: TFloatField
      FieldName = 'Month2'
    end
    object cdsTotalMonthDataMonth3: TFloatField
      FieldName = 'Month3'
    end
    object cdsTotalMonthDataMonth4: TFloatField
      FieldName = 'Month4'
    end
    object cdsTotalMonthDataMonth5: TFloatField
      FieldName = 'Month5'
    end
    object cdsTotalMonthDataMonth6: TFloatField
      FieldName = 'Month6'
    end
    object cdsTotalMonthDataMonth7: TFloatField
      FieldName = 'Month7'
    end
    object cdsTotalMonthDataMonth8: TFloatField
      FieldName = 'Month8'
    end
    object cdsTotalMonthDataMonth9: TFloatField
      FieldName = 'Month9'
    end
    object cdsTotalMonthDataMonth10: TFloatField
      FieldName = 'Month10'
    end
    object cdsTotalMonthDataMonth11: TFloatField
      FieldName = 'Month11'
    end
    object cdsTotalMonthDataMonth12: TFloatField
      FieldName = 'Month12'
    end
    object cdsTotalMonthDataOtherMonth: TFloatField
      FieldName = 'OtherMonth'
    end
    object cdsTotalMonthDataTotalComm: TFloatField
      FieldName = 'TotalComm'
    end
  end
  object qryComm1: TADOQuery
    Connection = frmMainMenu.ADOConnection1
    Parameters = <>
    Left = 432
    Top = 400
  end
  object dspComm1: TDataSetProvider
    DataSet = qryComm1
    Options = [poAllowCommandText]
    Left = 448
    Top = 408
  end
  object cdsComm1: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'dspComm1'
    Left = 456
    Top = 424
  end
end
