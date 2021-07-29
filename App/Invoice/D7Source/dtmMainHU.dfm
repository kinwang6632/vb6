object dtmMainH: TdtmMainH
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Left = 351
  Top = 160
  Height = 441
  Width = 437
  object adoComm: TADOQuery
    CacheSize = 100
    Connection = dtmMain.InvConnection
    CursorType = ctStatic
    Parameters = <>
    Left = 44
    Top = 24
  end
  object adoA07: TADOQuery
    CacheSize = 100
    Connection = dtmMain.InvConnection
    OnCalcFields = adoA07CalcFields
    Parameters = <>
    Prepared = True
    SQL.Strings = (
      'SELECT * FROM INV014 WHERE YEARMONTH IS NULL')
    Left = 220
    Top = 160
  end
  object cdsInvFormat: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'CodeNo'
        DataType = ftInteger
      end
      item
        Name = 'Description'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 48
    Top = 344
  end
  object cdsTaxType: TClientDataSet
    Aggregates = <>
    FieldDefs = <
      item
        Name = 'CodeNo'
        DataType = ftInteger
      end
      item
        Name = 'Description'
        DataType = ftString
        Size = 20
      end>
    IndexDefs = <>
    Params = <>
    StoreDefs = True
    Left = 48
    Top = 288
  end
  object adoInv019: TADOQuery
    Connection = dtmMain.InvConnection
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        'select TITLESNAME,TITLENAME INVTITLE,BUSINESSID,MZIPCODE ZIPCODE' +
        ',MAILADDR MAILADDRESS,INVADDR INVADDRESS,MEMO from Inv019 where ' +
        'CUSTID='#39#39)
    Left = 136
    Top = 100
    object adoInv019TITLESNAME: TStringField
      DisplayLabel = #25260#38957#31777#31281
      FieldName = 'TITLESNAME'
      Size = 50
    end
    object adoInv019INVTITLE: TStringField
      DisplayLabel = #25260#38957
      FieldName = 'INVTITLE'
      Size = 50
    end
    object adoInv019BUSINESSID: TStringField
      DisplayLabel = #32113#32232
      FieldName = 'BUSINESSID'
      Size = 8
    end
    object adoInv019ZIPCODE: TStringField
      DisplayLabel = #37109#36958#21312#34399
      FieldName = 'ZIPCODE'
      Size = 5
    end
    object adoInv019MAILADDRESS: TStringField
      DisplayLabel = #37109#23492#22320#22336
      FieldName = 'MAILADDRESS'
      Size = 60
    end
    object adoInv019INVADDRESS: TStringField
      DisplayLabel = #30332#31080#22320#22336
      FieldName = 'INVADDRESS'
      Size = 60
    end
    object adoInv019MEMO: TStringField
      DisplayLabel = #20633#35387
      FieldName = 'MEMO'
      Size = 60
    end
  end
  object adoInv099: TADOQuery
    Connection = dtmMain.InvConnection
    CursorType = ctStatic
    OnCalcFields = adoInv099CalcFields
    Parameters = <>
    SQL.Strings = (
      'SELECT YEARMONTH,'
      '       INVFORMAT,'
      '       PREFIX,'
      '       STARTNUM,'
      '       ENDNUM,'
      '       CURNUM,'
      '       TO_CHAR( LASTINVDATE,'#39'YYYY/MM/DD'#39' ) LInvDate,'
      '       MEMO '
      '  FROM INV099 '
      ' WHERE YEARMONTH IS NULL '
      '   AND USEFUL = '#39'Y'#39)
    Left = 136
    Top = 164
    object adoInv099YEARMONTH: TStringField
      FieldName = 'YEARMONTH'
      Size = 6
    end
    object adoInv099INVFORMAT: TStringField
      FieldName = 'INVFORMAT'
      FixedChar = True
      Size = 1
    end
    object adoInv099PREFIX: TStringField
      FieldName = 'PREFIX'
      Size = 2
    end
    object adoInv099STARTNUM: TStringField
      FieldName = 'STARTNUM'
      Size = 8
    end
    object adoInv099ENDNUM: TStringField
      FieldName = 'ENDNUM'
      Size = 8
    end
    object adoInv099CURNUM: TStringField
      FieldName = 'CURNUM'
      Size = 8
    end
    object adoInv099LINVDATE: TStringField
      FieldName = 'LINVDATE'
      ReadOnly = True
      Size = 10
    end
    object adoInv099MEMO: TStringField
      FieldName = 'MEMO'
      Size = 50
    end
    object adoInv099InvFormatDesc: TStringField
      DisplayLabel = #26684#24335
      FieldKind = fkCalculated
      FieldName = 'InvFormatDesc'
      Calculated = True
    end
  end
  object adoInv007: TADOQuery
    CacheSize = 1000
    Connection = dtmMain.InvConnection
    CursorType = ctStatic
    Parameters = <>
    Prepared = True
    SQL.Strings = (
      'SELECT INVID,'
      '       INVDATE,'
      '       CUSTID,'
      '       CUSTSNAME,'
      '       CHARGEDATE,'
      '       INVTITLE,'
      '       ZIPCODE,'
      '       INVADDR,'
      '       MAILADDR,'
      '       BUSINESSID,'
      
        '       DECODE( INVFORMAT, '#39'1'#39', '#39#38651#23376#39', '#39'2'#39', '#39#25163#20108#39', '#39'3'#39', '#39#25163#19977#39', INVFO' +
        'RMAT ) AS INVFORMAT,'
      
        '       DECODE( TAXTYPE, '#39'1'#39', '#39#25033#31237#39', '#39'2'#39', '#39#38646#31237#29575#39', '#39'3'#39', '#39#20813#31237#39', TAXTYP' +
        'E ) AS TAXTYPE,'
      '       TAXRATE,'
      '       SALEAMOUNT,'
      '       TAXAMOUNT,'
      '       INVAMOUNT,'
      '       DECODE( ISOBSOLETE, '#39'Y'#39', '#39#26159#39', '#39#21542#39' ) AS ISOBSOLETE,'
      '       OBSOLETEREASON,'
      '       MEMO1,'
      '       MEMO2,'
      '       UPTTIME,'
      '       UPTEN'
      '  FROM INV007'
      ' WHERE 1 = 2 ')
    Left = 20
    Top = 164
  end
  object adoInv005: TADOQuery
    Connection = dtmMain.InvConnection
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      '')
    Left = 20
    Top = 100
    object adoInv005ITEMID: TIntegerField
      DisplayLabel = #20195#30908
      FieldName = 'ITEMID'
    end
    object adoInv005DESCRIPTION: TStringField
      DisplayLabel = #21697#21517
      DisplayWidth = 20
      FieldName = 'DESCRIPTION'
      Size = 40
    end
    object adoInv005TAXNAME: TStringField
      DisplayLabel = #31237#29575#21029
      FieldName = 'TAXNAME'
      Size = 4
    end
    object adoInv005SIGN: TStringField
      DisplayLabel = #25910#20837'/'#25903#20986
      FieldName = 'SIGN'
      Visible = False
      Size = 1
    end
  end
  object adoInv011: TADOQuery
    Connection = dtmMain.InvConnection
    CursorType = ctStatic
    Parameters = <>
    Left = 80
    Top = 100
  end
  object adoInv013: TADOQuery
    Connection = dtmMain.InvConnection
    CursorType = ctStatic
    Parameters = <>
    Left = 80
    Top = 220
  end
  object adoA02_Master: TADOQuery
    CacheSize = 100
    Connection = dtmMain.InvConnection
    LockType = ltBatchOptimistic
    Parameters = <>
    Left = 144
    Top = 284
  end
  object adoA02_Detail: TADOQuery
    CacheSize = 100
    Connection = dtmMain.InvConnection
    LockType = ltBatchOptimistic
    Parameters = <>
    Left = 144
    Top = 340
  end
  object adoEinv: TADOQuery
    CacheSize = 1000
    Connection = dtmMain.EInvConnection
    CursorType = ctStatic
    Parameters = <>
    Prepared = True
    Left = 258
    Top = 288
  end
  object adoInv014: TADOQuery
    CacheSize = 1000
    Connection = dtmMain.InvConnection
    CursorType = ctStatic
    Parameters = <>
    Prepared = True
    SQL.Strings = (
      'SELECT INVID,'
      '       INVDATE,'
      '       CUSTID,'
      '       CUSTSNAME,'
      '       CHARGEDATE,'
      '       INVTITLE,'
      '       ZIPCODE,'
      '       INVADDR,'
      '       MAILADDR,'
      '       BUSINESSID,'
      
        '       DECODE( INVFORMAT, '#39'1'#39', '#39#38651#23376#39', '#39'2'#39', '#39#25163#20108#39', '#39'3'#39', '#39#25163#19977#39', INVFO' +
        'RMAT ) AS INVFORMAT,'
      
        '       DECODE( TAXTYPE, '#39'1'#39', '#39#25033#31237#39', '#39'2'#39', '#39#38646#31237#29575#39', '#39'3'#39', '#39#20813#31237#39', TAXTYP' +
        'E ) AS TAXTYPE,'
      '       TAXRATE,'
      '       SALEAMOUNT,'
      '       TAXAMOUNT,'
      '       INVAMOUNT,'
      '       DECODE( ISOBSOLETE, '#39'Y'#39', '#39#26159#39', '#39#21542#39' ) AS ISOBSOLETE,'
      '       OBSOLETEREASON,'
      '       MEMO1,'
      '       MEMO2,'
      '       UPTTIME,'
      '       UPTEN'
      '  FROM INV007'
      ' WHERE 1 = 2 ')
    Left = 88
    Top = 162
  end
  object adoInv046: TADOQuery
    CacheSize = 100
    Connection = dtmMain.InvConnection
    LockType = ltBatchOptimistic
    Parameters = <>
    Left = 278
    Top = 356
  end
end
