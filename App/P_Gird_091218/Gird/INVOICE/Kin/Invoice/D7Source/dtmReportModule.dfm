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
    ReportOptions.CreateDate = 38580.669872141200000000
    ReportOptions.LastChange = 39402.452189004600000000
    ScriptLanguage = 'PascalScript'
    ScriptText.Strings = (
      'var'
      '  aChineseNumber: array [0..9] of String;'
      ''
      
        '{ --------------------------------------------------------------' +
        '---------------------- }'
      ''
      'function GetDescription(aIndex: Integer): String;'
      'var'
      '  aDescription: String;'
      '  aYear, aMonth, aDay: Integer;'
      '  aYear2, aMonth2, aDay2: Integer;'
      'begin'
      '  aYear := 0;'
      '  aMonth := 0;'
      '  aDay := 0;'
      '  {}'
      '  aYear2 := 0;'
      '  aMonth2 := 0;'
      '  aDay2 := 0;'
      '  {}'
      '  if ( aIndex = 1 ) then'
      '  begin'
      '    aDescription := <frxA05_1."'#21697#21517'('#19968')">;'
      '    if ( <frxA05_1."'#26377#25928#26085#36215'('#19968')"> <> 0 ) then'
      '      DecodeDate( <frxA05_1."'#26377#25928#26085#36215'('#19968')">, aYear, aMonth, aDay );'
      '    if ( <frxA05_1."'#26377#25928#26085#36804'('#19968')"> <> 0 ) then'
      
        '      DecodeDate( <frxA05_1."'#26377#25928#26085#36804'('#19968')">, aYear2, aMonth2, aDay2 )' +
        ';'
      '  end else'
      '  if ( aIndex = 2 ) then'
      '  begin'
      '    aDescription := <frxA05_1."'#21697#21517'('#20108')">;'
      '    if ( <frxA05_1."'#26377#25928#26085#36215'('#20108')"> <> 0 ) then'
      '      DecodeDate( <frxA05_1."'#26377#25928#26085#36215'('#20108')">, aYear, aMonth, aDay );'
      '    if ( <frxA05_1."'#26377#25928#26085#36804'('#20108')"> <> 0 ) then'
      
        '      DecodeDate( <frxA05_1."'#26377#25928#26085#36804'('#20108')">, aYear2, aMonth2, aDay2 )' +
        ';'
      '  end else'
      '  if ( aIndex = 3 ) then'
      '  begin'
      '    aDescription := <frxA05_1."'#21697#21517'('#19977')">;'
      '    if ( <frxA05_1."'#26377#25928#26085#36215'('#19977')"> <> 0 ) then'
      '      DecodeDate( <frxA05_1."'#26377#25928#26085#36215'('#19977')">, aYear, aMonth, aDay );'
      '    if ( <frxA05_1."'#26377#25928#26085#36804'('#19977')"> <> 0 ) then'
      
        '      DecodeDate( <frxA05_1."'#26377#25928#26085#36804'('#19977')">, aYear2, aMonth2, aDay2 )' +
        ';'
      '  end else'
      '  if ( aIndex = 4 ) then'
      '  begin'
      '    aDescription := <frxA05_1."'#21697#21517'('#22235')">;'
      '    if ( <frxA05_1."'#26377#25928#26085#36215'('#22235')"> <> 0 ) then'
      '      DecodeDate( <frxA05_1."'#26377#25928#26085#36215'('#22235')">, aYear, aMonth, aDay );'
      '    if ( <frxA05_1."'#26377#25928#26085#36804'('#22235')"> <> 0 ) then'
      
        '      DecodeDate( <frxA05_1."'#26377#25928#26085#36804'('#22235')">, aYear2, aMonth2, aDay2 )' +
        ';'
      '  end else'
      '  if ( aIndex = 5 ) then'
      '  begin'
      '    aDescription := <frxA05_1."'#21697#21517'('#20116')">;'
      '    if ( <frxA05_1."'#26377#25928#26085#36215'('#20116')"> <> 0 ) then'
      '      DecodeDate( <frxA05_1."'#26377#25928#26085#36215'('#20116')">, aYear, aMonth, aDay );'
      '    if ( <frxA05_1."'#26377#25928#26085#36215'('#20116')"> <> 0 ) then'
      
        '      DecodeDate( <frxA05_1."'#26377#25928#26085#36804'('#20116')">, aYear2, aMonth2, aDay2 )' +
        ';'
      '  end;'
      '  Result := aDescription;'
      '  if ( aYear > 0 ) then'
      
        '    Result := ( Result + Format( '#39'(%d/%d/%d'#39', [aYear-1911, aMont' +
        'h+0, aDay+0] ) );'
      '  if ( aYear2 > 0 ) then'
      
        '    Result := ( Result + '#39'-'#39' + Format( '#39'%d/%d/%d)'#39', [aYear2-1911' +
        ', aMonth2+0, aDay2+0] ) );'
      'end;'
      ''
      
        '{ --------------------------------------------------------------' +
        '---------------------- }'
      ''
      'procedure MasterData1OnBeforePrint(Sender: TfrxComponent);'
      'var'
      '  aYear, aMonth, aDay: Integer;'
      '  aIndex, aNum: Integer;'
      '  aObject: TfrxMemoView;'
      '  aStrNum: String;'
      'begin'
      ''
      '  { '#26159#21542#21015#21360#30332#31080#25260#38957' }'
      
        '  Memo5.Visible := ( Get( '#39#21015#21360#30332#31080#25260#38957#39' ) = '#39'Y'#39' ) or ( <frxA05_1."'#32113#19968#32232 +
        #34399'"> <> '#39#39' );'
      ''
      '  { '#26159#21542#21015#21360#30332#31080#22320' }'
      '  Memo1.Visible := False;  //( Get( '#39#21015#21360#30332#31080#22320#22336#39' ) = '#39'Y'#39' );'
      '  Memo3.Visible := False;   //( Get( '#39#21015#21360#30332#31080#22320#22336#39' ) = '#39'Y'#39' );'
      ''
      '  { '#26159#21542#21015#21360#20633#35387' }'
      '  Memo39.Visible := ( Get( '#39#21015#21360#20633#35387#39' ) = '#39'Y'#39' );'
      ''
      '  { '#32068#21512#25910#36027#21697#21517'  }'
      '  Set( '#39#21697#21517'('#19968')'#39', '#39#39#39#39' + GetDescription(1) + '#39#39#39#39' );'
      '  Set( '#39#21697#21517'('#20108')'#39', '#39#39#39#39' + GetDescription(2) + '#39#39#39#39' );'
      '  Set( '#39#21697#21517'('#19977')'#39', '#39#39#39#39' + GetDescription(3) + '#39#39#39#39' );'
      '  Set( '#39#21697#21517'('#22235')'#39', '#39#39#39#39' + GetDescription(4) + '#39#39#39#39' );'
      '  Set( '#39#21697#21517'('#20116')'#39', '#39#39#39#39' + GetDescription(5) + '#39#39#39#39' );'
      ''
      '  { '#30332#31080#26085#26399' }'
      '  DecodeDate( <frxA05_1."'#30332#31080#26085#26399'">, aYear, aMonth, aDay );'
      '  Set( '#39#30332#31080#26085'('#24180')'#39', Format( '#39'%d'#39', [aYear-1911] ) );'
      
        '  Set( '#39#30332#31080#26085'('#26376')'#39', Format( '#39'%d'#39', [aMonth+0] ) ); { '#19981#21152' 0 , '#21360#21040#34920#19978#26371#19981#27491#30906 +
        ', ??? }'
      
        '  Set( '#39#30332#31080#26085'('#26085')'#39', Format( '#39'%d'#39', [aDay+0] ) ); { '#19981#21152' 0 , '#21360#21040#34920#19978#26371#19981#27491#30906', ' +
        '??? }'
      ''
      '  { '#20108#32879#24335#19981#21360#31237#38989' }'
      '  if ( <frxA05_1."'#32113#19968#32232#34399'"> <> '#39#39' ) then'
      '  begin'
      '     Memo37.Memo.Text := '#39'[frxA05_1."'#31237#38989'"]'#39';'
      '     Memo37.DataField := '#39#31237#38989#39';'
      '     Memo37.DisplayFormat.Kind := fkNumeric;'
      '     Memo37.DisplayFormat.FormatStr := '#39'#,##0.##'#39';'
      '     {}'
      '     Memo33.Memo.Text := '#39'[frxA05_1."'#37559#21806#38989'"]'#39';'
      '     Memo33.DataField := '#39#37559#21806#38989#39';'
      '  end else'
      '  begin'
      '     Memo37.Memo.Text := '#39'**********'#39';'
      '     Memo37.DataField := '#39#39';'
      '     Memo37.DisplaYFormat.Kind := fkText;'
      '     Memo37.DisplayFormat.FormatStr := '#39#39';'
      '     {}'
      '     Memo33.Memo.Text := '#39'[frxA05_1."'#30332#31080#32317#37329#38989'"]'#39';'
      '     Memo33.Memo.DataField := '#39#30332#31080#32317#37329#38989#39';'
      '  end;'
      ''
      '  { '#23458#32232' }'
      '  Memo32.Visible := ( <frxA05_1."'#23458#25142#20195#30908'"> <> '#39#39' );'
      ''
      '  { '#31237#21029' }'
      '  Set( '#39#25033#31237#39', '#39#39#39#39#39#39' );'
      '  Set( '#39#38646#31237#29575#39', '#39#39#39#39#39#39' );'
      '  Set( '#39#20813#31237#39', '#39#39#39#39#39#39' );'
      ''
      '  if ( <frxA05_1."'#31237#21029'"> = '#39'1'#39' ) then'
      '    Set( '#39#25033#31237#39', '#39#39#39'V'#39#39#39'  )'
      '  else if ( <frxA05_1."'#31237#21029'"> ) = '#39'2'#39' then'
      '    Set( '#39#38646#31237#29575#39', '#39#39#39'V'#39#39#39' )'
      '  else if ( <frxA05_1."'#31237#21029'"> ) = '#39'3'#39' then'
      '    Set( '#39#20813#31237#39', '#39#39#39'V'#39#39#39' );'
      ''
      ''
      '  { '#30332#31080#37329#38989#36681#20013#25991#25976#23383' }'
      ''
      '  aStrNum := Format( '#39'%.8d'#39', [<frxA05_1."'#30332#31080#32317#37329#38989'">] );'
      '  for aIndex := 1 to Length( aStrNum ) do'
      '  begin'
      '    aNum := StrToInt( Copy( aStrNum, aIndex, 1 ) );'
      '    aObject := nil;'
      '    case aIndex of'
      '      1: Set( '#39#20191#33836#39', '#39#39#39#39' + aChineseNumber[aNum] + '#39#39#39#39' );'
      '      2: Set( '#39#20336#33836#39', '#39#39#39#39' + aChineseNumber[aNum] + '#39#39#39#39' );'
      '      3: Set( '#39#25342#33836#39', '#39#39#39#39' + aChineseNumber[aNum] + '#39#39#39#39' );'
      '      4: Set( '#39#33836#39', '#39#39#39#39' + aChineseNumber[aNum] + '#39#39#39#39' );'
      '      5: Set( '#39#20191#39', '#39#39#39#39' + aChineseNumber[aNum] + '#39#39#39#39' );'
      '      6: Set( '#39#20336#39', '#39#39#39#39' + aChineseNumber[aNum] + '#39#39#39#39' );'
      '      7: Set( '#39#25342#39', '#39#39#39#39' + aChineseNumber[aNum] + '#39#39#39#39' );'
      '      8: Set( '#39#20491#39', '#39#39#39#39' + aChineseNumber[aNum] + '#39#39#39#39' );'
      '    end;'
      '  end;'
      ''
      'end;'
      ''
      
        '{ --------------------------------------------------------------' +
        '---------------------- }'
      ''
      'begin'
      '  aChineseNumber[0] := '#39#38646#39';'
      '  aChineseNumber[1] := '#39#22777#39';'
      '  aChineseNumber[2] := '#39#36019#39';'
      '  aChineseNumber[3] := '#39#21443#39';'
      '  aChineseNumber[4] := '#39#32902#39';'
      '  aChineseNumber[5] := '#39#20237#39';'
      '  aChineseNumber[6] := '#39#38520#39';'
      '  aChineseNumber[7] := '#39#26578#39';'
      '  aChineseNumber[8] := '#39#25420#39';'
      '  aChineseNumber[9] := '#39#29590#39';'
      'end.')
    Left = 36
    Top = 24
    Datasets = <
      item
        DataSet = frxA05_1
        DataSetName = 'frxA05_1'
      end>
    Variables = <
      item
        Name = ' '#20027#35201#27396#20301
        Value = Null
      end
      item
        Name = #30332#31080#26085'('#24180')'
        Value = Null
      end
      item
        Name = #30332#31080#26085'('#26376')'
        Value = Null
      end
      item
        Name = #30332#31080#26085'('#26085')'
        Value = Null
      end
      item
        Name = #21697#21517'('#19968')'
        Value = Null
      end
      item
        Name = #21697#21517'('#20108')'
        Value = Null
      end
      item
        Name = #21697#21517'('#19977')'
        Value = Null
      end
      item
        Name = #21697#21517'('#22235')'
        Value = Null
      end
      item
        Name = #21697#21517'('#20116')'
        Value = Null
      end
      item
        Name = ' '#31237#21029
        Value = Null
      end
      item
        Name = #25033#31237
        Value = Null
      end
      item
        Name = #38646#31237#29575
        Value = Null
      end
      item
        Name = #20813#31237
        Value = Null
      end
      item
        Name = ' '#20013#25991#23383#37329#38989
        Value = Null
      end
      item
        Name = #20191#33836
        Value = Null
      end
      item
        Name = #20336#33836
        Value = Null
      end
      item
        Name = #25342#33836
        Value = Null
      end
      item
        Name = #33836
        Value = Null
      end
      item
        Name = #20191
        Value = Null
      end
      item
        Name = #20336
        Value = Null
      end
      item
        Name = #25342
        Value = Null
      end
      item
        Name = #20491
        Value = Null
      end
      item
        Name = ' '#20854#23427
        Value = Null
      end
      item
        Name = #21015#21360#30332#31080#25260#38957
        Value = Null
      end
      item
        Name = #21015#21360#30332#31080#22320#22336
        Value = Null
      end
      item
        Name = #21015#21360#20633#35387
        Value = Null
      end>
    Style = <>
    object Page1: TfrxReportPage
      PaperWidth = 215.000000000000000000
      PaperHeight = 139.000000000000000000
      PaperSize = 256
      LeftMargin = 10.000000000000000000
      RightMargin = 10.000000000000000000
      TopMargin = 1.000000000000000000
      BottomMargin = 1.000000000000000000
      object MasterData1: TfrxMasterData
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = #27161#26999#39636
        Font.Style = []
        Height = 495.118430000000000000
        ParentFont = False
        Top = 18.897650000000000000
        Width = 737.008350000000000000
        OnBeforePrint = 'MasterData1OnBeforePrint'
        DataSet = frxA05_1
        DataSetName = 'frxA05_1'
        RowCount = 0
        object Memo1: TfrxMemoView
          Left = 90.708720000000000000
          Top = 147.401670000000000000
          Width = 71.811070000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clBlack
          Font.Height = -13
          Font.Name = #27161#26999#39636
          Font.Style = []
          Memo.Strings = (
            '[frxA05_1."'#37109#36958#21312#34399'"]')
          ParentFont = False
        end
        object Memo3: TfrxMemoView
          Left = 124.724490000000000000
          Top = 147.401670000000000000
          Width = 154.960730000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #30332#31080#25260#38957
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          Font.Charset = CHINESEBIG5_CHARSET
          Font.Color = clBlack
          Font.Height = -15
          Font.Name = #27161#26999#39636
          Font.Style = []
          Memo.Strings = (
            '[frxA05_1."'#30332#31080#25260#38957'"]')
          ParentFont = False
        end
        object Memo4: TfrxMemoView
          Left = 90.708720000000000000
          Top = 102.047310000000000000
          Width = 151.181200000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #30332#31080#34399#30908
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          Memo.Strings = (
            '[frxA05_1."'#30332#31080#34399#30908'"]')
        end
        object Memo5: TfrxMemoView
          Left = 90.708720000000000000
          Top = 117.165430000000000000
          Width = 151.181200000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          Memo.Strings = (
            '[frxA05_1."'#30332#31080#25260#38957'"]')
        end
        object Memo6: TfrxMemoView
          Left = 90.708720000000000000
          Top = 132.283550000000000000
          Width = 151.181200000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #32113#19968#32232#34399
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          Memo.Strings = (
            '[frxA05_1."'#32113#19968#32232#34399'"]')
        end
        object Memo7: TfrxMemoView
          Left = 317.480520000000000000
          Top = 102.047310000000000000
          Width = 37.795300000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            '['#30332#31080#26085'('#24180')]')
        end
        object Memo8: TfrxMemoView
          Left = 370.393940000000000000
          Top = 102.047310000000000000
          Width = 45.354360000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            '['#30332#31080#26085'('#26376')]')
        end
        object Memo9: TfrxMemoView
          Left = 430.866156380000000000
          Top = 102.047310000000000000
          Width = 49.133890000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            '['#30332#31080#26085'('#26085')]')
        end
        object Memo10: TfrxMemoView
          Left = 472.441250000000000000
          Top = 132.283550000000000000
          Width = 136.063080000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #27298#26597#30908
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          Memo.Strings = (
            '[frxA05_1."'#27298#26597#30908'"]')
        end
        object Memo11: TfrxMemoView
          Left = 30.236240000000000000
          Top = 200.315090000000000000
          Width = 79.370130000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            '['#21697#21517'('#19968')]')
        end
        object Memo12: TfrxMemoView
          Left = 30.236240000000000000
          Top = 219.212740000000000000
          Width = 79.370130000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            '['#21697#21517'('#20108')]')
        end
        object Memo13: TfrxMemoView
          Left = 30.236240000000000000
          Top = 238.110390000000000000
          Width = 79.370130000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            '['#21697#21517'('#19977')]')
        end
        object Memo14: TfrxMemoView
          Left = 30.236240000000000000
          Top = 257.008040000000000000
          Width = 79.370130000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            '['#21697#21517'('#22235')]')
        end
        object Memo15: TfrxMemoView
          Left = 30.236240000000000000
          Top = 275.905690000000000000
          Width = 79.370130000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          Memo.Strings = (
            '['#21697#21517'('#20116')]')
        end
        object Memo16: TfrxMemoView
          Left = 309.921460000000000000
          Top = 200.315090000000000000
          Width = 56.692913390000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #25976#37327'('#19968')'
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#25976#37327'('#19968')"]')
        end
        object Memo17: TfrxMemoView
          Left = 309.921460000000000000
          Top = 219.212740000000000000
          Width = 56.692913390000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #25976#37327'('#20108')'
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#25976#37327'('#20108')"]')
        end
        object Memo18: TfrxMemoView
          Left = 309.921460000000000000
          Top = 238.110390000000000000
          Width = 56.692913390000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #25976#37327'('#19977')'
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#25976#37327'('#19977')"]')
        end
        object Memo19: TfrxMemoView
          Left = 309.921460000000000000
          Top = 257.008040000000000000
          Width = 56.692913390000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #25976#37327'('#22235')'
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#25976#37327'('#22235')"]')
        end
        object Memo20: TfrxMemoView
          Left = 309.921460000000000000
          Top = 275.905690000000000000
          Width = 56.692913390000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #25976#37327'('#20116')'
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#25976#37327'('#20116')"]')
        end
        object Memo21: TfrxMemoView
          Left = 400.630180000000000000
          Top = 200.315090000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #21934#20729'('#19968')'
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          DisplayFormat.DecimalSeparator = '.'
          DisplayFormat.FormatStr = '#,##0.##'
          DisplayFormat.Kind = fkNumeric
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#21934#20729'('#19968')"]')
        end
        object Memo22: TfrxMemoView
          Left = 400.630180000000000000
          Top = 219.212740000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #21934#20729'('#20108')'
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          DisplayFormat.DecimalSeparator = '.'
          DisplayFormat.FormatStr = '#,##0.##'
          DisplayFormat.Kind = fkNumeric
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#21934#20729'('#20108')"]')
        end
        object Memo23: TfrxMemoView
          Left = 400.630180000000000000
          Top = 238.110390000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #21934#20729'('#19977')'
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          DisplayFormat.DecimalSeparator = '.'
          DisplayFormat.FormatStr = '#,##0.##'
          DisplayFormat.Kind = fkNumeric
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#21934#20729'('#19977')"]')
        end
        object Memo24: TfrxMemoView
          Left = 400.630180000000000000
          Top = 257.008040000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #21934#20729'('#22235')'
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          DisplayFormat.DecimalSeparator = '.'
          DisplayFormat.FormatStr = '#,##0.##'
          DisplayFormat.Kind = fkNumeric
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#21934#20729'('#22235')"]')
        end
        object Memo25: TfrxMemoView
          Left = 400.630180000000000000
          Top = 275.905690000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #21934#20729'('#20116')'
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          DisplayFormat.DecimalSeparator = '.'
          DisplayFormat.FormatStr = '#,##0.##'
          DisplayFormat.Kind = fkNumeric
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#21934#20729'('#20116')"]')
        end
        object Memo26: TfrxMemoView
          Left = 495.118430000000000000
          Top = 200.315090000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #32317#37329#38989'('#19968')'
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          DisplayFormat.DecimalSeparator = '.'
          DisplayFormat.FormatStr = '#,##0.##'
          DisplayFormat.Kind = fkNumeric
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#32317#37329#38989'('#19968')"]')
        end
        object Memo27: TfrxMemoView
          Left = 495.118430000000000000
          Top = 219.212740000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #32317#37329#38989'('#20108')'
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          DisplayFormat.DecimalSeparator = '.'
          DisplayFormat.FormatStr = '#,##0.##'
          DisplayFormat.Kind = fkNumeric
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#32317#37329#38989'('#20108')"]')
        end
        object Memo28: TfrxMemoView
          Left = 495.118430000000000000
          Top = 238.110390000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #32317#37329#38989'('#19977')'
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          DisplayFormat.DecimalSeparator = '.'
          DisplayFormat.FormatStr = '#,##0.##'
          DisplayFormat.Kind = fkNumeric
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#32317#37329#38989'('#19977')"]')
        end
        object Memo29: TfrxMemoView
          Left = 495.118430000000000000
          Top = 257.008040000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #32317#37329#38989'('#22235')'
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          DisplayFormat.DecimalSeparator = '.'
          DisplayFormat.FormatStr = '#,##0.##'
          DisplayFormat.Kind = fkNumeric
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#32317#37329#38989'('#22235')"]')
        end
        object Memo30: TfrxMemoView
          Left = 495.118430000000000000
          Top = 275.905690000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #32317#37329#38989'('#20116')'
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          DisplayFormat.DecimalSeparator = '.'
          DisplayFormat.FormatStr = '#,##0.##'
          DisplayFormat.Kind = fkNumeric
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#32317#37329#38989'('#20116')"]')
        end
        object Memo31: TfrxMemoView
          Left = 600.945270000000000000
          Top = 200.315090000000000000
          Width = 41.574830000000000000
          Height = 18.897650000000000000
          Memo.Strings = (
            #23458#32232':')
        end
        object Memo32: TfrxMemoView
          Left = 650.079160000000000000
          Top = 200.315090000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #23458#25142#20195#30908
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          Memo.Strings = (
            '[frxA05_1."'#23458#25142#20195#30908'"]')
        end
        object Memo33: TfrxMemoView
          Left = 491.338900000000000000
          Top = 359.055350000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #37559#21806#38989
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          DisplayFormat.DecimalSeparator = '.'
          DisplayFormat.FormatStr = '#,##0.##'
          DisplayFormat.Kind = fkNumeric
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#37559#21806#38989'"]')
        end
        object Memo34: TfrxMemoView
          Left = 340.157700000000000000
          Top = 400.630180000000000000
          Width = 37.795300000000000000
          Height = 18.897650000000000000
          HAlign = haCenter
          Memo.Strings = (
            '['#25033#31237']')
        end
        object Memo35: TfrxMemoView
          Left = 385.512060000000000000
          Top = 400.630180000000000000
          Width = 37.795300000000000000
          Height = 18.897650000000000000
          HAlign = haCenter
          Memo.Strings = (
            '['#38646#31237#29575']')
        end
        object Memo36: TfrxMemoView
          Left = 427.086890000000000000
          Top = 400.630180000000000000
          Width = 37.795300000000000000
          Height = 18.897650000000000000
          HAlign = haCenter
          Memo.Strings = (
            '['#20813#31237']')
        end
        object Memo37: TfrxMemoView
          Left = 517.795610000000000000
          Top = 389.291590000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          DisplayFormat.DecimalSeparator = '.'
          DisplayFormat.Kind = fkNumeric
          HAlign = haCenter
          Memo.Strings = (
            '[frxA05_1."'#31237#38989'"]')
        end
        object Memo38: TfrxMemoView
          Left = 491.338900000000000000
          Top = 419.527830000000000000
          Width = 94.488250000000000000
          Height = 18.897650000000000000
          AutoWidth = True
          DataField = #30332#31080#32317#37329#38989
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          DisplayFormat.DecimalSeparator = '.'
          DisplayFormat.FormatStr = '#,##0.##'
          DisplayFormat.Kind = fkNumeric
          HAlign = haRight
          Memo.Strings = (
            '[frxA05_1."'#30332#31080#32317#37329#38989'"]')
        end
        object Memo39: TfrxMemoView
          Left = 600.945270000000000000
          Top = 222.992270000000000000
          Width = 147.401670000000000000
          Height = 18.897650000000000000
          StretchMode = smActualHeight
          DataField = #20633#35387
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          Memo.Strings = (
            '[frxA05_1."'#20633#35387'"]')
        end
        object Memo40: TfrxMemoView
          Left = 120.944960000000000000
          Top = 445.984540000000000000
          Width = 41.574830000000000000
          Height = 18.897650000000000000
          HAlign = haCenter
          Memo.Strings = (
            '['#20191#33836']')
        end
        object Memo41: TfrxMemoView
          Left = 180.337574290000000000
          Top = 445.984540000000000000
          Width = 41.574830000000000000
          Height = 18.897650000000000000
          HAlign = haCenter
          Memo.Strings = (
            '['#20336#33836']')
        end
        object Memo42: TfrxMemoView
          Left = 239.730188570000000000
          Top = 445.984540000000000000
          Width = 41.574830000000000000
          Height = 18.897650000000000000
          HAlign = haCenter
          Memo.Strings = (
            '['#25342#33836']')
        end
        object Memo43: TfrxMemoView
          Left = 299.122802860000000000
          Top = 445.984540000000000000
          Width = 41.574830000000000000
          Height = 18.897650000000000000
          HAlign = haCenter
          Memo.Strings = (
            '['#33836']')
        end
        object Memo44: TfrxMemoView
          Left = 358.515417140000000000
          Top = 445.984540000000000000
          Width = 41.574830000000000000
          Height = 18.897650000000000000
          HAlign = haCenter
          Memo.Strings = (
            '['#20191']')
        end
        object Memo45: TfrxMemoView
          Left = 414.128501430000000000
          Top = 445.984540000000000000
          Width = 41.574830000000000000
          Height = 18.897650000000000000
          HAlign = haCenter
          Memo.Strings = (
            '['#20336']')
        end
        object Memo46: TfrxMemoView
          Left = 469.741585710000000000
          Top = 445.984540000000000000
          Width = 41.574830000000000000
          Height = 18.897650000000000000
          HAlign = haCenter
          Memo.Strings = (
            '['#25342']')
        end
        object Memo47: TfrxMemoView
          Left = 529.134200000000000000
          Top = 445.984540000000000000
          Width = 41.574830000000000000
          Height = 18.897650000000000000
          HAlign = haCenter
          Memo.Strings = (
            '['#20491']')
        end
        object Memo2: TfrxMemoView
          Left = 68.031540000000000000
          Top = 26.456710000000000000
          Width = 578.268090000000000000
          Height = 22.677180000000000000
          DataField = #37109#36958#21312#34399
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -19
          Font.Name = #27161#26999#39636
          Font.Style = []
          Memo.Strings = (
            '[frxA05_1."'#37109#36958#21312#34399'"]')
          ParentFont = False
        end
        object Memo48: TfrxMemoView
          Left = 139.842610000000000000
          Top = 26.456710000000000000
          Width = 801.260360000000000000
          Height = 18.897650000000000000
          DataField = #37109#23492#22320#22336
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -19
          Font.Name = #27161#26999#39636
          Font.Style = []
          Memo.Strings = (
            '[frxA05_1."'#37109#23492#22320#22336'"]')
          ParentFont = False
        end
        object Memo49: TfrxMemoView
          Left = 68.031540000000000000
          Top = 49.133889999999990000
          Width = 801.260360000000000000
          Height = 18.897650000000000000
          DataField = #23458#25142#31777#31281
          DataSet = frxA05_1
          DataSetName = 'frxA05_1'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -19
          Font.Name = #27161#26999#39636
          Font.Style = []
          Memo.Strings = (
            '[frxA05_1."'#23458#25142#31777#31281'"]')
          ParentFont = False
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
end
