unit dtmSOU;

interface

uses
  SysUtils, Classes, DB, ADODB, dtmMainU;

type
  TdtmSO = class(TDataModule)
    SOConnection: TADOConnection;
    adoComm: TADOQuery;
    adoQrySynSO041: TADOQuery;
    adoQrySynCD019: TADOQuery;
    adoQrySynCD032: TADOQuery;
    adoQrySynCD046: TADOQuery;
    adoQrySynSO001: TADOQuery;
    adoQrySynSO002: TADOQuery;
    adoMdCustomer: TADOQuery;
    adoChargeInfo34: TADOQuery;
    adoChargeInfo33: TADOQuery;
    adoA08: TADOQuery;
    adoQrySynCD095: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure ActiveMdCustomer(const aCustId: String);
    procedure ActiveMdCustomerEx(const aCustId: String);
    procedure GetChargeInfo(const aCustId: String; const aBillId: String = '';
      const aItemId: String = ''; const aTaxCode: String = ''; const aTaxRate: String = '');
    procedure DropInvData(var aParam: TCommonParam);
    procedure RecoverInvData(var aParam: TCommonParam);
    procedure GetMisAccountAndFacisNo(const aBillNo, aItem, aCustId: String;
      var aAcctNo, aFacisNo: String);
    procedure CreateAllowance(const aBillNo, aItem, aCustId, aNote: String);
    procedure DropAllowance(const aBillNo, aItem, aCustId, aAllowanceNo: String);
    function CheckAllowacneExists(const aBillNo, aItem, aCustId: String): Boolean;
  end;

var
  dtmSO: TdtmSO;

implementation


{$R *.dfm}

{ TdtmSO }

{ ---------------------------------------------------------------------------- }

procedure TdtmSO.DataModuleCreate(Sender: TObject);
begin
  {...}
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmSO.DataModuleDestroy(Sender: TObject);
begin
  {...}
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmSO.ActiveMdCustomer(const aCustId: String);
var
  aOwner: String;
begin
  adoMdCustomer.Close;
  if dtmMain.GetLinkToMIS then
  begin
    aOwner := dtmMain.getMisDbOwner;
    adoMdCustomer.Close;
    adoMdCustomer.SQL.Text := Format(
     '    SELECT ''SO002'' SOURCE,                 ' +
     '           A.CUSTNAME,                       ' +
     '           A.COMPCODE,                       ' +
     '           A.CUSTID,                         ' +
     '           A.MAILADDRESS,                    ' +
     '           C.SERVICETYPE,                    ' +
     '           C.INVTITLE,                       ' +
     '           C.INVNO BUSINESSID,               ' +
     '           C.INVADDRESS,                     ' +
     '           A.INSTADDRESS,                    ' +
     '           C.ACCOUNTNO,                      ' +
     '           B.ZIPCODE                         ' +
     '      FROM %s.SO001 A, %s.SO014 B, %s.SO002 C, %s.SO014 D ' +
     '     WHERE A.MAILADDRNO = B.ADDRNO           ' +
     '       AND A.INSTADDRNO = D.ADDRNO           ' +
     '       AND A.CUSTID = C.CUSTID               ' +
     '       AND A.CUSTID =                        ' + QuotedStr( aCustId ) +
     '       AND C.SERVICETYPE IN ( ' + dtmMain.GetServiceTypeStrEx + ' ) ' +
     ' UNION ALL                                   ' +
     '    SELECT ''SO002A'' SOURCE,                ' +
     '           A.CUSTNAME,                       ' +
     '           A.COMPCODE,                       ' +
     '           A.CUSTID,                         ' +
     '           A.MAILADDRESS,                    ' +
     '           null,                             ' +
     '           C.INVTITLE,                       ' +
     '           C.INVNO BUSINESSID,               ' +
     '           C.INVADDRESS,                     ' +
     '           A.INSTADDRESS,                    ' +
     '           C.ACCOUNTNO,                      ' +
     '           B.ZIPCODE                         ' +
     '      FROM %s.SO001 A, %s.SO014 B, %s.SO002A C, %s.SO014 D ' +
     '     WHERE A.MAILADDRNO = B.ADDRNO           ' +
     '       AND A.INSTADDRNO = D.ADDRNO           ' +
     '       AND A.CUSTID = C.CUSTID               ' +
     '       AND A.CUSTID =                        ' + QuotedStr( aCustId ) +
     '       AND C.STOPFLAG = 0                    ',
     [aOwner, aOwner, aOwner, aOwner, aOwner, aOwner, aOwner, aOwner] );
     adoMdCustomer.Open;
  end;
end;

{ ---------------------------------------------------------------------------- }

{ �s�����o�����Y - �h�b��, �o�����Y��Ʋ��� SO138,
  �� SO002A <-> SO002AD ������ SO138 }
procedure TdtmSO.ActiveMdCustomerEx(const aCustId: String);
var
  aOwner: String;
begin
  adoMdCustomer.Close;
  if dtmMain.GetLinkToMIS then
  begin
    aOwner := dtmMain.getMisDbOwner;
    adoMdCustomer.Close;
    adoMdCustomer.SQL.Text := Format(
     '    SELECT ''SO002'' SOURCE,                 ' +
     '           A.CUSTNAME,                       ' +
     '           A.COMPCODE,                       ' +
     '           A.CUSTID,                         ' +
     '           A.MAILADDRESS,                    ' +
     '           C.SERVICETYPE,                    ' +
     '           C.INVTITLE,                       ' +
     '           C.INVNO BUSINESSID,               ' +
     '           C.INVADDRESS,                     ' +
     '           A.INSTADDRESS,                    ' +
     '           C.ACCOUNTNO,                      ' +
     '           B.ZIPCODE,                        ' +
     '           null CHARGETITLE,                 ' +
     '           null INVSEQNO,                    ' +
     '           C.INVPURPOSECODE,                 ' +
     '           C.INVPURPOSENAME                  ' +
     '      FROM %s.SO001 A, %s.SO014 B, %s.SO002 C, %s.SO014 D ' +
     '     WHERE A.MAILADDRNO = B.ADDRNO           ' +
     '       AND A.INSTADDRNO = D.ADDRNO           ' +
     '       AND A.CUSTID = C.CUSTID               ' +
     '       AND A.CUSTID =                        ' + QuotedStr( aCustId ) +
     '       AND C.SERVICETYPE IN ( ' + dtmMain.GetServiceTypeStrEx + ' ) ' +
     ' UNION ALL                                   ' +
     '    SELECT UNIQUE                            ' +
     '      ''SO002A'' SOURCE,                     ' +
     '      A.CUSTNAME,                            ' +
     '      A.COMPCODE,                            ' +
     '      A.CUSTID,                              ' +
     '      D.MAILADDRESS,                         ' +
     '      null,                                  ' +
     '      D.INVTITLE,                            ' +
     '      D.INVNO BUSINESSID,                    ' +
     '      D.INVADDRESS,                          ' +
     '      A.INSTADDRESS,                         ' +
     '      C.ACCOUNTNO,                           ' +
     '      B.ZIPCODE,                             ' +
     '      D.CHARGETITLE,                         ' +
     '      D.INVSEQNO,                            ' +
     '      D.INVPURPOSECODE,                      ' +
     '      D.INVPURPOSENAME                       ' +
     ' FROM %s.SO001 A, %s.SO014 B, %s.SO002AD C, %s.SO138 D, %s.SO014 E ' +
     'WHERE D.MAILADDRNO = B.ADDRNO                ' +
     '  AND A.INSTADDRNO = E.ADDRNO                ' +
     '  AND A.CUSTID = C.CUSTID                    ' +
     '  AND C.INVSEQNO = D.INVSEQNO                ' +
     '  AND A.CUSTID =                             ' + QuotedStr( aCustId ) +
     '  AND D.STOPFLAG = 0                         ',
     [aOwner, aOwner, aOwner, aOwner, aOwner, aOwner, aOwner, aOwner, aOwner] );
     adoMdCustomer.Open;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmSO.GetChargeInfo(const aCustId, aBillId, aItemId, aTaxCode,
  aTaxRate: String);
var
  aOwner: String;
begin
  adoChargeInfo33.Close;
  adoChargeInfo34.Close;
  if dtmMain.GetLinkToMIS then
  begin
    aOwner := dtmMain.getMisDbOwner;
    adoChargeInfo33.SQL.Text :=
      '  SELECT 33 SOURCE,                            ' +
      '         A.COMPCODE,                           ' +
      '         A.CUSTID,                             ' +
      '         A.BILLNO,                             ' +
      '         A.ITEM,                               ' +
      '         A.CITEMCODE,                          ' +
      '         A.CITEMNAME,                          ' +
      '         A.SHOULDDATE,                         ' +
      '         A.SHOULDAMT,                          ' +
      '         A.REALDATE,                           ' +
      '         A.REALAMT,                            ' +
      '         A.REALPERIOD,                         ' +
      '         A.REALSTARTDATE,                      ' +
      '         A.REALSTOPDATE,                       ' +
      '         A.CLCTNAME,                           ' +
      '         A.STNAME,                             ' +
      '         A.ACCOUNTNO,                          ' +
      '         A.SERVICETYPE,                        ' +
      '         C.RATE1,                              ' +
      '         B.TAXCODE,                            ' +
      '         C.DESCRIPTION,                        ' +
      '         A.INVSEQNO                            ' +
      '    FROM %s.SO033 A, %s.CD019 B, %s.CD033 C    ' +
      '   WHERE A.CITEMCODE = B.CODENO                ' +
      '     AND B.TAXCODE = C.CODENO                  ' +
      '     AND C.TAXFLAG = 1                         ' +
      '     AND A.CANCELFLAG = 0                      ' +
      '     AND ( A.SHOULDAMT + A.REALAMT <> 0 )      ' +
      '     AND A.GUINO IS NULL                       ' +
      '     AND A.INVOICETIME IS NULL                 ' +
      '     AND A.CUSTID = ' + QuotedStr( aCustID  ) +
      '     AND A.SERVICETYPE IN ( ' + dtmMain.GetServiceTypeStrEx + ' ) ';

    adoChargeInfo33.SQL.Text := Format( adoChargeInfo33.SQL.Text, [
       aOwner, aOwner, aOwner] );

    if ( aBillId <> EmptyStr ) and ( aItemId <> EmptyStr ) then
    begin
      adoChargeInfo33.SQL.Text := ( adoChargeInfo33.SQL.Text + Format(
       '  AND ( A.BILLNO NOT IN ( %s ) OR A.ITEM NOT IN ( %s ) ) ', [aBillId, aItemId] ) );
    end;

    if ( aTaxCode <> '' ) and ( aTaxRate <> '' ) then
    begin
      adoChargeInfo33.SQL.Text := ( adoChargeInfo33.SQL.Text + Format(
       '  AND B.TAXCODE = %s  AND C.RATE1 = %s ', [aTaxCode, aTaxRate] ) );
    end;

    adoChargeInfo33.Open;

    adoChargeInfo34.SQL.Text :=
      '  SELECT 34 SOURCE,                            ' +
      '         A.COMPCODE,                           ' +
      '         A.CUSTID,                             ' +
      '         A.BILLNO,                             ' +
      '         A.ITEM,                               ' +
      '         A.CITEMCODE,                          ' +
      '         A.CITEMNAME,                          ' +
      '         A.SHOULDDATE,                         ' +
      '         A.SHOULDAMT,                          ' +
      '         A.REALDATE,                           ' +
      '         A.REALAMT,                            ' +
      '         A.REALPERIOD,                         ' +
      '         A.REALSTARTDATE,                      ' +
      '         A.REALSTOPDATE,                       ' +
      '         A.CLCTNAME,                           ' +
      '         A.STNAME,                             ' +
      '         A.ACCOUNTNO,                          ' +
      '         A.SERVICETYPE,                        ' +
      '         C.RATE1,                              ' +
      '         B.TAXCODE,                            ' +
      '         C.DESCRIPTION,                        ' +
      '         A.INVSEQNO                            ' +
      '    FROM %s.SO034 A, %s.CD019 B, %s.CD033 C    ' +
      '   WHERE A.CITEMCODE = B.CODENO                ' +
      '     AND B.TAXCODE = C.CODENO                  ' +
      '     AND C.TAXFLAG = 1                         ' +
      '     AND A.CANCELFLAG = 0                      ' +
      '     AND ( A.SHOULDAMT + A.REALAMT <> 0 )      ' +
      '     AND A.GUINO IS NULL                       ' +
      '     AND A.INVOICETIME IS NULL                 ' +
      '     AND A.CUSTID = ' + QuotedStr( aCustID  ) +
      '     AND A.SERVICETYPE IN ( ' + dtmMain.GetServiceTypeStrEx + ' ) ';

    adoChargeInfo34.SQL.Text := Format( adoChargeInfo34.SQL.Text, [
       aOwner, aOwner, aOwner] );

    if ( aBillId <> EmptyStr ) and ( aItemId <> EmptyStr ) then
    begin
      adoChargeInfo34.SQL.Text := ( adoChargeInfo34.SQL.Text + Format(
       '  AND ( A.BILLNO NOT IN ( %s ) OR A.ITEM NOT IN ( %s ) ) ', [aBillId, aItemId] ) );
    end;

    if ( aTaxCode <> EmptyStr ) and ( aTaxRate <> EmptyStr ) then
    begin
      adoChargeInfo34.SQL.Text := ( adoChargeInfo34.SQL.Text + Format(
       '  AND B.TAXCODE = %s  AND C.RATE1 = %s ', [aTaxCode, aTaxRate] ) );
    end;

    adoChargeInfo34.Open;
    
   end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmSO.DropInvData(var aParam: TCommonParam);
begin
  adoComm.Close;
  adoComm.SQL.Text := Format(
    '   UPDATE %s.SO033                      ' +
    '      SET PREINVOICE = NULL,            ' +
    '          GUINO = NULL,                 ' +
    '          INVDATE = NULL,               ' +
    '          INVOICETIME = NULL,           ' +
    '          INVPURPOSECODE = NULL,        ' +
    '          INVPURPOSENAME = NULL         ' +        
    '    WHERE GUINO = ''%s''                ',
    [dtmMain.getMisDbOwner, aParam.InStr1] );
  adoComm.ExecSQL;
  {}
  adoComm.SQL.Text := Format(
    '   UPDATE %s.SO034                       ' +
    '      SET PREINVOICE = NULL,             ' +
    '          GUINO = NULL,                  ' +
    '          INVDATE = NULL,                ' +
    '          INVOICETIME = NULL,            ' +
    '          INVPURPOSECODE = NULL,         ' +
    '          INVPURPOSENAME = NULL          ' +
    '    WHERE GUINO = ''%s''                 ' +
    '      AND ( PREINVOICE IS NULL OR        ' +
    '            PREINVOICE IN ( 0, 1, 2 ) )  ',
    [dtmMain.getMisDbOwner, aParam.InStr1] );
  adoComm.ExecSQL;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmSO.RecoverInvData(var aParam: TCommonParam);
var
  aEffect: Integer;
begin
  { aParam.InStr1 = �o�����X
    aParam.InStr2 = �o�����
    aParam.InStr3 = ���O�渹
    aParam.InStr4 = ���O����
    aParam.InStr5 = PREINVOICE
    aParam.InStr6 = �Ƚs,
    aParam.InStr7 = �o���γ~�N�X
    aParam.InStr8 = �o���γ~�W�� }
  adoComm.Close;
  adoComm.SQL.Text := Format(
    '   UPDATE %s.SO033                                       ' +
    '      SET PREINVOICE = %s,                               ' +
    '          GUINO = ''%s'',                                ' +
    '          INVDATE = TO_DATE( ''%s'', ''YYYY/MM/DD'' ),   ' +
    '          INVOICETIME = SYSDATE,                         ' +
    '          INVPURPOSECODE = ''%s'',                       ' +
    '          INVPURPOSENAME = ''%s''                        ' +
    '    WHERE CUSTID = ''%s''                                ' +
    '      AND BILLNO = ''%s''                                ' +
    '      AND ITEM = ''%s''                                  ',
    [dtmMain.getMisDbOwner, aParam.InStr5, aParam.InStr1, aParam.InStr2,
     aParam.InStr7, aParam.InStr8,
     aParam.InStr6, aParam.InStr3, aParam.InStr4] );
  aEffect := adoComm.ExecSQL;
  {}
  if ( aEffect <= 0 ) then
  begin
    adoComm.SQL.Text := Format(
      '   UPDATE %s.SO034                                       ' +
      '      SET PREINVOICE = %s,                               ' +
      '          GUINO = ''%s'',                                ' +
      '          INVDATE = TO_DATE( ''%s'', ''YYYY/MM/DD'' ),   ' +
      '          INVOICETIME = SYSDATE,                         ' +
    '            INVPURPOSECODE = ''%s'',                       ' +
    '            INVPURPOSENAME = ''%s''                        ' +
      '    WHERE CUSTID = ''%s''                                ' +
      '      AND BILLNO = ''%s''                                ' +
      '      AND ITEM = ''%s''                                  ',
      [dtmMain.getMisDbOwner, aParam.InStr5, aParam.InStr1, aParam.InStr2,
       aParam.InStr7, aParam.InStr8,
       aParam.InStr6, aParam.InStr3, aParam.InStr4] );
    adoComm.ExecSQL;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmSO.GetMisAccountAndFacisNo(const aBillNo, aItem,
  aCustId: String; var aAcctNo, aFacisNo: String);
var
  aTableName, aServiceType, aAccId: String;
begin
  { ���H�Υd����CP���� }
  aAcctNo := EmptyStr;
  aFacisNo := EmptyStr;
  aTableName := EmptyStr;
  adoComm.Close;
  adoComm.SQL.Text := Format(
    '  SELECT ''SO033'' TABLENAME, COUNT(1) COUNTS  FROM SO033 ' +
    '   WHERE BILLNO = ''%s'' AND ITEM = ''%s''   ' +
    ' UNION ALL                                   ' +
    '  SELECT ''SO034'' TABLENAME, COUNT(1) COUNTS  FROM SO034 ' +
    '   WHERE BILLNO = ''%s'' AND ITEM = ''%s''   ',
    [aBillNo, aItem, aBillNo, aItem] );
  adoComm.Open;
  adoComm.First;
  while not adoComm.Eof do
  begin
    if ( adoComm.FieldByName( 'COUNTS' ).AsInteger > 0 ) then
    begin
      aTableName := adoComm.FieldByName( 'TABLENAME' ).AsString;
      Break;
    end;
    adoComm.Next; 
  end;
  adoComm.Close;
  if ( aTableName = EmptyStr ) then Exit;
  adoComm.SQL.Text := Format(
    ' SELECT A.ACCOUNTNO, A.FACISNO, A.SERVICETYPE, B.ID  ' +
    '   FROM %s A, SO002A B                               ' +
    '  WHERE A.CUSTID = B.CUSTID(+)                       ' +
    '    AND A.ACCOUNTNO = B.ACCOUNTNO(+)                 ' +
    '    AND A.CUSTID = ''%s''                            ' +
    '    AND A.BILLNO = ''%s''                            ' +
    '    AND A.ITEM = ''%s''                              ',
    [aTableName, aCustId, aBillNo, aItem] );
  adoComm.Open;
  aAcctNo := adoComm.FieldByName( 'ACCOUNTNO' ).AsString;
  aFacisNo := adoComm.FieldByName( 'FACISNO' ).AsString;
  aServiceType := adoComm.FieldByName( 'SERVICETYPE' ).AsString;
  aAccId := adoComm.FieldByName( 'ID' ).AsString;
  adoComm.Close;
  { �A�ȧO�� P, �~��CP���� }
  if ( aServiceType <> 'P' ) then aFacisNo := EmptyStr;
  { �b���O��1, ��ܬ��H�Υd, �~���H�Υd��}
  if ( aAccId <> '1' ) then aAcctNo := EmptyStr;
  { �H�Υd���X, �u���� 4 �X }
  if ( aAcctNo <> EmptyStr ) then
  begin
    if ( Length( aAcctNo ) > 4 ) then
      aAcctNo := Copy( aAcctNo, Length( aAcctNo ) - 3, 4 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmSO.CreateAllowance(const aBillNo, aItem, aCustId,
  aNote: String);
begin
  adoComm.Close;
  adoComm.SQL.Text := Format(
    ' update so034 set               ' +
    '        preinvoice = 4,         ' +
    '        invoicetime = sysdate,  ' +
    '        note = ''%s''           ' +
    '  where custid = ''%s''         ' +
    '    and billno = ''%s''         ' +
    '    and item = ''%s''           ' +
    '    and preinvoice = ''3''      ',
    [aNote, aCustId, aBillNo, aItem] );
  adoComm.ExecSQL;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmSO.DropAllowance(const aBillNo, aItem, aCustId, aAllowanceNo: String);
var
  aNote: String;
begin
  adoComm.Close;
  adoComm.SQL.Text := Format(
    ' select note from so034         ' +
    '  where custid = ''%s''         ' +
    '    and billno = ''%s''         ' +
    '    and item = ''%s''           ' +
    '    and preinvoice = ''4''      ',
    [aCustId, aBillNo, aItem] );
  adoComm.Open;
  aNote := adoComm.Fields[0].AsString;
  adoComm.Close;
  {}
  if ( aNote <> EmptyStr ) then aNote := ( aNote + ',' );
  aNote := aNote + Format( '�R�������渹:%s', [aAllowanceNo] );
  {}
  adoComm.SQL.Text := Format(
    ' update so034 set               ' +
    '        preinvoice = 3,         ' +
    '        invoicetime = sysdate,  ' +
    '        note = ''%s''           ' +
    '  where custid = ''%s''         ' +
    '    and billno = ''%s''         ' +
    '    and item = ''%s''           ' +
    '    and preinvoice = ''4''      ',
    [aNote, aCustId, aBillNo, aItem] );
  adoComm.ExecSQL;
end;

{ ---------------------------------------------------------------------------- }

function TdtmSO.CheckAllowacneExists(const aBillNo, aItem,
  aCustId: String): Boolean;
begin
  adoComm.SQL.Text := Format(
    ' select count(1) from so034     ' +
    '  where custid = ''%s''         ' +
    '    and billno = ''%s''         ' +
    '    and item = ''%s''           ' +
    '    and preinvoice = 4          ',
    [aCustId, aBillNo, aItem] );
  adoComm.Open;
  Result := ( adoComm.Fields[0].AsInteger > 0 );
  adoComm.Close;
end;

{ ---------------------------------------------------------------------------- }

end.
