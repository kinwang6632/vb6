unit dtmSOU;

interface

uses
  SysUtils, Classes, DB, ADODB, dtmMainU,StrUtils;

type TvarAry = array of String;
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
    adoQrySynCD110: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    function Split(const aSplitStr : String) : TvarAry;

  public
    { Public declarations }
    procedure ActiveMdCustomer(const aCustId: String);
    procedure ActiveMdCustomerEx(const aCustId: String);
    procedure GetChargeInfo(const aCustId: String; const aBillId: String = '';
      const aItemId: String = ''; const aTaxCode: String = ''; const aTaxRate: String = '');
    procedure DropInvData(var aParam: TCommonParam);
    procedure RecoverInvData(var aParam: TCommonParam);
    procedure GetMisAccountAndFacisNo(const aBillNo, aItem, aCustId: String;
      var aAcctNo, aFacisNo, aSmartCard: String);
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
  {#5668 �W�[ INVOICEKIND ��� By Kin 2010/06/29 }
  adoMdCustomer.Close;
  if dtmMain.GetLinkToMIS then
  begin
    aOwner := dtmMain.getMisDbOwner;
    adoMdCustomer.Close;
    adoMdCustomer.SQL.Text := Format(
     '    SELECT ''SO002'' SOURCE,                                                              ' +
     '           A.CUSTNAME,                                                                    ' +
     '           A.COMPCODE,                                                                    ' +
     '           A.CUSTID,                                                                      ' +
     '           A.MAILADDRESS,                                                                 ' +
     '           C.SERVICETYPE,                                                                 ' +
     '           C.INVTITLE,                                                                    ' +
     '           C.INVNO BUSINESSID,                                                            ' +
     '           C.INVADDRESS,                                                                  ' +
     '           A.INSTADDRESS,                                                                 ' +
     '           C.ACCOUNTNO,                                                                   ' +
     '           B.ZIPCODE,                                                                     ' +
     '           null CHARGETITLE,                                                              ' +
     '           null INVSEQNO,                                                                 ' +
     '           C.INVPURPOSECODE,                                                              ' +
     '           C.INVPURPOSENAME,                                                              ' +
     '           Decode(Nvl(C.INVOICEKIND,0),0,''�q�l�p����o��'',''�q�l�o��'') INVOICEKIND,    ' +
     '           C.DENRECCODE,C.DENRECNAME,                                                     ' +
     '           C.EMAIL,A.TEL3                                                                 ' +
     '      FROM %s.SO001 A, %s.SO014 B, %s.SO002 C, %s.SO014 D                                 ' +
     '     WHERE A.MAILADDRNO = B.ADDRNO                                                        ' +
     '       AND A.INSTADDRNO = D.ADDRNO                                                        ' +
     '       AND A.CUSTID = C.CUSTID                                                            ' +
     '       AND A.CUSTID =                                         ' + QuotedStr( aCustId )      +
     '       AND C.SERVICETYPE IN ( ' + dtmMain.GetServiceTypeStrEx + ' )                       ' +
     ' UNION ALL                                                                                ' +
     '    SELECT UNIQUE                                                                         ' +
     '      ''SO002A'' SOURCE,                                                                  ' +
     '      A.CUSTNAME,                                                                         ' +
     '      A.COMPCODE,                                                                         ' +
     '      A.CUSTID,                                                                           ' +
     '      D.MAILADDRESS,                                                                      ' +
     '      null,                                                                               ' +
     '      D.INVTITLE,                                                                         ' +
     '      D.INVNO BUSINESSID,                                                                 ' +
     '      D.INVADDRESS,                                                                       ' +
     '      A.INSTADDRESS,                                                                      ' +
     '      C.ACCOUNTNO,                                                                        ' +
     '      B.ZIPCODE,                                                                          ' +
     '      D.CHARGETITLE,                                                                      ' +
     '      D.INVSEQNO,                                                                         ' +
     '      D.INVPURPOSECODE,                                                                   ' +
     '      D.INVPURPOSENAME,                                                                   ' +
     '      DECODE(Nvl(D.INVOICEKIND,0),0,''�q�l�p����o��'',''�q�l�o��'') INVOICEKIND,         ' +
     '      D.DENRECCODE,D.DENRECNAME,                                                          ' +
     '      NULL EMAIL,A.TEL3                                                                   ' +
     ' FROM %s.SO001 A, %s.SO014 B, %s.SO002AD C, %s.SO138 D, %s.SO014 E                        ' +
     'WHERE D.MAILADDRNO = B.ADDRNO                                                             ' +
     '  AND A.INSTADDRNO = E.ADDRNO                                                             ' +
     '  AND A.CUSTID = C.CUSTID                                                                 ' +
     '  AND C.INVSEQNO = D.INVSEQNO                                                             ' +
     '  AND A.CUSTID =                                              ' + QuotedStr( aCustId )      +
     '  AND D.STOPFLAG = 0                              ',
     [aOwner, aOwner, aOwner, aOwner, aOwner, aOwner, aOwner, aOwner, aOwner] );
     adoMdCustomer.Open;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmSO.GetChargeInfo(const aCustId, aBillId, aItemId, aTaxCode,
  aTaxRate: String);
var
  aOwner,aWhere: String;
  aAryBill,aAryItem : TvarAry;
  aIndex : Integer;
begin
  adoChargeInfo33.Close;
  adoChargeInfo34.Close;
  {#5285 �P�_���B�n�Υ�����],�ҥH�h�W�[Uccode By Kin 2009/09/09}
  {#5439 �W�[���O��ƪ��X�֦��O���� By Kin }
  if dtmMain.GetLinkToMIS then
  begin
    aOwner := dtmMain.getMisDbOwner;
    adoChargeInfo33.SQL.Text :=
      '  SELECT 33 SOURCE,                                                                   ' +
      '         A.COMPCODE,                                                                  ' +
      '         A.CUSTID,                                                                    ' +
      '         A.BILLNO,                                                                    ' +
      '         A.ITEM,                                                                      ' +
      '         A.CITEMCODE,                                                                 ' +
      '         A.CITEMNAME,                                                                 ' +
      '         A.SHOULDDATE,                                                                ' +
      '         DECODE(NVL(A.COMBAMOUNT,0),0,A.SHOULDAMT,A.COMBAMOUNT) SHOULDAMT,            ' +
      '         A.REALDATE,                                                                  ' +
      '         DECODE(NVL(A.COMBAMOUNT,0),0,A.REALAMT,A.COMBAMOUNT) REALAMT,                ' +
      '         DECODE(A.COMBSTARTDATE,NULL,A.REALSTARTDATE,COMBSTARTDATE) REALSTARTDATE,    ' +
      '         DECODE(A.COMBSTOPDATE,NULL,A.REALSTOPDATE,A.COMBSTOPDATE) REALSTOPDATE,      ' +
      '         A.REALSTOPDATE,                                                              ' +
      '         A.CLCTNAME,                                                                  ' +
      '         A.STNAME,                                                                    ' +
      '         A.ACCOUNTNO,                                                                 ' +
      '         A.SERVICETYPE,                                                               ' +
      '         C.RATE1,                                                                     ' +
      '         B.TAXCODE,                                                                   ' +
      '         C.DESCRIPTION,                                                               ' +
      '         A.INVSEQNO ,                                                                 ' +
      '         A.UCCode,                                                                    ' +
      '         A.COMBCITEMCODE,                                                             ' +
      '         A.COMBCITEMNAME                                                              ' +
      '    FROM %s.SO033 A, %s.CD019 B, %s.CD033 C                                           ' +
      '   WHERE A.CITEMCODE = B.CODENO                                                       ' +
      '     AND B.TAXCODE = C.CODENO                                                         ' +
      '     AND C.TAXFLAG = 1                                                                ' +
      '     AND A.CANCELFLAG = 0                                                             ' +
      '     AND ( A.SHOULDAMT + A.REALAMT <> 0 )                                             ' +
      '     AND A.GUINO IS NULL                                                              ' +
      '     AND A.INVOICETIME IS NULL                                                        ' +
      '     AND A.CUSTID = ' + QuotedStr( aCustID  ) +
      '     AND A.SERVICETYPE IN ( ' + dtmMain.GetServiceTypeStrEx + ' ) ';

    adoChargeInfo33.SQL.Text := Format( adoChargeInfo33.SQL.Text, [
       aOwner, aOwner, aOwner] );

    if ( aBillId <> EmptyStr ) and ( aItemId <> EmptyStr ) then
    begin
      { #5014 �쥻��BillNo IN Item IN�|�����D,��y�k--��Rowid By Kin 2009/04/03 }
      aAryBill := Split(aBillId);
      aAryItem := Split(aItemId);
      aWhere := EmptyStr;
      for aIndex := Low(aAryBill) to High(aAryBill) do
      begin
        if ( aWhere <> EmptyStr ) then
          aWhere := aWhere + ' OR ';
        aWhere :=  aWhere + Format('( BILLNO = %s AND ITEM = %s )',[aAryBill[aIndex],aAryItem[aIndex]]);
      end;
      adoChargeInfo33.SQL.Text := ( adoChargeInfo33.SQL.Text +
          Format(' AND A.ROWID NOT IN ( SELECT ROWID FROM %s.SO033 Where %s )',[aOwner,aWhere]) );
      {
      adoChargeInfo33.SQL.Text := ( adoChargeInfo33.SQL.Text + Format(
       '  AND ( A.BILLNO NOT IN ( %s ) OR A.ITEM NOT IN ( %s ) ) ', [aBillId, aItemId] ) );
      }
    end;

    if ( aTaxCode <> '' ) and ( aTaxRate <> '' ) then
    begin
      adoChargeInfo33.SQL.Text := ( adoChargeInfo33.SQL.Text + Format(
       '  AND B.TAXCODE = %s  AND C.RATE1 = %s ', [aTaxCode, aTaxRate] ) );
    end;
    {#5285 �P�_���B�n�Υ�����],�ҥH�h�W�[Uccode By Kin 2009/09/09}
    adoChargeInfo33.Open;

    adoChargeInfo34.SQL.Text :=
      '  SELECT 34 SOURCE,                                                                   ' +
      '         A.COMPCODE,                                                                  ' +
      '         A.CUSTID,                                                                    ' +
      '         A.BILLNO,                                                                    ' +
      '         A.ITEM,                                                                      ' +
      '         A.CITEMCODE,                                                                 ' +
      '         A.CITEMNAME,                                                                 ' +
      '         A.SHOULDDATE,                                                                ' +
      '         A.SHOULDAMT,                                                                 ' +
      '         A.REALDATE,                                                                  ' +
      '         DECODE(NVL(A.COMBAMOUNT,0),0,A.REALAMT,A.COMBAMOUNT) REALAMT,                ' +
      '         A.REALPERIOD,                                                                ' +
      '         DECODE(A.COMBSTARTDATE,NULL,A.REALSTARTDATE,COMBSTARTDATE) REALSTARTDATE,    ' +
      '         DECODE(A.COMBSTOPDATE,NULL,A.REALSTOPDATE,A.COMBSTOPDATE) REALSTOPDATE,      ' +
      '         A.CLCTNAME,                                                                  ' +
      '         A.STNAME,                                                                    ' +
      '         A.ACCOUNTNO,                                                                 ' +
      '         A.SERVICETYPE,                                                               ' +
      '         C.RATE1,                                                                     ' +
      '         B.TAXCODE,                                                                   ' +
      '         C.DESCRIPTION,                                                               ' +
      '         A.INVSEQNO,                                                                  ' +
      '         A.UCCode,                                                                    ' +
      '         A.COMBCITEMCODE,                                                             ' +
      '         A.COMBCITEMNAME                                                              ' +
      '    FROM %s.SO034 A, %s.CD019 B, %s.CD033 C                                           ' +
      '   WHERE A.CITEMCODE = B.CODENO                                                       ' +
      '     AND B.TAXCODE = C.CODENO                                                         ' +
      '     AND C.TAXFLAG = 1                                                                ' +
      '     AND A.CANCELFLAG = 0                                                             ' +
      '     AND ( A.SHOULDAMT + A.REALAMT <> 0 )                                             ' +
      '     AND A.GUINO IS NULL                                                              ' +
      '     AND A.INVOICETIME IS NULL                                                        ' +
      '     AND A.CUSTID = ' + QuotedStr( aCustID  ) +
      '     AND A.SERVICETYPE IN ( ' + dtmMain.GetServiceTypeStrEx + ' ) ';

    adoChargeInfo34.SQL.Text := Format( adoChargeInfo34.SQL.Text, [
       aOwner, aOwner, aOwner] );

    if ( aBillId <> EmptyStr ) and ( aItemId <> EmptyStr ) then
    begin
      adoChargeInfo34.SQL.Text := ( adoChargeInfo34.SQL.Text +
          Format(' AND A.ROWID NOT IN ( SELECT ROWID FROM %s.SO034 Where %s )',[aOwner,aWhere]) )
      {
      adoChargeInfo34.SQL.Text := ( adoChargeInfo34.SQL.Text + Format(
       '  AND ( A.BILLNO NOT IN ( %s ) OR A.ITEM NOT IN ( %s ) ) ', [aBillId, aItemId] ) );
      }
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
  aCustId: String; var aAcctNo, aFacisNo, aSmartCard: String);
var
  aTableName, aServiceType, aAccId,aOwner: String;
begin
  { ���H�Υd����CP���� }
  aOwner := dtmMain.getMisDbOwner;
  if aOwner <> '' then
  begin
    if Pos('.' ,aOwner) <= 0 then
    begin
      aOwner := aOwner + '.'
    end;
  end;


  aAcctNo := EmptyStr;
  aFacisNo := EmptyStr;
  aTableName := EmptyStr;
  aSmartCard := EmptyStr;
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
  {#5668 �p�G�]�w��ܳ]��,�n�NSmartCardNo ��X�� By Kin 2010/07/09 }
  if ( aFacisNo <> EmptyStr ) and ( dtmMain.GetShowFaci = 1) then
  begin
    adoComm.SQL.Text := Format('SELECT SMARTCARDNO FROM %sSO004     ' +
                  ' WHERE FACISNO = ''%s'' AND CUSTID = %s          ' ,
                  [ aOwner,aFacisNo,aCustId ] );
    adoComm.Open;
    if adoComm.IsEmpty then
      aSmartCard := EmptyStr
    else
      aSmartCard := adoComm.FieldByName( 'SMARTCARDNO' ).AsString;
    adoComm.Close;
  end;
  { �A�ȧO�� P, �~��CP���� }
  { #5358 �W�[ShowFaci �P�_�b�}�ߵo���ɬO�_�n�N�]�ƧǸ��g�JMemo1 By Kin 2009/12/14 }
  if dtmMain.GetShowFaci = 0 then
  begin
    if ( aServiceType <> 'P' ) then aFacisNo := EmptyStr;
  end;

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
{ #5014 �N�r�꦳','���ά��}�C ��VB��Split By Kin 2009/04/03 }
function TdtmSO.Split(const aSplitStr : String): TvarAry;
var
  aStr: string;
  aSource: String;
  aIndex : Integer;
  aArry : TvarAry;
begin
  if Length(aSplitStr) = 0 then
  begin
    Result := nil;
  end else
  begin
    aStr := EmptyStr;
    aSource := aSplitStr;
    if Pos(',',aSource) > 0 then
    begin
      for aIndex := 1 to Length( aSource ) do
      begin
        if ( aSource[aIndex] <> ',' ) then
        begin
          aStr := ( aStr + aSource[aIndex] );
        end
        else begin
          SetLength( aArry, Length( aArry ) + 1 );
          aArry[ Length( aArry ) - 1 ] := aStr;
          aStr := '';
        end;
      end;
      if ( aStr <> '' ) then
      begin
        SetLength( aArry, Length( aArry ) + 1 );
        aArry[ Length( aArry ) - 1 ] := aStr;
        aStr := '';
      end;
    end else
    begin
      SetLength(aArry,1);
      aArry[ Length( aArry ) - 1 ] := aSource;
    end;
  end;
  Result := aArry;
end;

{ ---------------------------------------------------------------------------- }
end.
