
CREATE OR REPLACE function SF_TranInv1(p_CompCode number, 
      p_CustId1 number, p_CustId2 number, p_ClassSQL varchar2, 
      p_BillTypeSQL varchar2, p_CitemSQL varchar2, 
      p_RealDate1 varchar2, p_RealDate2 varchar2, 
      p_ClctEn varchar2, p_UpdDate1 varchar2, p_UpdDate2 varchar2, 
      p_UpdEn varchar2, p_CMCode number, p_InvType number, p_AddrType number, 
      p_RetMsg out varchar2) return number
  AS

  c39		char(1) := chr(39);
  v_StartExecTime date;
  v_SQL1 	varchar2(1000);
  v_SQL2 	varchar2(1000);
  v_Char  	char(1);
  v_RcdCount 	number := 0;
  v_Now 	date;

  v_Cursor	number;			-- ��dynamic SQL
  v_RetCode	number;
  v_CurrCustId 	number;
  v_CurrCitemCode number;
  v_TaxCode 	number;

  z_CustId 	number;
  z_BillNo	varchar2(15);
  z_RealDate	date;
  z_RealAmt	number;
  z_CitemCode	number;
  z_CitemName	varchar2(12);
  z_UpdTime	varchar2(19);
  z_RealStartDate date;
  z_RealStopDate date;
  z_Note	varchar2(250);
  z_Address	varchar2(60);
  z_AddrNo	number;

  z_InvNo	number;
  z_InvTitle	varchar2(50);
  z_InvAddress	varchar2(60);
  z_ZipCode	varchar2(5);
  z_TaxType	char(1);

  z_UpdA	date;
  z_UpdB	number;

  --e_Final exception;

BEGIN

  -- �Ѽ��ˬd
  if p_InvType is null or (p_InvType not in (1,2,3)) or 
    p_AddrType is null or (p_AddrType not in (1,2,3)) then
    p_RetMsg := '�Ѽƿ��~';
    return -1;
  end if;

  -- �R���U�Ȧs��
  delete from Tmp003;

  -- �]�w�U�ܼƪ��
  v_StartExecTime := sysdate;		-- �}�l����ɶ�
  v_SQL1 := null;

  -- �Y���Ȥ�s������
  if nvl(p_CustId1,0)>0 then
    v_SQL1 := v_SQL1 || ' and A.CustId>=' || p_CustId1 || ' and A.CustId<=' || p_CustId2;
  end if;  
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  -- �Y�����O���ر���
  if p_CitemSQL is not null then
    v_SQL1 := v_SQL1 || ' and A.CitemCode ' || p_CitemSQL;
  end if;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  -- �Y���Ȥ����O����
  if p_ClassSQL is not null then
    v_SQL1 := v_SQL1 || ' and A.ClassCode ' || p_ClassSQL;
  end if;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  -- �Y�����q�O����
  if p_CompCode is not null then
    v_SQL1 := v_SQL1 || ' and A.CompCode=' || p_CompCode;
  end if;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  -- �Y���J�b�������
  if p_RealDate1 is not null then
    v_SQL1 := v_SQL1 || ' and A.RealDate>='||'to_date('||chr(39)||p_RealDate1||chr(39)||','||chr(39)||'YYYYMMDD'||chr(39)||')'||
		' and A.RealDate<=to_date('||chr(39)||p_RealDate1||chr(39)||','||chr(39)||'YYYYMMDD'||chr(39)||')';
  end if;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  -- �Y�����O�H������
  if p_ClctEn is not null then
    v_SQL1 := v_SQL1 || ' and A.ClctEn=' || c39 || p_ClctEn || c39;
  end if;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  -- �Y�����ʤ������
  if p_UpdDate1 is not null then
    v_SQL1 := v_SQL1 || ' A.UpdTime>='||c39||ltrim(to_char(to_number(substr(p_UpdDate1,1,4))-1911,'99'))||
		'/'||substr(p_UpdDate1,5,2)||'/'||substr(p_UpdDate1,7,2)||c39||
		' and A.UpdTime<='||c39||ltrim(to_char(to_number(substr(p_UpdDate2,1,4))-1911,'99'))||
		'/'||substr(p_UpdDate2,5,2)||'/'||substr(p_UpdDate2,7,2)||' 23:59:59'||c39;
  end if;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  -- �Y�����ʤH������
  if p_UpdEn is not null then
    --v_SQL1 := v_SQL1 || ' and A.UpdEn=' || c39 || p_UpdEn || c39;
    v_SQL1 := ' and A.UpdEn=' || c39 || p_UpdEn || c39;
  end if;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  -- �Y�����O�覡����
  if p_CMCode is not null then
    v_SQL1 := v_SQL1 || ' and A.CMCode=' || p_CMCode;
  end if;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  -- �o������
  if p_InvType!=1 then
    v_SQL1 := v_SQL1 || ' and B.InvoiceType=' || p_InvType;
  else
    v_SQL1 := v_SQL1 || ' and nvl(B.InvoiceType,0)!=0';
  end if;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  -- �Y����ں�������
  if p_BillTypeSQL is not null then
    v_Char := substr(p_BillTypeSQL, 1, 1);
    if v_Char='=' or v_Char='!' then
      if p_BillTypeSQL like '%B%' or p_BillTypeSQL like '%T%'  then
	v_SQL1 := v_SQL1 || ' and substr(A.BillNo,7,1)'||p_BillTypeSQL;
      else
	v_SQL1 := v_SQL1 || ' and substr(A.BillNo,1,1)'||p_BillTypeSQL;
      end if;
    else
      if p_BillTypeSQL like '%B%' and p_BillTypeSQL like '%T%'  then
	v_SQL1 := v_SQL1 || ' and (substr(A.BillNo,7,1) '||p_BillTypeSQL||')';
      else
	v_SQL1 := v_SQL1 || ' and (substr(A.BillNo,1,1) '||p_BillTypeSQL||
		' or substr(A.BillNo,7,1) '||p_BillTypeSQL||')';
      end if;
    end if;
  end if;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  --v_SQL1 := ltrim(v_SQL1,' and');
  v_SQL2 := 'select A.CustId,A.BillNo,A.RealDate,A.RealAmt,A.CitemCode,A.CitemName,'||
		'A.UpdTime,A.RealStartDate,A.RealStopDate,A.Note,';
  if p_AddrType=1 then
    v_SQL2 := v_SQL2 || 'B.InstAddress Address, B.InstAddrNo AddrNo';
  elsif p_AddrType=2 then
    v_SQL2 := v_SQL2 || 'B.ChargeAddress Address, B.ChargeAddrNo AddrNo';
  else
    v_SQL2 := v_SQL2 || 'B.MailAddress Address, B.MailAddrNo AddrNo';
  end if;
  v_SQL1 := v_SQL2 || ' from SO034 A, SO001 B where A.CustId=B.CustId ' || 
	v_SQL1 || ' order by A.CustId';
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);

  -- ����Ӭd��
  begin
    v_Cursor := DBMS_SQL.Open_cursor;
    DBMS_SQL.Parse(v_Cursor, v_SQL1, DBMS_SQL.V7);

    -- �w�q���X�����
    DBMS_SQL.Define_Column(v_Cursor, 1, z_CustId);
    DBMS_SQL.Define_Column(v_Cursor, 2, z_BillNo, 15);
    DBMS_SQL.Define_Column(v_Cursor, 3, z_RealDate);
    DBMS_SQL.Define_Column(v_Cursor, 4, z_RealAmt);
    DBMS_SQL.Define_Column(v_Cursor, 5, z_CitemCode);
    DBMS_SQL.Define_Column(v_Cursor, 6, z_CitemName, 12);
    DBMS_SQL.Define_Column(v_Cursor, 7, z_UpdTime, 19);
    DBMS_SQL.Define_Column(v_Cursor, 8, z_RealStartDate);
    DBMS_SQL.Define_Column(v_Cursor, 9, z_RealStopDate);
    DBMS_SQL.Define_Column(v_Cursor, 10, z_Note, 250);
    DBMS_SQL.Define_Column(v_Cursor, 11, z_Address, 60);
    DBMS_SQL.Define_Column(v_Cursor, 12, z_AddrNo);

    v_RetCode := DBMS_SQL.Execute(v_Cursor);
    --DBMS_SQL.Close_cursor(v_Cursor);
  exception
    when others then
      DBMS_SQL.Close_cursor(v_Cursor);
      rollback;
      p_RetMsg := 'SQL1���~: ' || v_SQL1;
      return -2;
  end;

  v_CurrCustId := -1;
  v_CurrCitemCode := -1;

  loop					-- ���쪺�C�@�����
    if DBMS_SQL.Fetch_rows(v_Cursor) <= 0 then		-- Fetch 1 rcd
      exit;				-- No more rows to fetch
    end if;
    -- ���U����
    DBMS_SQL.Column_Value(v_Cursor, 1, z_CustId);
    DBMS_SQL.Column_Value(v_Cursor, 2, z_BillNo);
    DBMS_SQL.Column_Value(v_Cursor, 3, z_RealDate);
    DBMS_SQL.Column_Value(v_Cursor, 4, z_RealAmt);
    DBMS_SQL.Column_Value(v_Cursor, 5, z_CitemCode);
    DBMS_SQL.Column_Value(v_Cursor, 6, z_CitemName);
    DBMS_SQL.Column_Value(v_Cursor, 7, z_UpdTime);
    DBMS_SQL.Column_Value(v_Cursor, 8, z_RealStartDate);
    DBMS_SQL.Column_Value(v_Cursor, 9, z_RealStopDate);
    DBMS_SQL.Column_Value(v_Cursor, 10, z_Note);
    DBMS_SQL.Column_Value(v_Cursor, 11, z_Address);
    DBMS_SQL.Column_Value(v_Cursor, 12, z_AddrNo);

    -- �Y���P�Ȥ�, �h���ӫȤ�������
    if v_CurrCustId != z_CustId then
      v_CurrCustId := z_CustId;
      -- ���Ȥ᪺�o�����
      begin
        select InvNo, InvTitle, InvAddress into z_InvNo, z_InvTitle, z_InvAddress
		from SO002 where CustId=v_CurrCustId;
	if z_InvTitle is null then
	  z_InvTitle := '�Ƚs'||z_CustId;
	end if;
	if z_InvAddress is null then
	  z_InvAddress := z_Address;
	end if;
      exception
	when others then
	  z_InvNo := null;
	  z_InvTitle := null;
	  z_InvAddress := null;
      end;

      -- ���H��a�}���l���ϸ�
      begin
        select ZipCode into z_ZipCode from SO014 where AddrNo=z_AddrNo;
      exception
	when others then
	  z_ZipCode := null;
      end;
    end if;

    -- �M�w�ҵ|�O
    if v_CurrCitemCode != z_CitemCode then
      v_CurrCitemCode := z_CitemCode;
      select TaxCode into v_TaxCode from CD019 where CodeNo=v_CurrCitemCode;
      if v_TaxCode = 1 then
	z_TaxType := '1';
      elsif v_TaxCode = 3 then
	z_TaxType := '2';
      else
	z_TaxType := '3';
      end if;
    end if;

    -- �����ʤ���P�ɶ�
    z_UpdA := to_date(to_char(to_number(substr(z_UpdTime,1,2))+1911)||substr(z_UpdTime,3,6),'YYYY/MM/DD');
    z_UpdB := to_number(substr(z_UpdTime, 10, 2)||substr(z_UpdTime, 13, 2));

    -- �s�W�ܵo�����ɼȦs��
    insert into Tmp003 
	(Cust_Id, Invoice_No, Title, Address, Cod_Invoic, 
	Bill_No, Date_c, C_Real, CGName, Acct_Id, 
	CGCode, Ip_Dat, Ip_Nam, Tax_Type, Tel_No, 
	Ip_Tim, Zip_No, Charg_addr, Date_Start, Date_Stop, 
	Notes) 
	values 
	(to_char(v_CurrCustId), to_char(z_InvNo), z_InvTitle, z_InvAddress, '1', 
	z_BillNo, z_RealDate, z_RealAmt, z_CitemName, null, 
	to_char(z_CitemCode), z_UpdA, null, z_TaxType, null, 
	z_UpdB, z_ZipCode, z_Address, z_RealStartDate, z_RealStopDate, 
	substr(z_Note,1,40));

    v_RcdCount := v_RcdCount + 1;
  end loop;

  DBMS_SQL.Close_cursor(v_Cursor);	-- Close cursor

  -- �R���Ȧs��
  --delete from Tmp003;

  --raise e_Final;
  p_RetMsg := '���ͧ���, �@��O'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||
		'��, �B�z����='||to_char(v_RcdCount);

  commit;
  return 0;


EXCEPTION
/*
  when e_Final then
    p_RetMsg := '���ͧ���, �@��O'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||
		'��, ���͵���='||to_char(v_RcdCount);
    -- �s�W�@����ƦܥX��@�~log��(SO059), *** ����b�e�ݰ� ***
    -- insert into SO059 (UpdTime, UpdName, Para, Result) values 
	-- (v_UpdTime, p_UpdName, v_SQL1, p_RetMsg);
--    commit;		-- ��VB��Commit ?
    return 0;
*/
  WHEN OTHERS THEN 
    rollback;
    DBMS_SQL.Close_cursor(v_Cursor);
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    p_RetMsg := '��L���~';
    return -99;
END;
/