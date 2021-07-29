 
CREATE OR REPLACE FUNCTION SF_MduToSingle
  (p_MduId in VARCHAR2, p_ClassSQL in VARCHAR2, p_StatusSQL in VARCHAR2, p_BCCodeNo in NUMBER,
   p_BCDescription in VARCHAR2, p_Period in NUMBER, p_Amount in NUMBER, p_ClctDate in VARCHAR2,
   p_UserId in VARCHAR2, p_RetMsg out VARCHAR2) RETURN number
  AS
  v_Cursor	NUMBER;			-- ��dynamic SQL
  v_RetCode	NUMBER;
  v_Command	VARCHAR2(2000);
  v_SubQuery	varchar2(2000) := '';
 
  v_CustId	number;
  v_Count	number := 0;
  v_UpdTime	varchar2(19);
  v_UserName	varchar2(20);
  v_Now 	date;
 
begin
 
  -- �ˬd�Ѽ�
  if (p_MduId is null) or (p_BCCodeNo is null) or (p_BCDescription is null) or 
   (p_Period is null) or (p_Amount is null) or (p_ClctDate is null) or 
   (p_UserId is null) then
    p_RetMsg := 'Wrong parameter.';
    return -1;
  end if;
 
  -- �ھڶǨӰѼơA�զ�����l��һݪ�SQL���O�G<SubQuery>
  v_SubQuery := 'select CustId from SO001 where MduId='||chr(39)||p_MduId||chr(39)||
		' and ChargeType=3';
  if p_ClassSQL is not null then
    v_SubQuery := v_SubQuery || ' and ClassCode1 ' || p_ClassSQL;
  end if;
  if p_StatusSQL is not null then
    v_SubQuery := v_SubQuery || ' and CustStatusCode ' || p_StatusSQL;
  end if;
 
  -- �ھ�<p_UserId>��<���ʤH���W��>, �P<���ʮɶ�>
  select sysdate into v_Now from dual;
  v_UpdTime := GiPackage.GetDTString2(v_Now);
  begin
    select UserName into v_UserName from SO026 where UserId=p_UserId; 
  exception
    when no_data_found then
      p_RetMsg := 'No such user id.'||chr(39)||p_UserId||chr(39);
      return -2;
  end;
 
  -- �R���U�l�᪺�򥻥x�g�����ظ��
  v_Command := 'delete from SO003 where CitemCode='||to_char(p_BCCodeNo)||
	' and CustId in ('||v_SubQuery||')';
  begin
    v_Cursor := DBMS_SQL.Open_cursor;
    DBMS_SQL.Parse(v_Cursor, v_Command, DBMS_SQL.V7);
    v_RetCode := DBMS_SQL.Execute(v_Cursor);
    DBMS_SQL.Close_cursor(v_Cursor); 
  exception
    when others then
      DBMS_SQL.Close_cursor(v_Cursor);
      p_RetMsg := 'SQL���O���~: ' || v_Command;
      return -3;
  end;
 
  -- Loop�Ӥj�ӨC�@�l��:
  --   ����s�W�@���򥻥x�g�����ظ��
  --   ��s���Ȥ����ɤ������O�ݩʦ�"�j�ӭӦ�"�A�ΰ򥻥x�����馨<p_ClctDate>
  v_Cursor := DBMS_SQL.Open_cursor;
  DBMS_SQL.Parse(v_Cursor, v_SubQuery, DBMS_SQL.V7);
  DBMS_SQL.Define_Column(v_Cursor, 1, v_CustId);	-- �w�q���X�����
  v_RetCode := DBMS_SQL.Execute(v_Cursor);
 
  loop
    if DBMS_SQL.Fetch_rows(v_Cursor) > 0 then		-- Fetch 1 row
      DBMS_SQL.Column_Value(v_Cursor, 1, v_CustId);	-- ��CustId����
 
      insert into SO003 (CustId, CitemCode, CitemName, Period, Amount, ClctDate,
	UpdTime, UpdEn) values (v_CustId, p_BCCodeNo, p_BCDescription, p_Period,
	p_Amount, to_date(p_ClctDate, 'YYYYMMDD'), v_UpdTime, v_UserName);
 
      update SO001 set ChargeType=2, NextBCDate=To_Date(p_ClctDate,'YYYYMMDD'),
	UpdTime=v_UpdTime, UpdEn=v_UserName where CustId=v_CustId;
      v_Count := v_Count + 1;
 
    else					-- No more rows to fetch
      exit;
    end if;
  end loop;
 
  DBMS_SQL.Close_cursor(v_Cursor);
 
  -- �վ�Ӥj�Ӱ򥻸�ƪ����e
  update SO017 set MainCustId=null, MainCustName=null, ClctMethod=2, 
    BCAmount=null, BCPeriod=null, UpdTime=v_UpdTime, UpdEn=v_UserName where MduId=p_MduId;
 
  commit;
  p_RetMsg := 'Execution OK, Cust. count='||to_char(v_Count);
  return 0;
 
exception
  when others then
    DBMS_SQL.Close_cursor(v_Cursor);
    rollback;
    p_RetMsg := 'Other error.';
    return -99;
end;
/
