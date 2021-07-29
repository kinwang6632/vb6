 
CREATE OR REPLACE FUNCTION SF_MduToSingle
  (p_MduId in VARCHAR2, p_ClassSQL in VARCHAR2, p_StatusSQL in VARCHAR2, p_BCCodeNo in NUMBER,
   p_BCDescription in VARCHAR2, p_Period in NUMBER, p_Amount in NUMBER, p_ClctDate in VARCHAR2,
   p_UserId in VARCHAR2, p_RetMsg out VARCHAR2) RETURN number
  AS
  v_Cursor	NUMBER;			-- 供dynamic SQL
  v_RetCode	NUMBER;
  v_Command	VARCHAR2(2000);
  v_SubQuery	varchar2(2000) := '';
 
  v_CustId	number;
  v_Count	number := 0;
  v_UpdTime	varchar2(19);
  v_UserName	varchar2(20);
  v_Now 	date;
 
begin
 
  -- 檢查參數
  if (p_MduId is null) or (p_BCCodeNo is null) or (p_BCDescription is null) or 
   (p_Period is null) or (p_Amount is null) or (p_ClctDate is null) or 
   (p_UserId is null) then
    p_RetMsg := 'Wrong parameter.';
    return -1;
  end if;
 
  -- 根據傳來參數，組成選取子戶所需的SQL指令：<SubQuery>
  v_SubQuery := 'select CustId from SO001 where MduId='||chr(39)||p_MduId||chr(39)||
		' and ChargeType=3';
  if p_ClassSQL is not null then
    v_SubQuery := v_SubQuery || ' and ClassCode1 ' || p_ClassSQL;
  end if;
  if p_StatusSQL is not null then
    v_SubQuery := v_SubQuery || ' and CustStatusCode ' || p_StatusSQL;
  end if;
 
  -- 根據<p_UserId>取<異動人員名稱>, 與<異動時間>
  select sysdate into v_Now from dual;
  v_UpdTime := GiPackage.GetDTString2(v_Now);
  begin
    select UserName into v_UserName from SO026 where UserId=p_UserId; 
  exception
    when no_data_found then
      p_RetMsg := 'No such user id.'||chr(39)||p_UserId||chr(39);
      return -2;
  end;
 
  -- 刪除各子戶的基本台週期項目資料
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
      p_RetMsg := 'SQL指令錯誤: ' || v_Command;
      return -3;
  end;
 
  -- Loop該大樓每一子戶:
  --   為其新增一筆基本台週期項目資料
  --   更新戶於客戶資料檔中的收費屬性成"大樓個收"，及基本台次收日成<p_ClctDate>
  v_Cursor := DBMS_SQL.Open_cursor;
  DBMS_SQL.Parse(v_Cursor, v_SubQuery, DBMS_SQL.V7);
  DBMS_SQL.Define_Column(v_Cursor, 1, v_CustId);	-- 定義取出的欄位
  v_RetCode := DBMS_SQL.Execute(v_Cursor);
 
  loop
    if DBMS_SQL.Fetch_rows(v_Cursor) > 0 then		-- Fetch 1 row
      DBMS_SQL.Column_Value(v_Cursor, 1, v_CustId);	-- 取CustId欄位值
 
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
 
  -- 調整該大樓基本資料的內容
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
