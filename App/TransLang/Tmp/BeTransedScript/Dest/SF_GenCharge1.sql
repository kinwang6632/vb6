 
CREATE OR REPLACE function SF_GenCharge1(p_YM varchar2, p_Day1 number, 
  p_Day2 number, p_CompSQL varchar2, p_CitemSQL varchar2, p_ClassSQL varchar2, 
  p_AreaSQL varchar2, p_ServSQL varchar2, p_ClctAreaSQL varchar2, 
  p_MduSQL varchar2, p_StrtSQL varchar2, p_AddrType number, p_GenOverdue number, 
  p_GenPRCust number, p_ToEndDate number, p_CustIdList varchar2, 
  p_UpdEn varchar2, p_UpdName varchar2, p_RetMsg out varchar2) return number
  AS
  v_Para3 	number;
  v_Para5 	number;
  v_Para10	number;
  v_Para15	number;
  v_xUCCode1 	number;
  v_xUCCode2 	number;
  v_xUCName1 	varchar2(12);
  v_xUCName2 	varchar2(12);
  c39		char(1) := chr(39);
  cr		char(1) := chr(13);
 
  v_StartExecTime 	date;
  v_SQL1 		varchar2(1000);
  v_StartDate 		varchar2(10);
  v_StopDate  		varchar2(10);
  v_RcdCount 		number := 0;
  v_TmpTableName	varchar2(20);
  v_CustIdList		varchar2(200);
  v_Index 		number;
  v_CustStatusCode	number;
  v_CurrCustId 		number;
  v_Item 		number;
  v_DupCount		number;
  v_DupCount1		number;
  v_DupCount2		number;
 
  z_CustId 	number;
  z_CitemCode 	number;
  z_CitemName 	varchar2(12);
  z_Period 	number;
  z_Amount 	number;
  z_StartDate 	date;
  z_StopDate  	date;
  z_ClctDate  	date;
  z_LastDate	date;
 
  x_CMCode 	number;
  x_CMName 	varchar2(12);
  x_CompCode 	number;
  x_AddrNo 	number;
  --x_Address 	varchar2(60);
  x_ChargeType 	number;
  x_MduId 	varchar2(8);
  x_ServCode 	varchar2(3);
  x_ClctAreaCode varchar2(3);
  x_ClassCode1 	number;
  x_ClctEn 	varchar2(10);
  x_ClctName 	varchar2(20);
  x_StrtCode 	number;
  x_CustCount 	number;
  x_AreaCode	varchar2(3);
  x_CDueDate	date;
 
  x_BillNo 	varchar2(15);
  x_UCCode 	number;
  x_UCName 	varchar2(12);
 
  x_ShouldDate	date;
  x_StartDate	date;
  x_StopDate	date;
 
  v_Now 	date;
  v_UpdTime 	varchar2(19);
 
  v_Cursor	number;			-- 供dynamic SQL
  v_RetCode	number;
  v_Command	varchar2(2000);
  v_D1		date;
  v_D2		date;
  i 		number;
 
  e_Final exception;
 
BEGIN
 
  -- 參數檢查
  if p_YM is null or nvl(p_Day1,0)<=0 or nvl(p_Day2,0)<=0 or 
    p_Day1>31 or p_Day2>31 or (nvl(p_AddrType,0) not in (1,2)) or
    p_GenOverdue is null or (p_GenOverdue not in (0,1)) or
    p_GenPRCust is null or (p_GenPRCust not in (0,1)) or
    p_ToEndDate is null or (p_ToEndDate not in (0,1)) then
    p_RetMsg := 'Wrong parameter.';
    return -1;
  end if;
 
 
  /* 取收費參數檔(SO043)中的次期起始日同本期截止日(Para3),
     延後幾日為次收日(Para5), 使用費率表(Para10), 破月依據(Para15),
     並調整Para3 */
  begin
    select Para3, Para5, Para10, Para15 into v_Para3, v_Para5, v_Para10, v_Para15 from SO043;
    if v_Para3 = 1 then
      v_Para3 := 0;
    else
      v_Para3 := 1;
    end if;
  exception 
    when others then
      p_RetMsg := 'No data or too many data in SO043.';
      return -2;
  end;
 
  /* 取未繳費原因代碼檔中(CD013)參考號為1(待收), 為2(轉帳不成功)
     之代碼/名稱, 假設為<待收之代碼>, <待收之名稱>, <轉帳不成功之代碼>,
     <轉帳不成功之名稱>, 做為後續使用. */
  begin
    select CodeNo, Description into v_xUCCode1, v_xUCName1 from CD013 where RefNo=1;
    select CodeNo, Description into v_xUCCode2, v_xUCName2 from CD013 where RefNo=2;
  exception 
    when no_data_found then
      p_RetMsg := '未繳費原因代碼檔中未設參考號為1或2的資料';
      return -3;
    when too_many_rows then
      p_RetMsg := '未繳費原因代碼檔中設參考號為1或2的資料太多';
      return -4;
  end;
 
  -- 刪除各暫存檔
  delete from SO065;
  delete from SO066;
  delete from SO067;
  delete from SO068;
 
  -- 設定各變數初值
  v_StartExecTime := sysdate;		-- 開始執行時間
  v_SQL1 := null;
  v_StartDate := p_YM||ltrim(to_char(p_Day1,'09'));
  v_StopDate  := p_YM||ltrim(to_char(p_Day2,'09'));
  v_D1 := to_date(v_StartDate, 'YYYYMMDD');
  v_D2 := to_date(v_StopDate,  'YYYYMMDD');
 
/* 本段是多餘的 !!!
  -- 若產生欠費客戶於本期收費資料
  if p_GenOverdue = 1 then
    v_SQL1 := v_SQL1||' NextBCDate<=to_date('||c39||v_StopDate||c39||','||c39||
	'YYYYMMDD'||c39||')';
  else
    v_SQL1 := v_SQL1||' NextBCDate>=to_date('||c39||v_StartDate||c39||','||c39||
	'YYYYMMDD'||c39||') and NextBCDate<=to_date('||c39||v_StopDate||c39||','||c39||
	'YYYYMMDD'||c39||')';
  end if;
*/
  -- 若有公司別條件
  if p_CompSQL is not null then
    v_SQL1 := v_SQL1 || ' and CompCode ' || p_CompSQL;
  end if;
 
  -- 若有客戶類別條件
  if p_ClassSQL is not null then
    v_SQL1 := v_SQL1 || ' and ClassCode1 ' || p_ClassSQL;
  end if;
 
  -- 若有服務區條件
  if p_AreaSQL is not null then
    v_SQL1 := v_SQL1 || ' and ServCode ' || p_AreaSQL;
  end if;
 
  -- 若有收費區條件
  if p_ClctAreaSQL is not null then
    v_SQL1 := v_SQL1 || ' and ClctAreaCode ' || p_ClctAreaSQL;
  end if;
 
  -- 若有大樓條件
  if p_MduSQL is not null then
    v_SQL1 := v_SQL1 || ' and MduId ' || p_MduSQL;
  end if;
 
  -- 若要產生待拆客戶於本期待收資料, 則
  if p_GenPRCust = 1 then
    v_SQL1 := v_SQL1 || ' and (CustStatusCode=1 or CustStatusCode=6)';
  else
    v_SQL1 := v_SQL1 || ' and CustStatusCode=1';
  end if;
 
  -- 進行第一階段查詢, 並查詢將結果存入應收資料產生暫存檔A(SO065)
  v_SQL1 := 'insert into SO065 (select CustId, InstAddrNo from SO001 where ' || ltrim(v_SQL1,' and') || ')';
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);
 
  begin
    v_Cursor := DBMS_SQL.Open_cursor;
    DBMS_SQL.Parse(v_Cursor, v_SQL1, DBMS_SQL.V7);
    v_RetCode := DBMS_SQL.Execute(v_Cursor);
    DBMS_SQL.Close_cursor(v_Cursor);
  exception
    when others then
      DBMS_SQL.Close_cursor(v_Cursor);
      rollback;
      p_RetMsg := 'SQL1錯誤: ' || v_SQL1;
      return -5;
  end;
 
  v_TmpTableName := 'SO065';
  select count(*) into v_RcdCount from SO065;
  if v_RcdCount = 0 then
    raise e_Final;
  end if;
 
 
  -- 若行政區條件或街道條件有值, 則進行第二階段查詢
  if p_AreaSQL is not null or p_StrtSQL is not null then
    v_SQL1 := 'select A.* from SO065 A, SO014 B where A.InstAddrNo=B.AddrNo';
    -- 若有行政區條件
    if p_AreaSQL is not null then
      v_SQL1 := v_SQL1 || ' and B.AreaCode ' || p_AreaSQL;
    end if;
    -- 若有街道條件
    if p_StrtSQL is not null then
      v_SQL1 := v_SQL1 || ' and B.StrtCode ' || p_StrtSQL;
    end if;
    -- 進行第二階段查詢, 並將結果存入應收資料產生暫存檔B(SO066)
    v_SQL1 := 'insert into SO066 (' || v_SQL1 || ')';
    -- DBMS_OUTPUT.PUT_LINE(v_SQL1);
 
    begin
      v_Cursor := DBMS_SQL.Open_cursor;
      DBMS_SQL.Parse(v_Cursor, v_SQL1, DBMS_SQL.V7);
      v_RetCode := DBMS_SQL.Execute(v_Cursor);
      DBMS_SQL.Close_cursor(v_Cursor);
    exception
      when others then
        DBMS_SQL.Close_cursor(v_Cursor);
        rollback;
        p_RetMsg := 'SQL2錯誤: ' || v_SQL1;
        return -5;
    end;
 
    v_TmpTableName := 'SO066';
    select count(*) into v_RcdCount from SO066;
    if v_RcdCount = 0 then
      raise e_Final;
    end if;
  end if;
 
 
  /*
    若客戶編號字串有值, 則以本項為主, 處理方式如下:
    loop將字串中以','隔開的客戶編號取出
	根據此客戶編號至客戶基本資料檔取客戶狀態代碼
	若該客戶狀態=1 or 產生待拆客戶於本期收費資料且該客戶狀態=6,
	    則新增一筆資料至應收資料產生暫存檔B(SO066)
    <暫存檔名> = 'SO066'
  */
 
  if p_CustIdList is not null then
    v_RcdCount := 0;
    v_CustIdList := p_CustIdList;
    v_Index := 1;
    while v_CustIdList is not null loop
      v_Index := instr(v_CustIdList, ',');
      z_CustId := 0;
      if v_Index > 0 then
	begin
	  z_CustId := to_number(rtrim(ltrim(substr(v_CustIdList, 1, v_Index-1))));
	  v_CustIdList := substr(v_CustIdList, v_Index+1);
	exception
	  when others then
	    v_CustIdList := null;
	end;
      else
	begin
	  z_CustId := to_number(v_CustIdList);
	  v_CustIdList := null;
	exception
	  when others then
	    v_CustIdList := null;
	end;
      end if; 
      if nvl(z_CustId,0) > 0 then
	--DBMS_OUTPUT.PUT_LINE('I='||to_char(v_Index)||', CId='||to_char(z_CustId));
	begin
	  select CustStatusCode into v_CustStatusCode from SO001 where CustId=z_CustId;
	  if v_CustStatusCode=1 or p_GenPRCust=1 and v_CustStatusCode=6 then
	    insert into SO066 (CustId) values (z_CustId);
	    v_RcdCount := v_RcdCount + 1;
	  end if;
	exception
	  when others then
	    v_RcdCount := v_RcdCount;
	end;
      end if;
    end loop;
    v_TmpTableName := 'SO066';
    if v_RcdCount = 0 then
      raise e_Final;
    end if;
  end if;
 
 
  -- 進行第三階段查詢, 將產生應收資料, 存放於SO067
  v_SQL1 := 'select A.CustId, B.CitemCode, B.CitemName, B.Period, B.Amount, B.StartDate, B.StopDate, B.ClctDate, B.LastDate from '
		|| v_TmpTableName || ' A, SO003 B where A.CustId=B.CustId';
 
  -- 若收費項目條件有值
  if p_CitemSQL is not null then
    v_SQL1 := v_SQL1 || ' and B.CitemCode ' || p_CitemSQL;
  end if;
 
  -- 若產生欠費客戶於本期收費資料
  if p_GenOverdue = 1 then
    v_SQL1 := v_SQL1||' and B.Amount!=0 and B.ClctDate<='||'to_date('||c39||v_StopDate
		||c39||','||c39||'YYYYMMDD'||c39||') order by A.CustId';
  else -- 只產生未欠費客戶於本期的
    v_SQL1 := v_SQL1||' and B.Amount!=0 and B.ClctDate>='||'to_date('||c39||v_StartDate
		||c39||','||c39||'YYYYMMDD'||c39||')'||
		' and B.ClctDate<='||'to_date('||c39||v_StopDate
		||c39||','||c39||'YYYYMMDD'||c39||') order by A.CustId';
  end if;
  --DBMS_OUTPUT.PUT_LINE(v_SQL1);
 
  v_CurrCustId := 0;
  v_Item := 0;
  v_RcdCount := 0;
 
  -- 執行該查詢
  begin
    v_Cursor := DBMS_SQL.Open_cursor;
    DBMS_SQL.Parse(v_Cursor, v_SQL1, DBMS_SQL.V7);
    -- 定義取出的欄位
    DBMS_SQL.Define_Column(v_Cursor, 1, z_CustId);
    DBMS_SQL.Define_Column(v_Cursor, 2, z_CitemCode);
    DBMS_SQL.Define_Column(v_Cursor, 3, z_CitemName, 12);
    DBMS_SQL.Define_Column(v_Cursor, 4, z_Period);
    DBMS_SQL.Define_Column(v_Cursor, 5, z_Amount);
    DBMS_SQL.Define_Column(v_Cursor, 6, z_StartDate);
    DBMS_SQL.Define_Column(v_Cursor, 7, z_StopDate);
    DBMS_SQL.Define_Column(v_Cursor, 8, z_ClctDate);
    --**DBMS_SQL.Define_Column(v_Cursor, 9, z_LastDate);
    v_RetCode := DBMS_SQL.Execute(v_Cursor);
  exception
    when others then
      DBMS_SQL.Close_cursor(v_Cursor);
      rollback;
      p_RetMsg := 'SQL3錯誤: ' || v_SQL1;
      return -5;
  end;
 
  -- 設定所有應收資料皆一樣的欄位
  v_Now := sysdate;				-- 產生時間
  v_UpdTime := GiPackage.GetDTString2(v_Now);	-- 異動時間
  --delete from t1;
 
  loop					-- 取到的每一筆資料
    if DBMS_SQL.Fetch_rows(v_Cursor) <= 0 then		-- Fetch 1 rcd
      exit;				-- No more rows to fetch
    end if;
    -- 取各欄位值
    DBMS_SQL.Column_Value(v_Cursor, 1, z_CustId);
    DBMS_SQL.Column_Value(v_Cursor, 2, z_CitemCode);
    DBMS_SQL.Column_Value(v_Cursor, 3, z_CitemName);
    DBMS_SQL.Column_Value(v_Cursor, 4, z_Period);
    DBMS_SQL.Column_Value(v_Cursor, 5, z_Amount);
    DBMS_SQL.Column_Value(v_Cursor, 6, z_StartDate);
    DBMS_SQL.Column_Value(v_Cursor, 7, z_StopDate);
    DBMS_SQL.Column_Value(v_Cursor, 8, z_ClctDate);
    --**BMS_SQL.Column_Value(v_Cursor, 9, z_LastDate);
    --DBMS_OUTPUT.PUT_LINE(z_CustId);
    --insert into t1 values (z_CustId, z_CitemCode, z_CitemName, z_Period, z_Amount, z_StartDate, z_StopDate, z_ClctDate);
 
    -- 若<目前客編>!=該筆客戶編號, 則: 至客戶基本資料檔取相關欄位
    if v_CurrCustId != z_CustId then
      v_CurrCustId := z_CustId;
      v_Item := 0;
 
      if p_AddrType = 1 then		-- 若地址依據為1(收費地址)
	select CMCode, CMName, CompCode, ChargeAddrNo, ChargeType, 
	  MduId, ServCode, ClctAreaCode, ClassCode1 into x_CMCode, x_CMName, 
	  x_CompCode, x_AddrNo, x_ChargeType, x_MduId, 
	  x_ServCode, x_ClctAreaCode, x_ClassCode1
	  from SO001 where CustId=v_CurrCustId;
      else
	select CMCode, CMName, CompCode, InstAddrNo, ChargeType, 
	  MduId, ServCode, ClctAreaCode, ClassCode1 into x_CMCode, x_CMName, 
	  x_CompCode, x_AddrNo, x_ChargeType, x_MduId, 
	  x_ServCode, x_ClctAreaCode, x_ClassCode1
	  from SO001 where CustId=v_CurrCustId;
      end if;
 
      -- 再根據取出之地址編號至地址資料檔取相關欄位: 收費人員代號,收費人員名稱,街道代碼
      select ClctEn, ClctName, StrtCode, AreaCode 
	into x_ClctEn, x_ClctName, x_StrtCode, x_AreaCode
	from SO014 where AddrNo=x_AddrNo;  -- 註: 新版之地址編號為唯一值
 
      -- 若為統收戶(大樓子戶), 則: 根據MduId至大樓資料檔(SO017)取有效戶數
      -- 否則: 有效戶數=1
      if x_ChargeType=3 and x_MduId is not null then
	-- x_CustCount 至少為1
	select nvl(InstCnt,1) into x_CustCount from SO017 where MduId=x_MduId;
      else
	x_CustCount := 1;
      end if;
 
      -- 單據編號(BillNo)='YYYYMM'+'B' + <8碼的客戶編號, 不足補0>
      x_BillNo := p_YM || 'B' || ltrim(to_char(v_CurrCustId,'09999999'));
 
      -- 若收費方式=4(轉帳), 則: 未收原因=<轉帳不成功之代碼/名稱>
      -- 否則: 未收原因=<待收之代碼/名稱>
      if x_CMCode = 4 then
	x_UCCode := v_xUCCode2;
	x_UCName := v_xUCName2;
      else
	x_UCCode := v_xUCCode1;
	x_UCName := v_xUCName1;
      end if;
    end if;  -- end of: if v_CurrCustId != z_CustId then
 
    -- 決定本筆應收資料的應收日, 有效期限, 金額, 期數, 再新增應收資料
    -- 若不產生欠費客戶於本期收費資料或要產生但該筆次收日介於起迄日範圍內
    if p_GenOverDue=0 or p_GenOverDue=1 and z_ClctDate>=v_D1 and z_ClctDate<=v_D2 then
      x_ShouldDate := z_ClctDate;
      x_StartDate := z_StopDate + v_Para3;
      x_StopDate  := add_months(z_StopDate, z_Period);
 
      -- 若為大樓個收戶, 及大樓個收戶收費至合約到期日, 若截止日超出合約到期日,
      -- 則調整截止日=合約到期日
      if p_ToEndDate=1 and x_ChargeType=2 and x_MduId is not null then	-- 大樓個收戶
	select CDueDate into x_CDueDate from SO017 where MduId=x_MduId;
	if x_CDueDate is not null and x_CDueDate < x_StopDate then
	  x_StopDate  := x_CDueDate;
	  v_RetCode := SF_GetAmount2(2, v_CurrCustId, z_CitemCode, z_Period, 
		to_char(x_StartDate, 'YYYYMMDD'), to_char(x_StopDate, 'YYYYMMDD'), z_Amount, p_RetMsg);
	end if;
      end if;
 
      -- 檢查該筆資料是否已產生於正式應收資料檔SO033 (以下為笨方法)
      v_DupCount:=0;
      v_DupCount1:=0;
      v_DupCount2:=0;
      select count(*) into v_DupCount from SO033 
	where CustId=v_CurrCustId and CitemCode=z_CitemCode and ShouldDate=x_ShouldDate;
      if nvl(v_DupCount,0)=0 then
         select count(*) into v_DupCount1 from SO034 
	   where CustId=v_CurrCustId and CitemCode=z_CitemCode and ShouldDate=x_ShouldDate;
      end if;
      if nvl(v_DupCount,0)=0 and nvl(v_DupCount1,0)=0 then
         select count(*) into v_DupCount2 from SO035 
	   where CustId=v_CurrCustId and CitemCode=z_CitemCode and ShouldDate=x_ShouldDate;
      end if;
 
      /* -- 以下方法假設應收資料檔無空隙資料
      if z_LastDate is null or x_ShouldDate > z_LastDate then
	v_DupCount := 0;
      else
	v_DupCount := 1;
      end if; */
 
      -- 新增一筆資料至暫時應收資料檔C(SO067)
      if v_DupCount=0 then
        v_Item := v_Item + 1;
        if v_DupCount1>0 then
           v_Item := v_DupCount1 + 1;
        end if;
        if v_DupCount2>0 and v_DupCount2>v_DupCount1 then
           v_Item := v_DupCount2 + 1;
        end if;
	begin
	  insert into SO067 
	    (CustId, BillNo, Item, CitemCode, CitemName, 
	    OldAmt, OldPeriod, OldStartDate, OldStopDate, ShouldDate, 
	    ShouldAmt, RealPeriod, RealStartDate, RealStopDate, ClctEn, 
	    ClctName, CMCode, CMName, UCCode, UCName, 
	    CreateTime, CreateEn, CompCode, AddrNo, StrtCode, 
	    MduId, ServCode, ClctAreaCode, OldClctEn, OldClctName, 
	    ClassCode, UpdTime, UpdEn, ClctYM, CustCount, AreaCode, RealAmt)
	    VALUES 
	    (v_CurrCustId, x_BillNo, v_Item, z_CitemCode, z_CitemName, 
	    z_Amount, z_Period, x_StartDate, x_StopDate, x_ShouldDate, 
	    Z_Amount, z_Period, x_StartDate, x_StopDate, x_ClctEn, 
	    x_ClctName, x_CMCode, x_CMName, x_UCCode, x_UCName, 
	    v_Now, p_UpdEn, x_CompCode, x_AddrNo, x_StrtCode, 
	    x_MduId, x_ServCode, x_ClctAreaCode, x_ClctEn, x_ClctName, 
	    x_ClassCode1, v_UpdTime, p_UpdName, to_number(p_YM), x_CustCount, x_AreaCode, 0);
	  --** update SO003 set LastDate=x_ShouldDate where CustId=v_CurrCustId and CitemCode=z_CitemCode;
	  v_RcdCount := v_RcdCount + 1;
	exception
	  when others then -- 若無法新增至暫時應收資料檔C的資料,將log至SO068
	    insert into SO068 values (v_CurrCustId, x_BillNo, v_Item, z_CitemCode, 
	      z_CitemName, x_ShouldDate, z_Period, z_Amount, x_StartDate, x_StopDate);
	end;
      end if;
 
    else	-- 產生欠費客戶之本期應收與過去欠費資料
      i := 0;
      x_ShouldDate := z_ClctDate;
      x_StopDate  := z_StopDate;
 
      while i<=24 and x_ShouldDate <= v_D2 loop
	x_StartDate := x_StopDate + v_Para3;
	x_StopDate  := add_months(x_StopDate, z_Period);
 
        -- 若為大樓個收戶, 及大樓個收戶收費至合約到期日, 若截止日超出合約到期日,
        -- 則調整截止日=合約到期日
        if p_ToEndDate=1 and x_ChargeType=2 and x_MduId is not null then	-- 大樓個收戶
	  select CDueDate into x_CDueDate from SO017 where MduId=x_MduId;
	  if x_CDueDate is not null and x_CDueDate < x_StopDate then
	    i := 24;		-- 強迫跳出迴圈
	    x_StopDate  := x_CDueDate;
	    v_RetCode := SF_GetAmount2(2, v_CurrCustId, z_CitemCode, z_Period, 
		to_char(x_StartDate, 'YYYYMMDD'), to_char(x_StopDate, 'YYYYMMDD'), z_Amount, p_RetMsg);
	  end if;
        end if;
 
	-- 檢查該筆資料是否已產生於正式應收資料檔SO033 (以下為笨方法)
        v_DupCount:=0;
        v_DupCount1:=0;
        v_DupCount2:=0;
	select count(*) into v_DupCount from SO033 
	  where CustId=v_CurrCustId and CitemCode=z_CitemCode and ShouldDate=x_ShouldDate;
        if nvl(v_DupCount,0)=0 then
           select count(*) into v_DupCount1 from SO034 
	     where CustId=v_CurrCustId and CitemCode=z_CitemCode and ShouldDate=x_ShouldDate;
        end if;
        if nvl(v_DupCount,0)=0 and nvl(v_DupCount1,0)=0 then
           select count(*) into v_DupCount2 from SO035 
	     where CustId=v_CurrCustId and CitemCode=z_CitemCode and ShouldDate=x_ShouldDate;
        end if;
 
	/* -- 以下方法假設應收資料檔無空隙資料
	if z_LastDate is null or x_ShouldDate > z_LastDate then
	  v_DupCount := 0;
	else
	  v_DupCount := 1;
	end if; */
 
	-- 新增一筆資料至暫時應收資料檔C(SO067)
	if v_DupCount=0 then
	  v_Item := v_Item + 1;
          if v_DupCount1>0 then
             v_Item := v_DupCount1 + 1;
          end if;
          if v_DupCount2>0 and v_DupCount2>v_DupCount1 then
             v_Item := v_DupCount2 + 1;
          end if;
	  begin
	    insert into SO067 
	      (CustId, BillNo, Item, CitemCode, CitemName, 
	      OldAmt, OldPeriod, OldStartDate, OldStopDate, ShouldDate, 
	      ShouldAmt, RealPeriod, RealStartDate, RealStopDate, ClctEn, 
	      ClctName, CMCode, CMName, UCCode, UCName, 
	      CreateTime, CreateEn, CompCode, AddrNo, StrtCode, 
	      MduId, ServCode, ClctAreaCode, OldClctEn, OldClctName, 
	      ClassCode, UpdTime, UpdEn, ClctYM, CustCount, AreaCode, RealAmt)
	      VALUES 
	      (v_CurrCustId, x_BillNo, v_Item, z_CitemCode, z_CitemName, 
	      z_Amount, z_Period, x_StartDate, x_StopDate, x_ShouldDate, 
	      Z_Amount, z_Period, x_StartDate, x_StopDate, x_ClctEn, 
	      x_ClctName, x_CMCode, x_CMName, x_UCCode, x_UCName, 
	      v_Now, p_UpdEn, x_CompCode, x_AddrNo, x_StrtCode, 
	      x_MduId, x_ServCode, x_ClctAreaCode, x_ClctEn, x_ClctName, 
	      x_ClassCode1, v_UpdTime, p_UpdName, to_number(p_YM), x_CustCount, x_AreaCode, 0);
	    --** update SO003 set LastDate=x_ShouldDate where CustId=v_CurrCustId and CitemCode=z_CitemCode;
	    v_RcdCount := v_RcdCount + 1;
	  exception
	    when others then -- 若無法新增至暫時應收資料檔C的資料將log至SO068
	      insert into SO068 values (v_CurrCustId, x_BillNo, v_Item, z_CitemCode, 
	        z_CitemName, x_ShouldDate, z_Period, z_Amount, x_StartDate, x_StopDate);
	  end;
	end if;
 
	i := i + 1;
	x_ShouldDate := x_StopDate + v_Para5;
      end loop;
 
    end if;
  end loop;
 
  DBMS_SQL.Close_cursor(v_Cursor);	-- Close cursor
 
  -- 將暫時應收資料檔C的內容複製到暫時應收資料檔(SO032)
  delete from SO032;
  begin
    insert into SO032 (select * from SO067);
  exception
    when others then
      rollback;
      p_RetMsg := 'Fail to copy from SO067 to SO032.';
      return -6;
  end;
/*
  -- 刪除各暫存檔
  delete from SO065;
  delete from SO066;
  delete from SO067;
  delete from SO068;
*/
 
  raise e_Final;
  return 0;
 
EXCEPTION
  when e_Final then
    p_RetMsg := 'Execution OK, it takes'||to_char(trunc(86400*(sysdate-v_StartExecTime)))||
		'seconds, Rcd count='||to_char(v_RcdCount);
    -- 新增一筆資料至出單作業log檔(SO059), *** 應放在前端做 ***
    -- insert into SO059 (UpdTime, UpdName, Para, Result) values 
	-- (v_UpdTime, p_UpdName, v_SQL1, p_RetMsg);
    commit;		-- 由VB端Commit ?
    return 0;
 
  WHEN OTHERS THEN 
    rollback;
    DBMS_SQL.Close_cursor(v_Cursor);
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    p_RetMsg := 'Other error.';
    return -99;
END;
/
