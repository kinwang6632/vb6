
create or replace function SF_GetAmount3
  (p_CustId number, p_CitemCode number, p_Period number) 
  return number as

/*  v_Para3 number;
  v_Para10 number; */

  v_ClassCode1 number;
  v_MduId varchar2(10);
  v_ChargeType number;

  v_Amount number;
  v_Amount2 number;
  v_FindFlag number := 0;

begin

  -- Check parameters
  if p_CustId is null or p_CitemCode is null or p_Period is null then
    -- p_RetMsg := '參數錯誤';
    return -1;
  end if;

  --p_RetMsg := '';
/*
  -- 取收費參數檔相關參數: Para3: 次期起使日同本期截止日, Para10: 使用費率表
  begin
    select Para3, Para10 into v_Para3, v_Para10 from SO043;
    if v_Para3=1 then
      v_Para3 := 0;
    else
      v_Para3 := 1;
    end if;
  exception
    when others then
      --p_RetMsg := '收費參數檔內容有誤';
      return -2;
  end;
*/

  -- 從週期性項目設定檔找相關資料, 以決定傳回之初值
  begin
    select nvl(Amount,0) into v_Amount2 from SO003 
      where CustId=p_CustId and CitemCode=p_CitemCode;
  exception
    when others then
      v_FindFlag := 0;
  end;

  -- 取該客戶相關資料: 客戶類別一, 大樓編號, 收費屬性
  begin
    select ClassCode1, MduId, ChargeType 
      into v_ClassCode1, v_MduId, v_ChargeType
      from SO001 where CustId=p_CustId;
  end;

  if v_MduId is not null and v_ChargeType>1 then -- 找大樓費率表, 1=一般
    begin 
      select nvl(Amount,0) into v_Amount
        from CD019SO017 where MduId=v_MduId and CitemCode=p_CitemCode 
	and Period=p_Period;
      v_FindFlag := 1;		-- 有該大樓費率表
    exception 			-- 無該大樓費率表, 故找非特定大樓費率表
      when others then
	begin
	  select nvl(Amount,0) into v_Amount
	    from CD019SO017 where MduId is null and CitemCode=p_CitemCode and Period=p_Period;
	  v_FindFlag := 1;
        exception		-- 也無非特定大樓費率表, 將再查客戶類別費率表
	  when others then
	    v_FindFlag := 0;
	end;
    end;
  end if;

  if v_FindFlag != 1 then	-- 找客戶類別費率表
    begin 
      select nvl(Amount,0) into v_Amount
        from CD019CD004 where CitemCode=p_CitemCode and Period=p_Period 
	and ClassCode=v_ClassCode1;
      v_FindFlag := 1;		-- 有該客戶類別費率表
    exception 			-- 無該客戶類別費率表, 故找非特定客戶類別費率表
      when others then
	begin
	  select nvl(Amount,0) into v_Amount
	    from CD019CD004 where ClassCode is null and CitemCode=p_CitemCode and Period=p_Period;
	  v_FindFlag := 1;
        exception	-- 也無非特定客戶類別費率表, 則沿用原收費項目代碼中預設值
	  when others then
	    v_FindFlag := 0;
	end;
    end;
  end if;

  if v_FindFlag = 1 then
    return v_Amount;
  else
    return v_Amount2;
  end if;

exception
  when others then
     DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    --p_RetMsg := '其他錯誤';
    return -99;
end;
/