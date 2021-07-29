 
create or replace function SF_GetAmount2
  (p_option number, p_CustId number, p_CitemCode number, p_Period number, 
   p_RealStartDate varchar2, p_RealStopDate varchar2, p_ShouldAmt out number, 
   p_RetMsg out varchar2) 
  return number as
 
  v_Tmp number;
  v_TmpDate date;
  v_PFlag number;
 
  v_Para3 number;
  v_Para10 number;
  v_Para15 number;
  v_StartDate date;
  v_StopDate date;
  v_ClassCode1 number;
  v_MduId varchar2(10);
  v_ChargeType number;
 
  v_Amount number;
  v_DayAmt number;
  v_TruncAmt number;
  v_PFlag1 number;
  v_PFlag2 number;
  v_FindFlag number := 0;
 
begin
 
  -- Check parameters
  if p_CustId is null or p_CitemCode is null or p_Period is null or 
    p_RealStartDate is null or p_RealStartDate is null then
    p_RetMsg := 'Wrong parameter.';
    return -1;
  end if;
 
  p_RetMsg := '';
  p_ShouldAmt := 0;
 
  -- 取收費參數檔相關參數: Para3: 次期起使日同本期截止日, Para10: 使用費率表, Para15: 破月依據
  begin
    select Para3, Para10, Para15 into v_Para3, v_Para10, v_Para15 from SO043;
    if v_Para3=1 then
      v_Para3 := 0;
    else
      v_Para3 := 1;
    end if;
  exception
    when others then
      p_RetMsg := 'No data or too many data in SO043.';
      return -2;
  end;
 
  -- 從週期性項目設定檔找相關資料, 以決定傳回之初值
  begin
    select Amount into v_Amount from SO003 
      where CustId=p_CustId and CitemCode=p_CitemCode;
  exception
    when others then
      v_FindFlag := 0;
  end;
  v_StartDate := to_date(p_RealStartDate, 'YYYYMMDD');
  v_StopDate  := to_date(p_RealStopDate, 'YYYYMMDD');
  v_Tmp := v_StopDate - v_StartDate + v_Para3;
  p_ShouldAmt := ROUND(nvl(v_Amount,0)*v_Tmp/30);
 
DBMS_OUTPUT.PUT_LINE('(1) '||to_char(ROUND(nvl(v_Amount,0)*v_Tmp/30)));
 
  -- 取該客戶相關資料: 客戶類別一, 大樓編號, 收費屬性
  begin
    select ClassCode1, MduId, ChargeType 
      into v_ClassCode1, v_MduId, v_ChargeType
      from SO001 where CustId=p_CustId;
  end;
 
  if v_MduId is not null and v_ChargeType>1 then -- 找大樓費率表, 1=一般
    begin 
      select Amount, DayAmt, TruncAmt, PFlag1, PFlag2 
        into v_Amount, v_DayAmt, v_TruncAmt, v_PFlag1, v_PFlag2
        from CD019SO017 where MduId=v_MduId and CitemCode=p_CitemCode 
	and Period=p_Period;
      v_FindFlag := 1;		-- 有該大樓費率表
    exception 			-- 無該大樓費率表, 故找非特定大樓費率表
      when others then
	begin
	  select Amount, DayAmt, TruncAmt, PFlag1, PFlag2 
	    into v_Amount, v_DayAmt, v_TruncAmt, v_PFlag1, v_PFlag2
	    from CD019SO017 where MduId is null and CitemCode=p_CitemCode 
	    and Period=p_Period;
	  v_FindFlag := 1;
        exception		-- 也無非特定大樓費率表, 將再查客戶類別費率表
	  when others then
	    v_FindFlag := 0;
	end;
    end;
  end if;
 
  if v_FindFlag != 1 then	-- 找客戶類別費率表
    begin 
      select Amount, DayAmt, TruncAmt, PFlag1, PFlag2 
        into v_Amount, v_DayAmt, v_TruncAmt, v_PFlag1, v_PFlag2
        from CD019CD004 where CitemCode=p_CitemCode and Period=p_Period 
	and ClassCode=v_ClassCode1;
      v_FindFlag := 1;		-- 有該客戶類別費率表
    exception 			-- 無該客戶類別費率表, 故找非特定客戶類別費率表
      when others then
	begin
	  select Amount, DayAmt, TruncAmt, PFlag1, PFlag2 
	    into v_Amount, v_DayAmt, v_TruncAmt, v_PFlag1, v_PFlag2
	    from CD019CD004 where ClassCode is null and CitemCode=p_CitemCode 
	    and Period=p_Period;
	  v_FindFlag := 1;
        exception	-- 也無非特定客戶類別費率表, 則沿用先前之計算值
	  when others then
	    v_FindFlag := 0;
	end;
    end;
  end if;
 
  if v_FindFlag = 1 then
    v_Amount := nvl(v_DayAmt,0) * v_Tmp;
    p_ShouldAmt := v_Amount;
    if nvl(v_TruncAmt,0) > 0 then	-- 金額歸整
      p_ShouldAmt := v_Amount - mod(v_Amount, v_TruncAmt);
    end if;
  end if;
 
  return 0;
 
exception
  when others then
    --DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    p_RetMsg := 'Other error.';
    return -99;
end;
/
