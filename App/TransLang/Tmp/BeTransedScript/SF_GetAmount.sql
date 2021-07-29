
create or replace function SF_GetAmount
  (p_option number, p_CustId number, p_CitemCode number, p_Period number, 
   p_RealStartDate varchar2, p_RealStopDate out varchar2, p_ShouldAmt out number, 
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

  v_Period number;
  v_Amount number;
  v_DayAmt number;
  v_TruncAmt number;
  v_PFlag1 number;
  v_PFlag2 number;
  v_FindFlag number := 0;

  v_OldP3 number;

begin

  -- Check parameters
  if p_CustId is null or p_CitemCode is null or p_Period is null or 
    p_RealStartDate is null then
    p_RetMsg := '參數錯誤';
    return -1;
  end if;

  p_RetMsg := '';
  p_ShouldAmt := 0;

  -- 取收費參數檔相關參數: Para3: 次期起使日同本期截止日, Para10: 使用費率表, Para15: 破月依據
  begin
    select Para3, Para10, Para15 into v_Para3, v_Para10, v_Para15 from SO043;
    v_OldP3 := v_Para3;
    if v_Para3=1 then
      v_Para3 := 0;
    else
      v_Para3 := 1;
    end if;
  exception
    when others then
      p_RetMsg := '收費參數檔內容有誤';
      return -2;
  end;

  -- 從週期性項目設定檔找相關資料, 以決定傳回之初值
  begin
    select Amount, nvl(Period,1) into v_Amount, v_Period from SO003 
      where CustId=p_CustId and CitemCode=p_CitemCode;
  exception
    when others then
      v_FindFlag := 0;
  end;
  v_StartDate := to_date(p_RealStartDate, 'YYYYMMDD');
  v_StopDate  := add_months(v_StartDate, p_Period);

  --v_StopDate  := add_months(v_StartDate, p_Period) - v_Para3;
  --p_RealStopDate := to_char(v_StopDate, 'YYYYMMDD');
  --p_ShouldAmt := trunc(nvl(v_Amount,0)*p_Period/v_Period);

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
	    from CD019CD004 where ClassCode is null and CitemCode=p_CitemCode and Period=p_Period;
	  v_FindFlag := 1;
        exception	-- 也無非特定客戶類別費率表, 則沿用原收費項目代碼中預設值
	  when others then
	    v_FindFlag := 0;
	end;
    end;
  end if;

  if v_FindFlag = 1 then
    if p_Option = 1 then
      v_PFlag := v_PFlag1;
    else
      v_PFlag := v_PFlag2;
    end if;

    if v_PFlag=0 or to_char(v_StartDate, 'DD')='01' then  -- 等於期數
      v_StopDate  := add_months(v_StartDate, p_Period) - v_Para3;
      p_RealStopDate := to_char(v_StopDate, 'YYYYMMDD');
      P_ShouldAmt := nvl(v_Amount,0);

    elsif v_PFlag=1 then	-- 小於期數
      v_Tmp := to_number(to_char(v_StopDate, 'DD'));
      v_StopDate := v_StopDate - v_Tmp + v_OldP3;
      p_RealStopDate := to_char(v_StopDate, 'YYYYMMDD');

      if v_Para15=1 then 	-- 依截止日, 故減去截止日數之金額
	v_Amount := v_Amount - nvl(v_DayAmt,0)*v_Tmp;
      else			-- 依起始日, 故計算該月日數之金額
	v_Tmp := last_day(v_StartDate) - v_StartDate + 1;
	v_Amount := nvl(v_DayAmt,0)*v_Tmp;
      end if; 

      p_ShouldAmt := v_Amount;
      if nvl(v_TruncAmt,0) > 0 then	-- 金額歸整
        p_ShouldAmt := v_Amount - mod(v_Amount, v_TruncAmt);
      end if; 

    else			-- 大於期數
      v_TmpDate := last_day(v_StopDate) + v_OldP3;
      p_RealStopDate := to_char(v_TmpDate, 'YYYYMMDD');

      if v_Para15=1 then 	-- 依截止日, 故加上截止日數之金額
	v_Tmp := last_day(v_StopDate) - v_StopDate + 1;
	v_Amount := v_Amount + nvl(v_DayAmt,0)*v_Tmp;
      else			-- 依起始日, 故加上起始日至月底之金額
	v_Tmp := last_day(v_StartDate) - v_StartDate + 1;
	v_Amount := v_Amount + nvl(v_DayAmt,0)*v_Tmp;
      end if; 

      p_ShouldAmt := v_Amount;
      if nvl(v_TruncAmt,0) > 0 then	-- 金額歸整
        p_ShouldAmt := v_Amount - mod(v_Amount, v_TruncAmt);
      end if;

    end if;
  else
    v_StopDate  := add_months(v_StartDate, p_Period) - v_Para3;
    p_RealStopDate := to_char(v_StopDate, 'YYYYMMDD');

  end if;

  if v_FindFlag = 0 then
    p_RetMsg := '無該期數之費率';
    return 0;
  end if;

  return 0;

exception
  when others then
    -- DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    p_RetMsg := '其他錯誤';
    return -99;
end;
/