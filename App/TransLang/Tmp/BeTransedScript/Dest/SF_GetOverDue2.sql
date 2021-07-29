 
create or replace function SF_GetOverDue2
  (p_CustId number, p_Para4 number, p_RetMsg out varchar2) return number as
 
  v_ShouldAmt number;
  v_Today date;
  v_ShouldDate date;
 
begin
  -- Check parameters
  if p_CustId is null or p_Para4 is null then
    return 0;
  end if;
 
  p_RetMsg := null;
  v_ShouldAmt := 0;
  v_Today := to_date(to_char(sysdate, 'YYYYMMDD'), 'YYYYMMDD');	-- 取今日
 
  -- 取欠費總合, 與最早之應收日期
  select sum(ShouldAmt), min(ShouldDate) into v_ShouldAmt, v_ShouldDate from SO033 
    where CustId=p_CustId and UCCode is not null and 
    (substr(BillNo, 7, 1)='B' and ShouldDate<v_Today-p_Para4 or 
     substr(BillNo, 7, 1)!='B' and ShouldDate<v_Today );
 
  if v_ShouldDate is not null then
    p_RetMsg := 'Non-pay cust. Charged date:' || GiPackage.ED2CDString(v_ShouldDate);
  else
    p_RetMsg := 'Active';
  end if;
 
  return 0-nvl(v_ShouldAmt,0);
 
exception
  when others then
    -- DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    return 0;
end;
/
