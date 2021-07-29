 
create or replace function SF_GetOverDue1
  (p_CustId number, p_Para4 number) return number as
 
  v_ShouldAmt number;
  v_Today date;
 
begin
  -- Check parameters
  if p_CustId is null or p_Para4 is null then
    return 0;
  end if;
 
  v_ShouldAmt := 0;
  v_Today := to_date(to_char(sysdate, 'YYYYMMDD'), 'YYYYMMDD');	-- 取今日
 
  -- 取欠費總合
  select sum(ShouldAmt) into v_ShouldAmt from SO033 
    where CustId=p_CustId and UCCode is not null and 
    (substr(BillNo, 7, 1)='B' and ShouldDate<v_Today-p_Para4 or 
     substr(BillNo, 7, 1)!='B' and ShouldDate<v_Today );
 
  return 0-nvl(v_ShouldAmt,0);
 
exception
  when others then
    -- DBMS_OUTPUT.PUT_LINE('SQLCODE='||SQLCODE||' '||'SQLERRM='||SQLERRM);
    return 0;
end;
/
