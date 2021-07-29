 
create or replace function SF_NGet1(p_RealAmt number, p_RealDate date, p_StartDate date, 
      p_StopDate date, p_TaxRate number, p_SameDay number, p_TaxAmt out number, 
      p_PureAmt out number, p_CurrAmt out number, p_NextAmt out number) return number
  AS
  v_FirstDay date;
  v_LastDay date;
  v_Amount number;
  v_StopDate date;
  v_SDate date;
  v_EDate date;
  v_AvgAmt number;
  v_ThisAmt number;
  v_NowAmt number := 0;
 
begin
/*
  -- 參數檢查
  if p_DueDate is null or p_UpdTime is null or p_UpdName is null then
    --p_RetMsg := 'Wrong parameter.';
    return -1;
  end if;
*/
 
  v_FirstDay := round(p_RealDate, 'MONTH');	-- 該月首日
  v_LastDay := last_day(p_RealDate);		-- 該月尾日
 
  -- 計算稅額
  if nvl(p_TaxRate,0) > 0 then
    v_Amount := round(p_RealAmt*100/(100+p_TaxRate));
  else
    v_Amount := p_RealAmt;
  end if;
  p_TaxAmt := p_RealAmt - v_Amount;
  p_PureAmt := v_Amount;
  --p_CurrAmt := 0;
  --p_NextAmt := 0;
 
  v_StopDate := p_StopDate - (1-p_SameDay);
  v_SDate := p_StartDate;
  v_EDate := least(last_day(v_SDate), v_StopDate);
 
  while v_EDate<=v_StopDate and v_SDate<=v_StopDate and v_EDate<=v_LastDay loop
--    v_AvgAmt := round(v_Amount/(v_StopDate-v_SDate+1),6);  -- 取日平均額
--    v_ThisAmt := round(v_AvgAmt * (v_EDate-v_SDate+1));    -- 該月份金額
    v_AvgAmt := round(v_Amount/(v_StopDate-v_SDate+(1-p_SameDay)),6);  -- 取日平均額
    v_ThisAmt := round(v_AvgAmt * (v_EDate-v_SDate+(1-p_SameDay)),6);   -- 該月份金額
    if v_EDate <= v_LastDay then
      v_NowAmt := v_NowAmt + v_ThisAmt;
    end if;
    v_Amount := v_Amount - v_ThisAmt;
    v_SDate := v_EDate + 1;
    v_EDate := least(last_day(v_SDate), v_StopDate);
  end loop; 
 
  p_NextAmt := v_Amount;
  p_CurrAmt := v_NowAmt;
 
  return 0;
 
exception
  when others then
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    --p_RetMsg := 'Other error.';
    return -99;
end;
/
