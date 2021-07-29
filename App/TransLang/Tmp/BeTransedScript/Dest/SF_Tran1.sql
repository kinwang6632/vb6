 
create or replace function SF_Tran1(p_TranOption number, p_StopDate varchar2, p_A1Code varchar2, 
  p_A1Name varchar2, p_A2Code varchar2, p_A2Name varchar2, p_CompCode varchar2, 
  p_PrtOption number, p_UserId varchar2, p_RetMsg out varchar2) return number
  AS
  v_Tmp number;
  v_Para3 number;
  v_Para5 number;
  v_Str1 varchar2(16);
  v_BCCode number;
  v_NextBCDate date;
 
  v_CitemCode number := 0;
  v_TaxCode number;
  v_TaxRate number := 0;
  v_Sign char(1);
  v_CurrSign char(1);
  v_Note  varchar2(100);
  v_Note2 varchar2(100);
  v_AccIdP1 varchar2(8);
  v_AccIdM1 varchar2(8);
  v_AccIdP2 varchar2(8);
  v_AccIdM2 varchar2(8);
 
  v_TaxAmt number;
  v_PureAmt number;
  v_CurrAmt number;
  v_NextAmt number;
 
  v_InstTime date;
  v_AccNo varchar2(8);
  v_TotalTax number := 0;
  v_TotalAmt number := 0;
  v_SumAmt number := 0;
  v_RcdCount number := 0;
  v_OldAmt number := 0;
  v_Now date;
 
  cursor cc1(v_STCode number) is 
    select * from SO033 where RealDate<=to_date(p_StopDate,'YYYYMMDD')
      and UCCode is null and CompCode=p_CompCode and 
      (STCode is null or STCode!=v_STCode)
      order by CitemCode;
 
 -- select ROWID, CitemCode, CustId, CancelFlag
begin
 
  -- 參數檢查
  if p_TranOption is null or (p_TranOption not in (0, 1)) or 
    p_StopDate is null or p_A1Code is null or p_A1Name is null or 
    p_A2Code is null or p_A2Name is null or p_CompCode is null or 
    p_PrtOption is null or p_UserId is null then
    p_RetMsg := 'Wrong parameter.';
    return -1;
  end if;
 
  -- 取短收原因代碼檔中(CD016)參考號為1(呆帳)之呆帳代碼/名稱
  begin
    select CodeNo into v_Tmp from CD016 where RefNo=1;
  exception 
    when others then
      p_RetMsg := '短收原因代碼檔中未設參考號為1(呆帳)的資料';
      return -2;
  end;
 
  -- 取次期起始日同本期截止日, 延後幾日為次收日
  begin
    select Para3, Para5 into v_Para3, v_Para5 from SO043;
  exception 
    when others then
      p_RetMsg := 'No data or too many data in SO043.';
      return -3;
  end;
 
  -- 取基本台訊號費代碼
  begin
    select CodeNo into v_BCCode from CD019 where RefNo=1;
  exception 
    when others then
      p_RetMsg := 'No RefNo=1 record in CD019.                    ';
      return -4;
  end;
 
  -- 檢查稅率是否設定正確
  select count(*) into v_RcdCount from CD019
	where TaxCode not in (select CodeNo from CD033);
  if v_RcdCount > 0 then
      p_RetMsg := 'No tax code in CD033.                                ';
      return -5;
  end if;
 
  DELETE FROM SO079;
 
  v_Str1 := 'End date:'||ltrim(to_char(to_number(substr(p_StopDate,1,4))-1911, '09'))||substr(p_StopDate,5);
  v_Now := sysdate;
  v_RcdCount := 0;
 
  -- Loop SO033 符合條件之每一筆, 進行以下動作
  for cr1 in cc1(v_Tmp) loop
    if cr1.CancelFlag = 1 then  -- 作廢資料搬移至壞帳資料檔SO035
      if p_TranOption=1 then
        insert into SO035 (AddrNo, AreaCode, BillNo, CancelFlag, CitemCode, 
	  CitemName, ClassCode, ClctAreaCode, ClctEn, ClctName, ClctYM, 
	  ClsEn, ClsTime, CMCode, CMName, CompCode, CreateEn, 
	  CreateTime, CustCount, CustId, EntryEn, GUINo, InvoiceTime, 
	  Item, ManualNo, MduId, Note, OldAmt, OldClctEn, OldClctName, 
	  OldPeriod, OldStartDate, OldStopDate, PrintEn, 
	  PrintTime, PrtClsTime, PrtCount, PrtSNo, PTCode, PTName, 
	  Quantity, RealAmt, RealDate, RealPeriod, RealStartDate, 
	  RealStopDate, ServCode, ShouldAmt, ShouldDate, STCode, STName, 
	  StrtCode, UCCode, UCName, UpdEn, UpdTime)
	  values (cr1.AddrNo, cr1.AreaCode, cr1.BillNo, cr1.CancelFlag, cr1.CitemCode, 
	  cr1.CitemName, cr1.ClassCode, cr1.ClctAreaCode, cr1.ClctEn, cr1.ClctName, cr1.ClctYM, 
	  p_UserId, v_Now, cr1.CMCode, cr1.CMName, cr1.CompCode, cr1.CreateEn, 
	  cr1.CreateTime, cr1.CustCount, cr1.CustId, cr1.EntryEn, cr1.GUINo, cr1.InvoiceTime, 
	  cr1.Item, cr1.ManualNo, cr1.MduId, cr1.Note, cr1.OldAmt, cr1.OldClctEn, 
	  cr1.OldClctName, cr1.OldPeriod, cr1.OldStartDate, cr1.OldStopDate, cr1.PrintEn, 
	  cr1.PrintTime, cr1.PrtClsTime, cr1.PrtCount, cr1.PrtSNo, cr1.PTCode, cr1.PTName, 
	  cr1.Quantity, cr1.RealAmt, cr1.RealDate, cr1.RealPeriod, cr1.RealStartDate, 
	  cr1.RealStopDate, cr1.ServCode, cr1.ShouldAmt, cr1.ShouldDate, cr1.STCode, cr1.STName, 
	  cr1.StrtCode, cr1.UCCode, cr1.UCName, cr1.UpdEn, cr1.UpdTime);
 
        delete from SO033 where BillNo=cr1.BillNo and Item=cr1.Item;
      end if;
 
    else			-- 一般資料
      if cr1.CitemCode != v_CitemCode then
	-- 取該收費項目之當期收入科目,當期預收科目,次期收入科目,次期預收科目
	v_CitemCode := cr1.CitemCode;
        select AccIdP1, AccIdM1, AccIdP2, AccIdM2, TaxCode, Sign
	  into v_AccIdP1, v_AccIdM1, v_AccIdP2, v_AccIdM2, v_TaxCode, v_Sign
	  from CD019 where CodeNo=v_CitemCode;
          if v_Sign='+' then
             v_Sign:='-';
          else
             v_Sign:='+';
          end if;
 
           if v_TaxCode is not null then	-- 取目前稅率
                  select nvl(Rate1,0) into v_TaxRate from CD033 where CodeNo=v_TaxCode;
	  --v_TaxRate := 5;	
           else
	  v_TaxRate := 0;
           end if;
      end if;
 
      v_CurrSign := '-';
      v_Note2 := cr1.BillNo || '-' || cr1.CitemName || ',' || GiPackage.ED2CDString(cr1.RealDate);
      v_TaxAmt := 0;
      v_PureAmt := 0;
      v_CurrAmt := 0;
      v_NextAmt := 0;
 
      if cr1.RealPeriod > 0 then	-- 週期性費用
	-- 計算期初攤提
	v_Tmp := SF_NGet1(cr1.ShouldAmt, cr1.RealDate, cr1.RealStartDate, cr1.RealStopDate, 
			v_TaxRate, v_Para3, v_TaxAmt, v_PureAmt, v_CurrAmt , v_NextAmt);
	if v_Sign = '-' then 
 	  v_CurrSign := '+';
	  v_PureAmt := 0 + v_PureAmt;
	  v_CurrAmt := 0 + v_CurrAmt;
	  v_NextAmt := 0 + v_NextAmt;
                   else                                                   -- Add 2000.08.09
 	  v_CurrSign := '-';
	  v_PureAmt := 0 - v_PureAmt;
	  v_CurrAmt := 0 - v_CurrAmt;
	  v_NextAmt := 0 - v_NextAmt;
	end if;
 
	-- 取該客戶之安裝時間, 當本筆為基本台費時, 再取基本台次收日
	if v_CitemCode = v_BCCode then
	  select InstTime, NextBCDate into v_InstTime, v_NextBCDate 
		from SO001 where CustId=cr1.CustId;
	  -- 更新該客戶之基本台次收日
	  if (cr1.RealStopDate+v_Para5 > v_NextBCDate) and p_TranOption=1 then 
	    update SO001 set NextBCDate=cr1.RealStopDate+v_Para5 where CustId=cr1.CustId;
	  end if;
	else
	  select InstTime into v_InstTime from SO001 where CustId=cr1.CustId;
	end if;
 
	if v_InstTime>=cr1.RealStartDate and v_InstTime<cr1.RealStopDate+1 then	-- 首次費
	  v_AccNo := v_AccIdM1;		-- 當期預收科目
	  v_Note := v_Str1;
	  if p_PrtOption = 1 then
	    v_Note := v_Note || ',' || v_Note2;
	  end if; 
	  if v_CurrAmt != 0 then	-- 新增一筆至暫存檔
                        if v_OldAmt != cr1.RealAmt then
                              insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt, 
		RealPeriod, CitemCode, Sign, Note)
		values (v_AccNo, cr1.CustId, v_CurrAmt, cr1.RealDate, cr1.RealAmt,
		cr1.RealPeriod, cr1.CitemCode, v_Sign, v_Note);
                              v_OldAmt := cr1.RealAmt;
                        else
                              insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt, 
		RealPeriod, CitemCode, Sign, Note)
		values (v_AccNo, cr1.CustId, v_CurrAmt, cr1.RealDate, 0,
		cr1.RealPeriod, cr1.CitemCode, v_Sign, v_Note);
                        end if; 
                     end if;
	  v_AccNo := v_AccIdP1;		-- 當期收入科目
 
	else			-- 非首次費
	  v_AccNo := v_AccIdM2;		-- 次期預收科目
	  v_Note := v_Str1;
	  if p_PrtOption = 1 then
	    v_Note := v_Note || ',' || v_Note2;
	  end if; 
	  if v_CurrAmt != 0 then	-- 新增一筆至暫存檔
                        if v_OldAmt != cr1.RealAmt then
                             insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt, 
		RealPeriod, CitemCode, Sign, Note)
		values (v_AccNo, cr1.CustId, v_CurrAmt, cr1.RealDate, cr1.RealAmt,
		cr1.RealPeriod, cr1.CitemCode, v_Sign, v_Note);
                             v_OldAmt := cr1.RealAmt;
                         else
                             insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt, 
		RealPeriod, CitemCode, Sign, Note)
		values (v_AccNo, cr1.CustId, v_CurrAmt, cr1.RealDate, 0,
		cr1.RealPeriod, cr1.CitemCode, v_Sign, v_Note);
                         end if;
	  end if; 
	  v_AccNo := v_AccIdP2;		-- 次期收入科目
	end if;
 
      else				-- 非週期性費用
	if v_TaxRate > 0 then		-- 計算銷售額
	  v_PureAmt := round(cr1.RealAmt/(1+v_TaxRate/100));
	else
	  v_PureAmt := cr1.RealAmt;
	end if;
	v_TaxAmt := cr1.RealAmt - v_PureAmt;	-- 計算稅額
	v_AccNo := v_AccIdP1;			-- 當期收入科目
	v_Note := v_Str1;
	if p_PrtOption = 1 then
	  v_Note := v_Note || ',' || v_Note2;
	end if; 
	v_NextAmt := v_PureAmt;
 
      end if;
    end if;
 
    if cr1.CancelFlag != 1 then
       if v_NextAmt != 0 and v_AccNo is not null then	-- 新增一筆至暫存檔
          if v_OldAmt != cr1.RealAmt then
              insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt, 
	RealPeriod, CitemCode, Sign, Note)
	values (v_AccNo, cr1.CustId, v_NextAmt, cr1.RealDate, cr1.RealAmt,
	0, cr1.CitemCode, v_Sign, v_Note);
             v_OldAmt := cr1.RealAmt; 
          else
              insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt, 
	RealPeriod, CitemCode, Sign, Note)
	values (v_AccNo, cr1.CustId, v_NextAmt, cr1.RealDate, 0,
	0, cr1.CitemCode, v_Sign, v_Note);
          end if;
       end if;
 
       if v_Sign = '-' then 
         v_TotalTax := v_TotalTax + v_TaxAmt;
         v_TotalAmt := v_TotalAmt + v_PureAmt;
       else
         v_TotalTax := v_TotalTax - v_TaxAmt;
         v_TotalAmt := v_TotalAmt - v_PureAmt;
       end if;
       insert into SO080 (Rcdno, Custid) values (v_TaxAmt, v_PureAmt);
   end if;
 
    -- 搬移至已日結資料檔SO034
    if p_TranOption=1 and cr1.CancelFlag != 1 then
      begin
        insert into SO034 (AddrNo, AreaCode, BillNo, CancelFlag, CitemCode, 
	  CitemName, ClassCode, ClctAreaCode, ClctEn, ClctName, ClctYM, 
	  ClsEn, ClsTime, CMCode, CMName, CompCode, CreateEn, 
	  CreateTime, CustCount, CustId, EntryEn, GUINo, InvoiceTime, 
	  Item, ManualNo, MduId, Note, OldAmt, OldClctEn, OldClctName, 
	  OldPeriod, OldStartDate, OldStopDate, PrintEn, 
	  PrintTime, PrtClsTime, PrtCount, PrtSNo, PTCode, PTName, 
	  Quantity, RealAmt, RealDate, RealPeriod, RealStartDate, 
	  RealStopDate, ServCode, ShouldAmt, ShouldDate, STCode, STName, 
	  StrtCode, UCCode, UCName, UpdEn, UpdTime, ActYM)
	  values (cr1.AddrNo, cr1.AreaCode, cr1.BillNo, cr1.CancelFlag, cr1.CitemCode, 
	  cr1.CitemName, cr1.ClassCode, cr1.ClctAreaCode, cr1.ClctEn, cr1.ClctName, cr1.ClctYM, 
	  cr1.ClsEn, cr1.ClsTime, cr1.CMCode, cr1.CMName, cr1.CompCode, cr1.CreateEn, 
	  cr1.CreateTime, cr1.CustCount, cr1.CustId, cr1.EntryEn, cr1.GUINo, cr1.InvoiceTime, 
	  cr1.Item, cr1.ManualNo, cr1.MduId, cr1.Note, cr1.OldAmt, cr1.OldClctEn, 
	  cr1.OldClctName, cr1.OldPeriod, cr1.OldStartDate, cr1.OldStopDate, cr1.PrintEn, 
	  cr1.PrintTime, cr1.PrtClsTime, cr1.PrtCount, cr1.PrtSNo, cr1.PTCode, cr1.PTName, 
	  cr1.Quantity, cr1.RealAmt, cr1.RealDate, cr1.RealPeriod, cr1.RealStartDate, 
	  cr1.RealStopDate, cr1.ServCode, cr1.ShouldAmt, cr1.ShouldDate, cr1.STCode, cr1.STName, 
	  cr1.StrtCode, cr1.UCCode, cr1.UCName, cr1.UpdEn, cr1.UpdTime, 
	  to_number(substr(p_StopDate,1,6)));
      exception
        when others then
	  dbms_output.put_line(cr1.CustId || ' ' || cr1.BillNo || ' ' || cr1.CitemCode || ' ' || cr1.CitemName || ' ' || to_char(cr1.ShouldDate, 'YYYYMMDD'));
	  DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
      end; 
      delete from SO033 where BillNo=cr1.BillNo and Item=cr1.Item;
    end if;
 
    if cr1.CancelFlag != 1 then
       v_RcdCount := v_RcdCount + 1;
    end if;
 
  end loop;
 
  v_SumAmt:=(v_TotalAmt+v_TotalTax);
  if v_SumAmt != 0 then		-- 新增一筆至暫存檔
    insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt, 
	RealPeriod, CitemCode, Sign, Note)
	values (p_A1Code, 0, v_SumAmt, to_date(p_StopDate,'YYYYMMDD'), v_SumAmt,
	0, null, '+', v_Str1);
  end if;
 
  if v_TotalTax != 0 then		-- 新增一筆至暫存檔
    insert into SO079 (AccNo, CustId, ShouldAmt, RealDate, RealAmt, 
	RealPeriod, CitemCode, Sign, Note)
	values (p_A2Code, 0,v_TotalTax, to_date(p_StopDate,'YYYYMMDD'), 0,
	0, null, '-', v_Str1);
  end if;
 
  p_RetMsg := 'Execution OK, Rcd count=' || v_RcdCount;
--  commit;
  return v_RcdCount;
 
exception
  when others then
    rollback;
    DBMS_OUTPUT.PUT_LINE('Error Code ='||SQLCODE||' '||'Error message='||SQLERRM);
    p_RetMsg := 'Other error.';
    return -99;
end;
/
