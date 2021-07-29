 
create or replace function SF_NewWip2
  (p_CustId in number) return number
  as
  v_NewWipCode2 number := 0;
  v_NewWipName2 varchar2(12);
  v_Count number;
 
begin
  -- Check parameters
  if p_CustId is null then
    return -1;
  end if;
 
  -- 計算該客戶於維修未完工log(SO072)檔筆數
  -- 若有資料, 將該客戶的派工狀態二改為21('in TC ')
  -- 若無資料, 則將派工狀態二改為0('no')
  select count(*) into v_Count from SO072 where CustId=p_CustId;
  if v_Count = 0 then
    v_NewWipCode2 := 0;
    begin	 -- 取名稱
      select Description into v_NewWipName2 from CD036 where CodeNo=v_NewWipCode2;
    exception
      when others then
	return -2;
    end;
  else
    v_NewWipCode2 := 21;
    begin	 -- 取名稱
      select Description into v_NewWipName2 from CD036 where CodeNo=v_NewWipCode2;
    exception
      when others then
	return -3;
    end;
  end if;
 
  update SO001 set WipCode2=v_NewWipCode2, WipName2=v_NewWipName2 where CustId=p_CustId;
 
  commit;
  return 0;
 
exception
  when others then
    rollback;
    return -99;
end;
/
