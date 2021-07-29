 
create or replace function SF_NewWip1
  (p_CustId in number) return number
  as
  v_NewWipCode1 number := 0;
  v_NewWipName1 varchar2(12);
  v_RefNo number;
 
  -- 該客戶於裝機未完工log(SO071)檔中,最近的一筆log資料
  cursor cc1 is
    select RefNo from SO071 where CustId=p_CustId order by AcceptTime desc;
 
begin
  -- Check parameters
  if p_CustId is null then
    return -1;
  end if;
 
  -- 取派工狀態類別代碼檔(CD036)代碼=0的名稱
  begin
    select Description into v_NewWipName1 from CD036 where CodeNo=0;
  exception 
    when others then
      return -2;
  end;
 
  -- 取該客戶於裝機未完工log(SO071)檔中,最近的一筆log資料的裝機類別代碼的參考號
  -- 若有資料, 判斷該參考號, 來決定將該客戶的派工狀態一改為何值
  -- 若無資料, 則將派工狀態一改為0
  v_RefNo := null;
  open cc1;
    fetch cc1 into v_RefNo;
  close cc1;
 
  if v_RefNo is null then
    v_NewWipCode1 := 0;		-- 無
  elsif v_RefNo=0 then
    v_NewWipCode1 := 5;		-- 其他(室內移機,換線)
  elsif v_RefNo=1 then
    v_NewWipCode1 := 1;		-- 裝機中
  elsif v_RefNo=5 or v_RefNo=7 then
    v_NewWipCode1 := 2;		-- 復機中
  elsif v_RefNo=6 then
    v_NewWipCode1 := 4;		-- 改裝
  else
    v_NewWipCode1 := 3;		-- 設備加裝
  end if;
  if v_RefNo is not null then
    select Description into v_NewWipName1 from CD036 where CodeNo=v_NewWipCode1; -- 取名稱
  end if;
  update SO001 set WipCode1=v_NewWipCode1, WipName1=v_NewWipName1 where CustId=p_CustId;
 
  commit;
  return 0;
 
exception
  when others then
    rollback;
    return -99;
end;
/
